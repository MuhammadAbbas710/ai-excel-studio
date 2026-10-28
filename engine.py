"""AI Excel Studio engine.

Contains (in one file, to keep deployment to 4 files):
  1. Spec validation / normalisation (the AI only ever produces a JSON "spec";
     it never produces code, and nothing from the AI is executed)
  2. Gemini calls (image analysis, planning, editing, transcription)
  3. Workbook builder (openpyxl, native charts, validation, conditional formatting)
  4. Safe editing operations on the spec
  5. Dataset import, quality checks, Python-side preview, PDF guide
"""
from __future__ import annotations

import ast
import copy
import io
import json
import operator
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from datetime import date, datetime
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.chart import (AreaChart, BarChart, DoughnutChart, LineChart,
                            PieChart, Reference, ScatterChart, Series)
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.formatting.rule import (CellIsRule, ColorScaleRule, DataBarRule,
                                      FormulaRule)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties

# ----------------------------------------------------------------- constants
HEADER_ROW = 4
FIRST_ROW = 5
COL_TYPES = ("text", "number", "currency", "percent", "date", "dropdown", "formula")
NUMERIC_TYPES = ("number", "currency", "percent", "formula")
CHART_TYPES = ("column", "bar", "line", "area", "pie", "doughnut", "scatter")
AGGS = ("sum", "average", "count", "max", "min")
FORMATS = ("number", "integer", "currency", "percent", "text", "date")
CF_RULES = ("data_bar", "color_scale", "progress", "greater_than", "less_than",
            "equal", "duplicates", "blank", "overdue")
FONTS = ("Calibri", "Arial", "Segoe UI", "Georgia", "Cambria", "Verdana",
         "Tahoma", "Trebuchet MS", "Garamond", "Century Gothic", "Times New Roman")
MAX_ROWS = 5000
INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\\.]")
BLOCKED_FUNCS = ("WEBSERVICE", "CALL", "REGISTER.ID", "EXEC", "DDE", "HYPERLINK",
                 "FILTERXML", "RTD", "SHELL", "SQL.REQUEST", "MSEXCEL")


class SpecError(Exception):
    """The workbook specification (or an edit) is invalid."""


class AIError(Exception):
    """The AI service failed or returned something unusable."""


# -------------------------------------------------------------------- themes
THEMES = {
    "Minimalist professional": dict(primary="1F2937", accent="6B7280", background="FFFFFF", surface="F3F4F6", text="111827"),
    "Modern corporate": dict(primary="1E3A8A", accent="3B82F6", background="FFFFFF", surface="EFF6FF", text="0F172A"),
    "Soft pastel": dict(primary="7E6BA8", accent="F4A6C0", background="FFFFFF", surface="FDF2F8", text="3B3355"),
    "Sage green": dict(primary="5F7A61", accent="A3B899", background="FAF7F0", surface="EEF2E9", text="2F3A2F"),
    "Neutral beige": dict(primary="8B7355", accent="C9B79C", background="FFFBF5", surface="F5EFE6", text="3E3328"),
    "Dark professional": dict(primary="38BDF8", accent="818CF8", background="111827", surface="1F2937", text="F9FAFB"),
    "Blue analytics": dict(primary="0B5394", accent="1C9CF6", background="FFFFFF", surface="E8F1FA", text="0B2540"),
    "Finance-focused": dict(primary="0B6E4F", accent="C9A227", background="FFFFFF", surface="E9F5EF", text="0A2F23"),
    "Wellness and lifestyle": dict(primary="2A9D8F", accent="E9C46A", background="FFFFFF", surface="E8F6F3", text="1F3D39"),
}


def _hex(c, default):
    c = str(c or "").strip().lstrip("#")
    return c.upper() if re.fullmatch(r"[0-9A-Fa-f]{6}", c) else default


def _lum(h):
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255


def _mix(a, b, t):
    ca = [int(a[i:i + 2], 16) for i in (0, 2, 4)]
    cb = [int(b[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(ca, cb))


def norm_theme(t) -> dict:
    """Return a theme dict with '#'-prefixed colours (used by the UI and stored in the spec)."""
    if isinstance(t, str):
        t = {"name": t}
    t = t if isinstance(t, dict) else {}
    name = t.get("name") if t.get("name") in THEMES else None
    base = THEMES[name or "Modern corporate"]
    out = {k: "#" + _hex(t.get(k), v) for k, v in base.items()}
    if name and any(out[k][1:] != base[k] for k in base):
        name = "Custom"
    elif not name:
        name = "Custom" if any(k in t for k in base) else "Modern corporate"
    out["name"] = name
    out["font"] = t.get("font") if t.get("font") in FONTS else "Calibri"
    return out


def _T(spec) -> dict:
    """Theme as plain hex (no '#') plus derived colours, for the builder."""
    t = {k: (v[1:] if isinstance(v, str) and v.startswith("#") else v)
         for k, v in norm_theme(spec.get("theme")).items()}
    t["on_primary"] = "FFFFFF" if _lum(t["primary"]) < 0.6 else "111111"
    t["dark"] = _lum(t["background"]) < 0.4
    t["muted"] = _mix(t["text"], t["background"], 0.45)
    t["border"] = _mix(t["text"], t["background"], 0.82)
    return t


def _palette(T):
    p, a, bg = T["primary"], T["accent"], T["background"]
    return [p, a, _mix(p, bg, 0.4), _mix(a, "000000", 0.3), _mix(p, a, 0.5),
            _mix(a, bg, 0.5), _mix(p, "000000", 0.35), _mix(p, bg, 0.65)]


# ------------------------------------------------------------ formula safety
def _check_formula_text(f: str):
    if len(f) > 1500:
        raise SpecError("Formula is too long")
    if not re.fullmatch(r"[\w\s\.\+\-\*/\^\(\),:;\"'<>=&%$!#]*", f):
        raise SpecError(f"Formula contains unsupported characters: {f[:60]}")
    if f.count("(") != f.count(")"):
        raise SpecError(f"Unbalanced parentheses in formula: {f[:60]}")
    if f.count('"') % 2:
        raise SpecError(f"Unbalanced quotes in formula: {f[:60]}")
    for fn in BLOCKED_FUNCS:
        if re.search(r"\b" + re.escape(fn) + r"\s*\(", f, re.I):
            raise SpecError(f"Function {fn} is not allowed")
    if re.search(r"\[|\]|\||\\|file:|https?:", f, re.I):
        raise SpecError("External references are not allowed in formulas")


def resolve_formula(expr, letters, r, first_row, ranges=None) -> str:
    """Turn the friendly column formula into an Excel formula (without '=').

    [Column]       -> the same row's cell in that column
    [Column:prev]  -> the previous row's cell (0 on the first row)
    {Sheet.Column} -> the whole data range of a column on a table sheet
    """
    expr = str(expr or "").strip()
    if expr.startswith("="):
        expr = expr[1:]
    if not expr:
        raise SpecError("Empty formula")

    def rng(m):
        sheet, col = m.group(1).strip(), m.group(2).strip()
        if not ranges or sheet not in ranges or col not in ranges[sheet]:
            raise SpecError(f"Unknown range {{{sheet}.{col}}}")
        return ranges[sheet][col]

    expr = re.sub(r"\{([^{}.]+)\.([^{}]+)\}", rng, expr)

    def ref(m):
        name, prev = m.group(1), False
        if name.endswith(":prev"):
            name, prev = name[:-5], True
        if name not in letters:
            raise SpecError(f'Unknown column "{name}"')
        if prev:
            return "0" if r == first_row else f"{letters[name]}{r - 1}"
        return f"{letters[name]}{r}"

    out = re.sub(r"\[([^\[\]]+)\]", ref, expr)
    _check_formula_text(out)
    return out


def _q(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


# ------------------------------------------------------------- value helpers
def _key(v) -> str:
    if isinstance(v, (datetime, date)):
        return v.date().isoformat() if isinstance(v, datetime) else v.isoformat()
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip().lower()


_DATE_FORMATS = ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%Y/%m/%d",
                 "%b %d, %Y", "%d %b %Y", "%Y-%m-%d %H:%M:%S")


def _coerce(v, ctype):
    if v is None or (isinstance(v, float) and v != v):
        return None
    if isinstance(v, str) and not v.strip():
        return None
    if ctype in ("number", "currency", "percent"):
        if isinstance(v, bool):
            return v
        if isinstance(v, (int, float)):
            return int(v) if float(v).is_integer() else float(v)
        if isinstance(v, str):
            s = v.strip().replace(",", "")
            pct = s.endswith("%")
            s = s.strip("$€£%").strip()
            try:
                x = float(s) / (100 if pct else 1)
                return int(x) if x.is_integer() else x
            except ValueError:
                return v
        return v
    if ctype == "date":
        if isinstance(v, datetime):
            return v
        if isinstance(v, date):
            return datetime(v.year, v.month, v.day)
        if isinstance(v, str):
            for f in _DATE_FORMATS:
                try:
                    return datetime.strptime(v.strip(), f)
                except ValueError:
                    pass
        return v
    if ctype == "formula":
        return None
    if isinstance(v, (datetime, date)):
        return _key(v)
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _where_match(x, w) -> bool:
    op = w.get("op", "=")
    val = w.get("equals", w.get("value"))
    if x is None:
        return False
    if op == "=":
        return _key(x) == _key(val)
    if op == "<>":
        return _key(x) != _key(val)
    try:
        a, b = float(x), float(val)
    except (TypeError, ValueError):
        return False
    return {">": a > b, "<": a < b, ">=": a >= b, "<=": a <= b}.get(op, False)


# ----------------------------------------------------- spec normalisation
def _clean_sheet_name(name, used):
    n = INVALID_SHEET_CHARS.sub(" ", str(name or "Sheet")).strip().strip("'") or "Sheet"
    n = re.sub(r"\s+", " ", n)[:31]
    base, i = n, 2
    while n.lower() in used:
        suf = f" {i}"
        n = base[:31 - len(suf)] + suf
        i += 1
    used.add(n.lower())
    return n


def _norm_columns(cols, warnings, sname):
    out, used = [], set()
    for c in cols or []:
        if isinstance(c, str):
            c = {"name": c}
        if not isinstance(c, dict):
            continue
        name = re.sub(r"[\[\]{}]", "", str(c.get("name", "")).strip())[:60]
        if not name:
            continue
        base, i = name, 2
        while name.lower() in used:
            name, i = f"{base} {i}", i + 1
        used.add(name.lower())
        t = c.get("type", "text")
        t = t if t in COL_TYPES else "text"
        col = {"name": name, "type": t}
        if t == "dropdown":
            opts = [str(o) for o in (c.get("options") or []) if str(o).strip()]
            if opts:
                col["options"] = opts[:200]
            else:
                col["type"] = "text"
        if t == "formula":
            col["formula"] = str(c.get("formula", ""))
            col["format"] = c.get("format") if c.get("format") in FORMATS else "number"
        for k in ("min", "max"):
            if isinstance(c.get(k), (int, float)) and not isinstance(c.get(k), bool):
                col[k] = c[k]
        if isinstance(c.get("width"), (int, float)):
            col["width"] = max(6, min(60, float(c["width"])))
        out.append(col)
    return out


def table_ranges(spec) -> dict:
    """{sheet: {column: "'Sheet'!$C$5:$C$104"}} for every table sheet."""
    res = {}
    for sh in spec["sheets"]:
        if sh["type"] != "table":
            continue
        last = FIRST_ROW + max(len(sh["rows"]) + sh["extra_rows"], 1) - 1
        res[sh["name"]] = {
            c["name"]: f"{_q(sh['name'])}!${get_column_letter(i + 1)}${FIRST_ROW}:${get_column_letter(i + 1)}${last}"
            for i, c in enumerate(sh["columns"])}
    return res


def _formula_refs(expr):
    return [m[:-5] if m.endswith(":prev") else m for m in re.findall(r"\[([^\[\]]+)\]", expr)]


def normalize_spec(spec) -> tuple[dict, list]:
    """Validate and repair a spec. Returns (clean_spec, warnings). Never executes anything."""
    warnings: list[str] = []
    if not isinstance(spec, dict):
        raise SpecError("The workbook specification is not a JSON object.")
    out: dict[str, Any] = {
        "workbook_name": re.sub(r"[^\w\- ]", "", str(spec.get("workbook_name") or "My Workbook")).strip()[:60] or "My Workbook",
        "theme": norm_theme(spec.get("theme")),
        "settings": {"currency_symbol": str((spec.get("settings") or {}).get("currency_symbol", "$"))[:3] or "$",
                     "add_instructions": bool((spec.get("settings") or {}).get("add_instructions", True))},
        "assumptions": [str(a) for a in (spec.get("assumptions") or [])][:30],
        "design_notes": str(spec.get("design_notes") or "")[:1000],
        "sheets": [],
    }
    used = {"lists"}
    raw_sheets = spec.get("sheets") or []
    if not raw_sheets:
        raise SpecError("The specification contains no worksheets.")
    for raw in raw_sheets[:30]:
        if not isinstance(raw, dict):
            continue
        t = raw.get("type", "table")
        t = t if t in ("table", "dashboard", "instructions") else "table"
        sh = {"name": _clean_sheet_name(raw.get("name"), used), "type": t,
              "title": str(raw.get("title") or raw.get("name") or "")[:100],
              "purpose": str(raw.get("purpose") or "")[:400],
              "notes": str(raw.get("notes") or "")[:400],
              "sample_data": bool(raw.get("sample_data", True))}
        if t == "table":
            sh["columns"] = _norm_columns(raw.get("columns"), warnings, sh["name"])
            if not sh["columns"]:
                warnings.append(f'Sheet "{sh["name"]}" had no valid columns; added a default "Item" column.')
                sh["columns"] = [{"name": "Item", "type": "text"}]
            n = len(sh["columns"])
            names = [c["name"] for c in sh["columns"]]
            rows = []
            for r in (raw.get("rows") or [])[:MAX_ROWS]:
                if isinstance(r, dict):
                    r = [r.get(nm) for nm in names]
                if isinstance(r, (list, tuple)):
                    r = list(r)[:n]
                    rows.append(r + [None] * (n - len(r)))
            sh["rows"] = rows
            try:
                ex = int(raw.get("extra_rows", 100))
            except (TypeError, ValueError):
                ex = 100
            sh["extra_rows"] = max(0, min(1000, ex))
            sh["conditional"] = [c for c in (raw.get("conditional") or []) if isinstance(c, dict)][:20]
        elif t == "dashboard":
            sh["kpis"] = [k for k in (raw.get("kpis") or []) if isinstance(k, dict)][:12]
            sh["charts"] = [c for c in (raw.get("charts") or []) if isinstance(c, dict)][:8]
        else:
            sh["lines"] = [str(x) for x in (raw.get("lines") or [])][:60]
        out["sheets"].append(sh)

    tables = {s["name"]: s for s in out["sheets"] if s["type"] == "table"}
    ranges = table_ranges(out)

    # pass 2: formula columns
    for sh in tables.values():
        letters = {c["name"]: get_column_letter(i + 1) for i, c in enumerate(sh["columns"])}
        fcols = {c["name"]: c for c in sh["columns"] if c["type"] == "formula"}
        for c in list(fcols.values()):
            try:
                resolve_formula(c["formula"], letters, FIRST_ROW + 1, FIRST_ROW, ranges)
                if c["name"] in [r for r in _formula_refs(c["formula"])
                                 if f"[{r}:prev]" not in c["formula"]]:
                    raise SpecError("formula refers to its own column")
            except SpecError as e:
                warnings.append(f'"{sh["name"]}" column "{c["name"]}": {e}. Converted to a plain text column.')
                c["type"] = "text"
                c.pop("formula", None)
                c.pop("format", None)
                fcols.pop(c["name"], None)
        # circular dependency check between formula columns
        resolved, changed = set(), True
        while changed:
            changed = False
            for nm, c in fcols.items():
                if nm in resolved:
                    continue
                deps = [d for d in _formula_refs(c["formula"]) if d in fcols
                        and f"[{d}:prev]" not in c["formula"]]
                if all(d in resolved for d in deps):
                    resolved.add(nm)
                    changed = True
        for nm, c in fcols.items():
            if nm not in resolved:
                warnings.append(f'"{sh["name"]}" column "{nm}" is part of a circular formula chain. Converted to text.')
                c["type"] = "text"
                c.pop("formula", None)
                c.pop("format", None)
        # conditional rules
        names = [c["name"] for c in sh["columns"]]
        good = []
        for r in sh["conditional"]:
            if r.get("column") in names and r.get("rule") in CF_RULES:
                good.append({"column": r["column"], "rule": r["rule"], "value": r.get("value"),
                             "color": r.get("color") if r.get("color") in ("red", "green", "amber", "blue") else "red"})
            else:
                warnings.append(f'Ignored an invalid conditional-formatting rule on "{sh["name"]}".')
        sh["conditional"] = good

    # pass 3: dashboards
    for sh in out["sheets"]:
        if sh["type"] != "dashboard":
            continue
        kpis, labels = [], set()
        for k in sh["kpis"]:
            label = str(k.get("label", "")).strip()[:40]
            if not label or label.lower() in labels:
                warnings.append("Skipped a KPI with a missing or duplicate label.")
                continue
            fmt = k.get("format") if k.get("format") in FORMATS else "number"
            if k.get("expr"):
                kpis.append({"label": label, "expr": str(k["expr"]), "format": fmt})
            else:
                s, col = tables.get(k.get("sheet")), k.get("column")
                names = [c["name"] for c in s["columns"]] if s else []
                agg = k.get("agg") if k.get("agg") in AGGS else "sum"
                if not s or (col not in names and agg != "count"):
                    warnings.append(f'KPI "{label}" refers to an unknown sheet/column and was skipped.')
                    continue
                item = {"label": label, "sheet": s["name"], "column": col if col in names else names[0],
                        "agg": agg, "format": fmt}
                w = k.get("where")
                if isinstance(w, dict) and w.get("column") in names:
                    item["where"] = {"column": w["column"], "equals": w.get("equals", w.get("value")),
                                     "op": w.get("op") if w.get("op") in ("=", "<>", ">", "<", ">=", "<=") else "="}
                kpis.append(item)
            labels.add(label.lower())
        # expression KPIs may only reference aggregate KPIs
        agg_labels = {k["label"] for k in kpis if "expr" not in k}
        final = []
        for k in kpis:
            if "expr" in k:
                refs = re.findall(r"\{([^{}]+)\}", k["expr"])
                try:
                    if not refs or any(r not in agg_labels for r in refs):
                        raise SpecError("expression must reference aggregate KPI labels in {braces}")
                    _check_formula_text(re.sub(r"\{[^{}]+\}", "1", k["expr"]).lstrip("="))
                except SpecError as e:
                    warnings.append(f'KPI "{k["label"]}": {e}. Skipped.')
                    continue
            final.append(k)
        sh["kpis"] = final
        charts = []
        for ch in sh["charts"]:
            kind = ch.get("type") if ch.get("type") in CHART_TYPES else "column"
            s = tables.get(ch.get("sheet"))
            title = str(ch.get("title") or "Chart")[:80]
            if not s:
                warnings.append(f'Chart "{title}" refers to an unknown sheet and was skipped.')
                continue
            names = [c["name"] for c in s["columns"]]
            types = {c["name"]: c["type"] for c in s["columns"]}
            vals = [v for v in (ch.get("value_columns") or []) if v in names and types[v] in NUMERIC_TYPES]
            item = {"type": kind, "title": title, "sheet": s["name"], "value_columns": vals[:4],
                    "agg": ch.get("agg") if ch.get("agg") in AGGS else "sum",
                    "stacked": bool(ch.get("stacked", False))}
            if kind == "scatter":
                if len(vals) < 2:
                    warnings.append(f'Scatter chart "{title}" needs two numeric columns; skipped.')
                    continue
                item["value_columns"] = vals[:2]
            else:
                cat = ch.get("category_column")
                if cat not in names:
                    warnings.append(f'Chart "{title}" has an unknown category column; skipped.')
                    continue
                item["category_column"] = cat
                if not vals:
                    item["agg"] = "count"
                w = ch.get("where")
                if isinstance(w, dict) and w.get("column") in names:
                    item["where"] = {"column": w["column"], "equals": w.get("equals", w.get("value")),
                                     "op": w.get("op") if w.get("op") in ("=", "<>", ">", "<", ">=", "<=") else "="}
                if not chart_categories(s, item):
                    warnings.append(f'Chart "{title}" has no categories to plot yet (no sample rows or dropdown options); skipped.')
                    continue
            charts.append(item)
        sh["charts"] = charts
    return out, warnings


def chart_categories(sh, ch):
    names = [c["name"] for c in sh["columns"]]
    cc = sh["columns"][names.index(ch["category_column"])]
    idx = names.index(ch["category_column"])
    w = ch.get("where")
    widx = names.index(w["column"]) if w else None
    cats, seen = [], set()
    for row in sh["rows"]:
        if w and not _where_match(row[widx], w):
            continue
        v = row[idx]
        if v in (None, ""):
            continue
        k = _key(v)
        if k not in seen:
            seen.add(k)
            cats.append(_coerce(v, cc["type"]))
    if not cats and cc["type"] == "dropdown":
        cats = list(cc.get("options", []))
    return cats[:40]


# ------------------------------------------------------------- workbook build
def _fill(h):
    return PatternFill("solid", start_color=h, end_color=h)


def _numfmt(ctype, fmt, settings):
    t = fmt if ctype == "formula" else ctype
    cur = settings.get("currency_symbol", "$")
    return {"currency": f'"{cur}"#,##0.00', "percent": "0.0%", "date": "yyyy-mm-dd",
            "number": "#,##0.##", "integer": "#,##0"}.get(t, "General")


def _style_font(T, **kw):
    return Font(name=T["font"], **kw)


class _Lists:
    def __init__(self, wb):
        self.wb, self.ws, self.n = wb, None, 0

    def ref(self, options):
        if self.ws is None:
            self.ws = self.wb.create_sheet("Lists")
            self.ws.sheet_state = "hidden"
        self.n += 1
        col = get_column_letter(self.n)
        for i, o in enumerate(options, 1):
            c = self.ws[f"{col}{i}"]
            c.value = o
            c.data_type = "s"
        return f"=Lists!${col}$1:${col}${len(options)}"


def _back_link(ws, cell, target, text, T):
    c = ws[cell]
    c.value = text
    c.hyperlink = Hyperlink(ref=cell, location=f"{_q(target)}!A1", display=text)
    c.font = _style_font(T, size=10, underline="single", color=T["primary"])


CF_COLORS = {"red": ("FFC7CE", "9C0006"), "green": ("C6EFCE", "006100"),
             "amber": ("FFEB9C", "9C5700"), "blue": ("DDEBF7", "1F4E78")}


def _apply_cf(ws, rule, letters, last, key_col, T):
    L = letters[rule["column"]]
    rng = f"{L}{FIRST_ROW}:{L}{last}"
    fc, tc = CF_COLORS[rule["color"]]
    fill, font = PatternFill(start_color=fc, end_color=fc, fill_type="solid"), Font(color=tc)
    k, v = rule["rule"], rule.get("value")
    cf = ws.conditional_formatting
    if k == "data_bar":
        cf.add(rng, DataBarRule(start_type="min", end_type="max", color=T["accent"]))
    elif k == "progress":
        cf.add(rng, DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=T["accent"]))
    elif k == "color_scale":
        cf.add(rng, ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile",
                                   mid_value=50, mid_color="FFEB84", end_type="max", end_color="63BE7B"))
    elif k in ("greater_than", "less_than", "equal"):
        try:
            f = str(float(v))
        except (TypeError, ValueError):
            f = '"' + str(v).replace('"', '""') + '"'
        op = {"greater_than": "greaterThan", "less_than": "lessThan", "equal": "equal"}[k]
        cf.add(rng, CellIsRule(operator=op, formula=[f], fill=fill, font=font))
    elif k == "duplicates":
        cf.add(rng, FormulaRule(formula=[f'AND({L}{FIRST_ROW}<>"",COUNTIF(${L}${FIRST_ROW}:${L}${last},{L}{FIRST_ROW})>1)'],
                                fill=fill, font=font))
    elif k == "blank":
        cf.add(rng, FormulaRule(formula=[f'AND(${key_col}{FIRST_ROW}<>"",{L}{FIRST_ROW}="")'], fill=fill, font=font))
    elif k == "overdue":
        cf.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({L}{FIRST_ROW}),{L}{FIRST_ROW}<TODAY())"], fill=fill, font=font))


def _build_table(ws, sh, spec, T, lists, ranges, sheet_names):
    settings = spec["settings"]
    cols = sh["columns"]
    n = len(cols)
    letters = {c["name"]: get_column_letter(i + 1) for i, c in enumerate(cols)}
    nrows = len(sh["rows"])
    last = FIRST_ROW + max(nrows + sh["extra_rows"], 1) - 1
    key_col = next((letters[c["name"]] for c in cols if c["type"] != "formula"), "A")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = T["accent"]
    lastL = get_column_letter(max(n, 4))

    ws["A1"] = sh["title"] or sh["name"]
    ws["A1"].font = _style_font(T, size=18, bold=True, color=T["primary"])
    ws.row_dimensions[1].height = 30
    note = sh["notes"] or sh["purpose"]
    if sh["sample_data"] and nrows:
        note = (note + "  " if note else "") + "SAMPLE DATA is included for demonstration. Overwrite or delete these rows with your own."
    ws["A2"] = note
    ws["A2"].font = _style_font(T, size=10, italic=True, color=T["muted"])
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(f"A2:{lastL}2")
    ws.row_dimensions[2].height = 32 if len(note) > 90 else 18
    dash = next((s["name"] for s in spec["sheets"] if s["type"] == "dashboard"), None)
    if dash:
        _back_link(ws, "A3", dash, "< Back to Dashboard", T)

    thin = Side(style="thin", color=T["border"])
    for j, c in enumerate(cols, 1):
        h = ws.cell(HEADER_ROW, j, c["name"])
        h.font = _style_font(T, bold=True, color=T["on_primary"])
        h.fill = _fill(T["primary"])
        h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        h.border = Border(bottom=Side(style="medium", color=T["accent"]))
    ws.row_dimensions[HEADER_ROW].height = 30

    body_font = _style_font(T, size=11, color=T["text"])
    for i in range(max(nrows + sh["extra_rows"], 1)):
        r = FIRST_ROW + i
        base = _fill(T["surface"]) if i % 2 else (_fill(T["background"]) if T["dark"] else None)
        for j, c in enumerate(cols, 1):
            cell = ws.cell(r, j)
            if c["type"] == "formula":
                f = resolve_formula(c["formula"], letters, r, FIRST_ROW, ranges)
                cell.value = f"={f}" if c["name"] == "" or key_col == letters[c["name"]] else f'=IF({key_col}{r}="","",{f})'
            elif i < nrows:
                v = _coerce(sh["rows"][i][j - 1], c["type"])
                cell.value = v
                if isinstance(v, str) and v.startswith("="):
                    cell.data_type = "s"
            cell.number_format = _numfmt(c["type"], c.get("format"), settings)
            cell.font = body_font
            cell.border = Border(bottom=thin)
            if base:
                cell.fill = base
            if c["type"] in ("number", "currency", "percent", "formula"):
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c["type"] in ("date", "dropdown"):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(vertical="center", wrap_text=False)

    # widths
    for j, c in enumerate(cols, 1):
        sample = [len(str(r[j - 1])) for r in sh["rows"][:200] if r[j - 1] not in (None, "")]
        w = c.get("width") or min(max(len(c["name"]) + 4, (max(sample) + 3) if sample else 0, 12), 42)
        ws.column_dimensions[get_column_letter(j)].width = w

    # validation
    for j, c in enumerate(cols, 1):
        L = get_column_letter(j)
        rng = f"{L}{FIRST_ROW}:{L}{last}"
        dv = None
        if c["type"] == "dropdown":
            opts = c["options"]
            inline = ",".join(opts)
            if len(inline) <= 250 and not any("," in o or '"' in o for o in opts):
                dv = DataValidation(type="list", formula1=f'"{inline}"', allow_blank=True)
            else:
                dv = DataValidation(type="list", formula1=lists.ref(opts), allow_blank=True)
            dv.error, dv.errorTitle = "Please choose a value from the list.", "Invalid entry"
        elif c["type"] in ("number", "currency"):
            lo, hi = c.get("min"), c.get("max")
            if lo is not None and hi is not None:
                dv = DataValidation(type="decimal", operator="between", formula1=str(lo), formula2=str(hi), allow_blank=True)
            elif lo is not None:
                dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1=str(lo), allow_blank=True)
            elif hi is not None:
                dv = DataValidation(type="decimal", operator="lessThanOrEqual", formula1=str(hi), allow_blank=True)
            else:
                dv = DataValidation(type="decimal", operator="between", formula1="-1000000000000", formula2="1000000000000", allow_blank=True)
            dv.error, dv.errorTitle = "Please enter a number.", "Invalid entry"
        elif c["type"] == "percent":
            dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=True)
            dv.error, dv.errorTitle = "Enter a percentage such as 25% (a value between 0 and 1).", "Invalid entry"
        elif c["type"] == "date":
            dv = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True)
            dv.error, dv.errorTitle = "Please enter a valid date.", "Invalid entry"
        if dv:
            dv.showErrorMessage = True
            dv.add(rng)
            ws.add_data_validation(dv)

    for rule in sh["conditional"]:
        _apply_cf(ws, rule, letters, last, key_col, T)

    ws.freeze_panes = f"A{FIRST_ROW}"
    ws.print_title_rows = f"{HEADER_ROW}:{HEADER_ROW}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    return {"letters": letters, "last": last, "nrows": nrows}


def _kpi_formula(k, ranges):
    rg = ranges[k["sheet"]][k["column"]]
    w = k.get("where")
    agg = k["agg"]
    if w:
        wr = ranges[k["sheet"]][w["column"]]
        op, val = w.get("op", "="), w.get("equals")
        try:
            float(val)
            isnum = True
        except (TypeError, ValueError):
            isnum = False
        sval = str(val).replace('"', '""')
        crit = (f'"{sval}"' if not isnum else str(val)) if op == "=" else f'"{op}{sval}"'
        return {"sum": f"SUMIFS({rg},{wr},{crit})",
                "average": f"IFERROR(AVERAGEIFS({rg},{wr},{crit}),0)",
                "count": f"COUNTIFS({wr},{crit})",
                "max": f"MAXIFS({rg},{wr},{crit})", "min": f"MINIFS({rg},{wr},{crit})"}[agg]
    return {"sum": f"SUM({rg})", "average": f"IFERROR(AVERAGE({rg}),0)",
            "count": f'SUMPRODUCT(--({rg}<>""))', "max": f"MAX({rg})", "min": f"MIN({rg})"}[agg]


def _make_chart(ch, T):
    kind = ch["type"]
    if kind in ("column", "bar"):
        c = BarChart()
        c.type = "bar" if kind == "bar" else "col"
        if ch.get("stacked"):
            c.grouping, c.overlap = "stacked", 100
    elif kind == "line":
        c = LineChart()
    elif kind == "area":
        c = AreaChart()
        if ch.get("stacked"):
            c.grouping = "stacked"
    elif kind == "pie":
        c = PieChart()
    elif kind == "doughnut":
        c = DoughnutChart()
        c.holeSize = 55
    else:
        c = ScatterChart()
        c.style = 13
    c.title = ch["title"]
    c.width, c.height = 13.0, 7.0
    if hasattr(c, "x_axis"):
        c.x_axis.delete = False
        c.y_axis.delete = False
    return c


def _style_series(c, ch, T, npoints):
    pal = _palette(T)
    kind = ch["type"]
    if kind in ("pie", "doughnut"):
        s = c.series[0]
        for i in range(npoints):
            pt = DataPoint(idx=i)
            pt.graphicalProperties.solidFill = pal[i % len(pal)]
            s.dPt.append(pt)
        c.dataLabels = DataLabelList()
        c.dataLabels.showPercent = True
        c.dataLabels.showVal = False
        c.dataLabels.showCatName = False
        c.dataLabels.showSerName = False
        c.dataLabels.showLegendKey = False
        return
    for s, color in zip(c.series, pal):
        if kind == "line":
            s.graphicalProperties.line.solidFill = color
            s.graphicalProperties.line.width = 28000
            s.smooth = False
        elif kind == "scatter":
            s.marker.symbol = "circle"
            s.marker.size = 7
            s.marker.graphicalProperties.solidFill = color
            s.graphicalProperties.line.noFill = True
        else:
            s.graphicalProperties.solidFill = color


def _build_dashboard(ws, sh, spec, T, meta, ranges, wb):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = T["primary"]
    bg = _fill(T["background"])
    for r in range(1, 121):
        for c in range(1, 34):
            ws.cell(r, c).fill = bg
    ws.column_dimensions["A"].width = 2
    for i in range(2, 18):
        ws.column_dimensions[get_column_letter(i)].width = 2.5 if (i - 1) % 4 == 0 else 11
    ws.column_dimensions["S"].width = 3
    ws.column_dimensions["T"].width = 24
    for i in range(21, 27):
        ws.column_dimensions[get_column_letter(i)].width = 14

    for c in range(1, 18):
        ws.cell(1, c).fill = _fill(T["primary"])
    ws.row_dimensions[1].height = 8
    ws["B2"] = sh["title"] or sh["name"]
    ws["B2"].font = _style_font(T, size=24, bold=True, color=T["primary"])
    ws["B2"].alignment = Alignment(vertical="center")
    ws.merge_cells("B2:P2")
    ws.row_dimensions[2].height = 40
    ws["B3"] = sh["purpose"]
    ws["B3"].font = _style_font(T, size=11, italic=True, color=T["muted"])
    ws.merge_cells("B3:P3")
    col = 2
    for s in spec["sheets"]:
        if s["name"] == sh["name"] or col > 16:
            continue
        cell = ws.cell(4, col)
        cell.value = "> " + s["name"]
        cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"{_q(s['name'])}!A1", display=cell.value)
        cell.font = _style_font(T, size=10, underline="single", color=T["primary"])
        col += 2

    # KPI cards
    kpi_cells = {}
    for i, k in enumerate(sh["kpis"]):
        grp, pos = divmod(i, 4)
        c0 = 2 + 4 * pos
        r0 = 6 + 3 * grp
        L0, L2 = get_column_letter(c0), get_column_letter(c0 + 2)
        ws.merge_cells(f"{L0}{r0}:{L2}{r0}")
        ws.merge_cells(f"{L0}{r0 + 1}:{L2}{r0 + 1}")
        for cc in range(c0, c0 + 3):
            ws.cell(r0, cc).fill = _fill(T["surface"])
            ws.cell(r0 + 1, cc).fill = _fill(T["surface"])
            ws.cell(r0, cc).border = Border(top=Side(style="thick", color=T["accent"]))
        lab, val = ws[f"{L0}{r0}"], ws[f"{L0}{r0 + 1}"]
        lab.value = k["label"]
        lab.font = _style_font(T, size=10, bold=True, color=T["muted"])
        lab.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        val.font = _style_font(T, size=22, bold=True, color=T["primary"])
        val.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        val.number_format = _numfmt(k["format"], None, spec["settings"])
        ws.row_dimensions[r0].height = 20
        ws.row_dimensions[r0 + 1].height = 36
        ws.row_dimensions[r0 + 2].height = 10
        kpi_cells[k["label"]] = f"{L0}{r0 + 1}"
        k["_cell"] = val
    for k in sh["kpis"]:
        cell = k.pop("_cell")
        if "expr" in k:
            e = re.sub(r"\{([^{}]+)\}", lambda m: kpi_cells[m.group(1)], k["expr"].lstrip("="))
            _check_formula_text(e)
            cell.value = f"={e}"
        else:
            cell.value = "=" + _kpi_formula(k, ranges)

    # charts + chart data blocks
    chart_row = 6 + 3 * max(1, -(-len(sh["kpis"]) // 4)) + 1
    block_row = 6
    for idx, ch in enumerate(sh["charts"]):
        s = next(x for x in spec["sheets"] if x["name"] == ch["sheet"])
        tws = wb[s["name"]]
        m = meta[s["name"]]
        c = _make_chart(ch, T)
        anchor = f"{'B' if idx % 2 == 0 else 'J'}{chart_row + (idx // 2) * 16}"
        if ch["type"] == "scatter":
            xn, yn = ch["value_columns"]
            xi = [i + 1 for i, cc in enumerate(s["columns"]) if cc["name"] == xn][0]
            yi = [i + 1 for i, cc in enumerate(s["columns"]) if cc["name"] == yn][0]
            lastr = FIRST_ROW + max(m["nrows"], 2) - 1
            ser = Series(Reference(tws, min_col=yi, min_row=FIRST_ROW, max_row=lastr),
                         Reference(tws, min_col=xi, min_row=FIRST_ROW, max_row=lastr), title=yn)
            c.series.append(ser)
            c.x_axis.title, c.y_axis.title = xn, yn
            c.legend = None
            _style_series(c, ch, T, 0)
        else:
            cats = chart_categories(s, ch)
            cat_rng = ranges[s["name"]][ch["category_column"]]
            heads = ch["value_columns"] or ["Count"]
            ws.cell(block_row, 20, ch["title"] + " (chart data)").font = _style_font(T, bold=True, color=T["primary"])
            hdr = block_row + 1
            ws.cell(hdr, 20, ch["category_column"])
            for k, h in enumerate(heads):
                ws.cell(hdr, 21 + k, h)
            for cc in range(20, 21 + len(heads)):
                x = ws.cell(hdr, cc)
                x.font = _style_font(T, bold=True, color=T["on_primary"])
                x.fill = _fill(T["primary"])
                x.alignment = Alignment(horizontal="center")
            cat_type = next(cc["type"] for cc in s["columns"] if cc["name"] == ch["category_column"])
            w = ch.get("where")
            for ri, cat in enumerate(cats, 1):
                r = hdr + ri
                cell = ws.cell(r, 20, cat)
                cell.font = _style_font(T, color=T["text"])
                if cat_type == "date":
                    cell.number_format = "yyyy-mm-dd"
                extra = ""
                if w:
                    wr = ranges[s["name"]][w["column"]]
                    op, val = w.get("op", "="), w.get("equals")
                    try:
                        float(val)
                        isnum = True
                    except (TypeError, ValueError):
                        isnum = False
                    sval = str(val).replace('"', '""')
                    crit = (f'"{sval}"' if not isnum else str(val)) if op == "=" else f'"{op}{sval}"'
                    extra = f",{wr},{crit}"
                for k, h in enumerate(heads):
                    vr = ranges[s["name"]][h] if ch["value_columns"] else None
                    if ch["agg"] == "count" or vr is None:
                        f = f"COUNTIFS({cat_rng},$T{r}{extra})"
                    elif ch["agg"] == "average":
                        f = f"IFERROR(AVERAGEIFS({vr},{cat_rng},$T{r}{extra}),0)"
                    elif ch["agg"] in ("max", "min"):
                        f = f"{ch['agg'].upper()}IFS({vr},{cat_rng},$T{r}{extra})"
                    else:
                        f = f"SUMIFS({vr},{cat_rng},$T{r}{extra})"
                    cc = ws.cell(r, 21 + k, "=" + f)
                    cc.font = _style_font(T, color=T["text"])
                    vcol = next((x for x in s["columns"] if x["name"] == h), None)
                    cc.number_format = _numfmt(vcol["type"], vcol.get("format"), spec["settings"]) if vcol else "0"
            last = hdr + len(cats)
            data = Reference(ws, min_col=21, max_col=20 + len(heads), min_row=hdr, max_row=last)
            c.add_data(data, titles_from_data=True)
            c.set_categories(Reference(ws, min_col=20, min_row=hdr + 1, max_row=last))
            if len(heads) == 1 and ch["type"] not in ("pie", "doughnut"):
                c.legend = None
            else:
                c.legend.position = "b"
            _style_series(c, ch, T, len(cats))
            block_row = last + 3
        ws.add_chart(c, anchor)
    if sh["charts"]:
        note_r = chart_row + ((len(sh["charts"]) + 1) // 2) * 16 + 1
        n = ws.cell(note_r, 2, "Charts read the formula-driven 'chart data' tables in columns T onward. "
                    "To plot a new category, type its name in the first column of that table.")
        n.font = _style_font(T, size=9, italic=True, color=T["muted"])


def _build_instructions(ws, spec, T):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = T["muted"]
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 90
    bg = _fill(T["background"])
    for r in range(1, 80):
        for c in range(1, 5):
            ws.cell(r, c).fill = bg
    ws["B2"] = "How to use this workbook"
    ws["B2"].font = _style_font(T, size=22, bold=True, color=T["primary"])
    ws.row_dimensions[2].height = 36
    r = 4
    sh = next((s for s in spec["sheets"] if s["type"] == "instructions"), None)
    lines = (sh or {}).get("lines") or [
        "1. Start on the Dashboard to see your summary.",
        "2. Enter your own information in the data sheets (the sample rows show the expected format).",
        "3. Cells with a drop-down arrow only accept values from the list.",
        "4. Formula columns (and the dashboard) update automatically. Do not overwrite them.",
        "5. Charts read from the 'chart data' tables on the dashboard (columns T onward).",
        "6. Excel recalculates all formulas when the file is opened; press F9 if values look stale."]
    for ln in lines:
        c = ws.cell(r, 2, ln)
        c.font = _style_font(T, size=11, color=T["text"])
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.row_dimensions[r].height = 20 if len(ln) < 110 else 34
        r += 1
    r += 1
    ws.cell(r, 2, "Sheet").font = _style_font(T, bold=True, color=T["on_primary"])
    ws.cell(r, 3, "Purpose").font = _style_font(T, bold=True, color=T["on_primary"])
    ws.cell(r, 2).fill = ws.cell(r, 3).fill = _fill(T["primary"])
    r += 1
    for s in spec["sheets"]:
        if s["type"] == "instructions":
            continue
        c = ws.cell(r, 2, s["name"])
        c.hyperlink = Hyperlink(ref=c.coordinate, location=f"{_q(s['name'])}!A1", display=s["name"])
        c.font = _style_font(T, underline="single", color=T["primary"])
        p = ws.cell(r, 3, s["purpose"] or s["title"])
        p.font = _style_font(T, color=T["text"])
        p.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    if any(s["type"] == "table" and s["sample_data"] and s["rows"] for s in spec["sheets"]):
        r += 1
        c = ws.cell(r, 2, "Sample data notice: the example rows are placeholders for demonstration only. Replace them with your own data.")
        c.font = _style_font(T, italic=True, color=T["muted"])


def build_workbook(spec) -> tuple[bytes, dict, list]:
    """Build a real .xlsx. Returns (xlsx_bytes, normalised_spec, warnings)."""
    spec, warnings = normalize_spec(spec)
    T = _T(spec)
    ranges = table_ranges(spec)
    wb = Workbook()
    wb.remove(wb.active)
    lists = _Lists(wb)
    sheets = {}
    for sh in spec["sheets"]:
        sheets[sh["name"]] = wb.create_sheet(sh["name"])
    meta = {}
    for sh in spec["sheets"]:
        if sh["type"] == "table":
            meta[sh["name"]] = _build_table(sheets[sh["name"]], sh, spec, T, lists, ranges, list(sheets))
    for sh in spec["sheets"]:
        if sh["type"] == "dashboard":
            _build_dashboard(sheets[sh["name"]], sh, spec, T, meta, ranges, wb)
        elif sh["type"] == "instructions":
            _build_instructions(sheets[sh["name"]], spec, T)
    if spec["settings"]["add_instructions"] and not any(s["type"] == "instructions" or s["name"].lower() == "instructions"
                                                        for s in spec["sheets"]):
        ws = wb.create_sheet("Instructions", len(spec["sheets"]))
        _build_instructions(ws, spec, T)
    if lists.ws is not None and "Lists" in wb.sheetnames:
        wb.move_sheet("Lists", offset=len(wb.sheetnames))
    wb.properties.title = spec["workbook_name"]
    wb.properties.creator = "AI Excel Studio"
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), spec, warnings


def suggest_filename(spec) -> str:
    n = re.sub(r"[^A-Za-z0-9]+", "_", spec.get("workbook_name", "Workbook")).strip("_") or "Workbook"
    return f"{n}.xlsx"


# --------------------------------------------------------- safe edit operations
class _Pending(Exception):
    pass


DESTRUCTIVE = {"remove_sheet", "remove_column"}


def _find_sheet(spec, name):
    for s in spec["sheets"]:
        if s["name"].lower() == str(name).strip().lower():
            return s
    raise SpecError(f'Sheet "{name}" was not found.')


def _find_col(sh, name):
    for i, c in enumerate(sh["columns"]):
        if c["name"].lower() == str(name).strip().lower():
            return i, c
    raise SpecError(f'Column "{name}" was not found on "{sh["name"]}".')


def _get_dash(spec, name=None):
    for s in spec["sheets"]:
        if s["type"] == "dashboard" and (not name or s["name"].lower() == str(name).lower()):
            return s
    d = {"name": "Dashboard", "type": "dashboard", "title": spec.get("workbook_name", "Dashboard"),
         "purpose": "Summary of your key numbers.", "kpis": [], "charts": []}
    spec["sheets"].insert(0, d)
    return d


def describe_destructive(ops) -> list[str]:
    out = []
    for op in ops:
        if op.get("op") == "remove_sheet":
            out.append(f'Delete worksheet "{op.get("name")}" and all of its data')
        elif op.get("op") == "remove_column":
            out.append(f'Delete column "{op.get("column")}" (and its data) from "{op.get("sheet")}"')
    return out


def _rename_refs(spec, old, new):
    for s in spec["sheets"]:
        for k in s.get("kpis", []) + s.get("charts", []):
            if str(k.get("sheet", "")).lower() == old.lower():
                k["sheet"] = new
        for c in s.get("columns", []):
            if c.get("formula"):
                c["formula"] = re.sub(r"\{" + re.escape(old) + r"\.", "{" + new + ".", c["formula"])


def _apply_one(spec, op):
    k = op.get("op")
    if k == "set_theme":
        cur = spec.get("theme") or {}
        base = {"name": op["theme"]} if op.get("theme") in THEMES else dict(cur)
        for key, v in (op.get("colors") or {}).items():
            if key in ("primary", "accent", "background", "surface", "text"):
                base[key] = v
        if op.get("font"):
            base["font"] = op["font"]
        if not op.get("theme") in THEMES and op.get("colors"):
            base["name"] = "Custom"
        spec["theme"] = norm_theme(base)
        return f'Theme set to {spec["theme"]["name"]}'
    if k == "set_workbook_name":
        spec["workbook_name"] = str(op["name"])
        return "Workbook renamed"
    if k == "set_setting":
        if op.get("key") in ("currency_symbol", "add_instructions"):
            spec.setdefault("settings", {})[op["key"]] = op.get("value")
            return f'Setting {op["key"]} updated'
        raise SpecError("Unknown setting")
    if k == "add_sheet":
        sh = op["sheet"]
        if any(s["name"].lower() == str(sh.get("name", "")).lower() for s in spec["sheets"]):
            raise SpecError(f'A sheet named "{sh.get("name")}" already exists.')
        pos = op.get("position")
        if isinstance(pos, int) and 0 <= pos <= len(spec["sheets"]):
            spec["sheets"].insert(pos, sh)
        else:
            spec["sheets"].append(sh)
        return f'Added sheet "{sh.get("name")}"'
    if k == "remove_sheet":
        s = _find_sheet(spec, op["name"])
        spec["sheets"].remove(s)
        return f'Removed sheet "{s["name"]}"'
    if k == "rename_sheet":
        s = _find_sheet(spec, op["name"])
        old, s["name"] = s["name"], str(op["new_name"])
        _rename_refs(spec, old, s["name"])
        return f'Renamed "{old}" to "{s["name"]}"'
    if k == "reorder_sheets":
        order = [str(x).lower() for x in op.get("order", [])]
        spec["sheets"].sort(key=lambda s: order.index(s["name"].lower()) if s["name"].lower() in order else 999)
        return "Reordered sheets"
    if k == "update_sheet":
        s = _find_sheet(spec, op["name"])
        for key, v in (op.get("changes") or {}).items():
            if key in ("title", "purpose", "notes", "sample_data", "extra_rows"):
                s[key] = v
        return f'Updated sheet "{s["name"]}"'
    if k == "add_column":
        s = _find_sheet(spec, op["sheet"])
        col = op["column"]
        if any(c["name"].lower() == str(col.get("name", "")).lower() for c in s["columns"]):
            raise SpecError(f'Column "{col.get("name")}" already exists on "{s["name"]}".')
        pos = op.get("position")
        pos = pos - 1 if isinstance(pos, int) and 1 <= pos <= len(s["columns"]) + 1 else len(s["columns"])
        s["columns"].insert(pos, col)
        vals = op.get("values") or []
        for i, row in enumerate(s["rows"]):
            row.insert(pos, vals[i] if i < len(vals) else None)
        return f'Added column "{col.get("name")}" to "{s["name"]}"'
    if k == "remove_column":
        s = _find_sheet(spec, op["sheet"])
        i, c = _find_col(s, op["column"])
        s["columns"].pop(i)
        for row in s["rows"]:
            if i < len(row):
                row.pop(i)
        return f'Removed column "{c["name"]}" from "{s["name"]}"'
    if k == "update_column":
        s = _find_sheet(spec, op["sheet"])
        i, c = _find_col(s, op["column"])
        ch = dict(op.get("changes") or {})
        new_name = ch.pop("name", None)
        c.update({kk: v for kk, v in ch.items() if kk in ("type", "options", "formula", "format", "min", "max", "width")})
        if new_name and new_name != c["name"]:
            old = c["name"]
            c["name"] = str(new_name)
            for cc in s["columns"]:
                if cc.get("formula"):
                    cc["formula"] = cc["formula"].replace(f"[{old}]", f"[{new_name}]").replace(f"[{old}:prev]", f"[{new_name}:prev]")
            for d in spec["sheets"]:
                for kk in d.get("kpis", []) + d.get("charts", []):
                    if kk.get("sheet") == s["name"]:
                        if kk.get("column") == old:
                            kk["column"] = new_name
                        if kk.get("category_column") == old:
                            kk["category_column"] = new_name
                        kk["value_columns"] = [new_name if v == old else v for v in kk.get("value_columns", [])]
                        if isinstance(kk.get("where"), dict) and kk["where"].get("column") == old:
                            kk["where"]["column"] = new_name
                for cc in d.get("columns", []):
                    if cc.get("formula"):
                        cc["formula"] = cc["formula"].replace(f"{{{s['name']}.{old}}}", f"{{{s['name']}.{new_name}}}")
            for r in s.get("conditional", []):
                if r.get("column") == old:
                    r["column"] = new_name
        return f'Updated column "{c["name"]}" on "{s["name"]}"'
    if k == "add_rows":
        s = _find_sheet(spec, op["sheet"])
        n = len(s["columns"])
        for r in op.get("rows", []):
            r = list(r)[:n]
            s["rows"].append(r + [None] * (n - len(r)))
        return f'Added {len(op.get("rows", []))} row(s) to "{s["name"]}"'
    if k == "set_extra_rows":
        s = _find_sheet(spec, op["sheet"])
        s["extra_rows"] = int(op["extra_rows"])
        return f'"{s["name"]}" now has {s["extra_rows"]} blank formula-ready rows'
    if k == "add_kpi":
        d = _get_dash(spec, op.get("dashboard"))
        d.setdefault("kpis", []).append(op["kpi"])
        return f'Added KPI "{op["kpi"].get("label")}"'
    if k == "remove_kpi":
        d = _get_dash(spec, op.get("dashboard"))
        d["kpis"] = [x for x in d.get("kpis", []) if x.get("label", "").lower() != str(op["label"]).lower()]
        return f'Removed KPI "{op["label"]}"'
    if k == "add_chart":
        d = _get_dash(spec, op.get("dashboard"))
        d.setdefault("charts", []).append(op["chart"])
        return f'Added chart "{op["chart"].get("title")}"'
    if k == "remove_chart":
        d = _get_dash(spec, op.get("dashboard"))
        d["charts"] = [x for x in d.get("charts", []) if x.get("title", "").lower() != str(op["title"]).lower()]
        return f'Removed chart "{op["title"]}"'
    if k == "add_conditional":
        s = _find_sheet(spec, op["sheet"])
        s.setdefault("conditional", []).append(op["rule"])
        return f'Added conditional formatting on "{s["name"]}"'
    raise SpecError(f"Unsupported operation: {k}")


def apply_operations(spec, ops, confirmed=False):
    """Apply a list of validated operations to a COPY of the spec.

    Returns (new_spec, log, pending_descriptions, warnings). If destructive
    operations are present and `confirmed` is False nothing is applied.
    """
    ops = [o for o in (ops or []) if isinstance(o, dict)][:40]
    pend = describe_destructive(ops)
    if pend and not confirmed:
        return spec, [], pend, []
    new, log = copy.deepcopy(spec), []
    for op in ops:
        try:
            log.append(_apply_one(new, op))
        except SpecError as e:
            log.append(f"Skipped: {e}")
        except (KeyError, TypeError, ValueError) as e:
            log.append(f"Skipped an invalid '{op.get('op')}' operation ({type(e).__name__}: {e}).")
    new, warns = normalize_spec(new)
    return new, log, [], warns


def preservation_report(old, new) -> list[str]:
    issues = []
    nm = {s["name"]: s for s in new["sheets"]}
    for s in old["sheets"]:
        ns = nm.get(s["name"])
        if ns is None:
            issues.append(f'Sheet "{s["name"]}" is no longer present under that name.')
            continue
        if s["type"] != "table" or ns["type"] != "table":
            continue
        if len(ns["rows"]) < len(s["rows"]):
            issues.append(f'"{s["name"]}" now has fewer data rows ({len(s["rows"])} -> {len(ns["rows"])}).')
        lost = {c["name"] for c in s["columns"]} - {c["name"] for c in ns["columns"]}
        if lost:
            issues.append(f'"{s["name"]}" lost column(s): {", ".join(sorted(lost))}.')
    return issues


# ---------------------------------------------------------------- Gemini / AI
def get_api_key() -> str | None:
    key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if key:
        return key
    try:
        import streamlit as st
        return st.secrets.get("GEMINI_API_KEY") or None
    except Exception:
        return None


def get_model() -> str:
    m = os.environ.get("GEMINI_MODEL")
    if not m:
        try:
            import streamlit as st
            m = st.secrets.get("GEMINI_MODEL")
        except Exception:
            m = None
    return m or "gemini-2.5-flash"


def _call_gemini(parts, system, json_mode=True, temperature=0.4) -> str:
    key = get_api_key()
    if not key:
        raise AIError("No Gemini API key found. Add GEMINI_API_KEY to Streamlit secrets or an environment variable.")
    try:
        from google import genai
        from google.genai import types
    except ImportError as e:
        raise AIError("The 'google-genai' package is not installed (see requirements.txt).") from e
    try:
        client = genai.Client(api_key=key)
        cfg = types.GenerateContentConfig(
            system_instruction=system, temperature=temperature,
            response_mime_type="application/json" if json_mode else None)
        resp = client.models.generate_content(model=get_model(), contents=parts, config=cfg)
        text = resp.text
    except Exception as e:  # noqa: BLE001 - convert any SDK error to a friendly one
        msg = str(e)
        if "API_KEY_INVALID" in msg or "API key not valid" in msg or "403" in msg[:40]:
            raise AIError("The Gemini API key is invalid or not permitted. Check GEMINI_API_KEY.") from e
        if "429" in msg or "RESOURCE_EXHAUSTED" in msg:
            raise AIError("Gemini free-tier quota or rate limit reached. Wait a minute and try again, or use a different key/model.") from e
        if "404" in msg or "NOT_FOUND" in msg:
            raise AIError(f"Model '{get_model()}' was not found. Set GEMINI_MODEL to a supported Gemini Flash model.") from e
        raise AIError(f"Gemini request failed: {msg[:300]}") from e
    if not text:
        raise AIError("Gemini returned an empty response (it may have been blocked). Try rephrasing.")
    return text


def _parse_json(text: str):
    t = text.strip()
    t = re.sub(r"^```(?:json)?|```$", "", t, flags=re.M).strip()
    try:
        return json.loads(t)
    except json.JSONDecodeError:
        a, b = t.find("{"), t.rfind("}")
        if a >= 0 and b > a:
            return json.loads(t[a:b + 1])
        raise


def _json_call(parts, system, temperature=0.4):
    text = _call_gemini(parts, system, True, temperature)
    try:
        return _parse_json(text)
    except json.JSONDecodeError:
        text = _call_gemini(parts + ["Your previous reply was not valid JSON. Reply again with ONLY the JSON object."],
                            system, True, 0.1)
        try:
            return _parse_json(text)
        except json.JSONDecodeError as e:
            raise AIError("The AI returned invalid JSON twice. Please try again.") from e


SPEC_GUIDE = """
You design Excel workbooks by emitting a JSON "workbook spec". A separate program turns the spec into a real .xlsx.
You never write code. Output ONLY a JSON object with this shape:

{
 "workbook_name": "Professional Life Planner",
 "theme": {"name": "<one of: %(themes)s>", "primary":"#RRGGBB","accent":"#RRGGBB","background":"#RRGGBB","surface":"#RRGGBB","text":"#RRGGBB","font":"Calibri"},
 "settings": {"currency_symbol": "$", "add_instructions": true},
 "assumptions": ["things you assumed because the request/images did not say"],
 "design_notes": "short note on the look",
 "sheets": [
  {"name":"Dashboard","type":"dashboard","title":"...","purpose":"...",
   "kpis":[
     {"label":"Total Income","sheet":"Transactions","column":"Amount","agg":"sum|average|count|max|min","format":"currency|number|integer|percent",
      "where":{"column":"Type","equals":"Income","op":"="}},
     {"label":"Net Savings","expr":"{Total Income}-{Total Expenses}","format":"currency"}],
   "charts":[
     {"type":"column|bar|line|area|pie|doughnut|scatter","title":"...","sheet":"Transactions",
      "category_column":"Category","value_columns":["Amount"],"agg":"sum|average|count|max|min","stacked":false,
      "where":{"column":"Type","equals":"Expense"}}]},
  {"name":"Transactions","type":"table","title":"...","purpose":"...","notes":"how to use this sheet",
   "sample_data":true,"extra_rows":100,
   "columns":[
     {"name":"Date","type":"date"},
     {"name":"Type","type":"dropdown","options":["Income","Expense"]},
     {"name":"Amount","type":"currency","min":0},
     {"name":"Tax","type":"formula","formula":"[Amount]*0.1","format":"currency"}],
   "rows":[["2025-01-03","Income",2500]],
   "conditional":[{"column":"Amount","rule":"data_bar|color_scale|progress|greater_than|less_than|equal|duplicates|blank|overdue","value":0,"color":"red|green|amber|blue"}]},
  {"name":"Instructions","type":"instructions","lines":["step 1 ...","step 2 ..."]}
 ]
}

RULES
- Column types: text, number, currency, percent (store as fraction 0-1, e.g. 0.25), date (YYYY-MM-DD), dropdown (needs options), formula.
- Formula columns: write a spreadsheet formula WITHOUT the leading "=". Refer to the same row with [Column Name], to the previous row with
  [Column Name:prev] (running balances), and to a whole column of a table sheet with {Sheet Name.Column Name}, e.g.
  SUMIFS({Transactions.Amount},{Transactions.Category},[Category]). Formula columns are NEVER included in "rows" (use null).
  Allowed: SUM, AVERAGE, MIN, MAX, COUNT, SUMIF(S), COUNTIF(S), IF, IFERROR, AND, OR, ROUND, TEXT, DATE, TODAY, INDEX, MATCH, XLOOKUP, arithmetic.
  Never reference other workbooks, URLs, macros or the WEBSERVICE/HYPERLINK/CALL functions.
- KPIs and charts live ONLY on "dashboard" sheets and read from "table" sheets. KPI "expr" may only use {Label} of other aggregate KPIs.
- Charts group a category column. Charts summarise whatever categories appear in the sample rows (or dropdown options), so for trend charts
  add a plain text "Month" (or Week) dropdown column rather than relying on dates. Use "where" to filter (e.g. only Expense rows).
- Every table sheet needs 8-30 realistic SAMPLE rows (sample_data true) unless the user supplied real data. Rows must have one value per column.
- Always include one dashboard sheet first, the data/tracker sheets, and optionally an instructions sheet. 3-8 sheets total.
- Sheet names: max 31 chars, no : \\ / ? * [ ] .
- Use a theme from the list unless a reference image gives palette colours; then set explicit hex colours.
- Habit/yes-no trackers: use dropdown columns with options like ["Yes","No"] or ["Done",""].
- Respond with JSON only.
""" % {"themes": ", ".join(THEMES)}

OPS_GUIDE = """
You edit an existing workbook by returning JSON: {"summary":"one-sentence plain description","operations":[...],"clarification":null}
If the request is ambiguous, or would destroy data and you are unsure, return {"summary":"","operations":[],"clarification":"your question"}.
Make the smallest change that satisfies the request; never rebuild the whole workbook for a local change.
Supported operations (anything else is rejected):
 {"op":"set_theme","theme":"<preset name>"}  or  {"op":"set_theme","colors":{"primary":"#hex","accent":"#hex","background":"#hex","surface":"#hex","text":"#hex"},"font":"Calibri"}
 {"op":"set_workbook_name","name":"..."}   {"op":"set_setting","key":"currency_symbol|add_instructions","value":...}
 {"op":"add_sheet","sheet":{<full sheet object as in the spec>},"position":int?}
 {"op":"remove_sheet","name":"..."}   {"op":"rename_sheet","name":"...","new_name":"..."}   {"op":"reorder_sheets","order":["..."]}
 {"op":"update_sheet","name":"...","changes":{"title":"","purpose":"","notes":"","sample_data":false,"extra_rows":200}}
 {"op":"add_column","sheet":"...","column":{<column object>},"position":int?,"values":[...]?}
 {"op":"remove_column","sheet":"...","column":"..."}
 {"op":"update_column","sheet":"...","column":"...","changes":{"name":"","type":"","options":[],"formula":"","format":"","min":0,"max":1}}
 {"op":"add_rows","sheet":"...","rows":[[...]]}   {"op":"set_extra_rows","sheet":"...","extra_rows":300}
 {"op":"add_kpi","kpi":{...},"dashboard":"name?"}   {"op":"remove_kpi","label":"..."}
 {"op":"add_chart","chart":{...},"dashboard":"name?"}   {"op":"remove_chart","title":"..."}
 {"op":"add_conditional","sheet":"...","rule":{"column":"...","rule":"...","value":...,"color":"..."}}
Formula/KPI/chart/column syntax is identical to the spec format (formulas use [Column], [Column:prev], {Sheet.Column}; no leading "=").
Respect existing sheet and column names exactly. Respond with JSON only.
"""

DESIGN_GUIDE = """
You are a spreadsheet UI analyst. Study the reference image(s) and return ONLY JSON describing the design:
{"overall_layout":"", "sheets_detected":[{"name":"","purpose":""}], "palette":{"primary":"#hex","accent":"#hex","background":"#hex","surface":"#hex","text":"#hex"},
 "fonts":{"family":"","title_size":0,"body_size":0,"hierarchy":""}, "tables":[{"position":"","columns":[""],"notes":""}],
 "kpi_cards":[{"label":"","position":""}], "charts":[{"type":"","position":"","size":"","data_shown":""}],
 "borders_spacing_dividers":"", "icons_navigation_decor":"", "calendars_trackers_progress":"", "visible_formulas_metrics":"",
 "interactive_controls":"", "assumptions":["what you had to guess"], "confidence":"what cannot be known from an image"}
Be honest: if something cannot be determined from the image, list it under assumptions/confidence rather than inventing certainty.
"""


def analyze_images(images, notes="") -> dict:
    from google.genai import types  # imported lazily so the module loads without the SDK
    parts = [f"Analyse these {len(images)} reference image(s). Extra notes from the user: {notes or 'none'}"]
    for name, data, mime in images:
        parts.append(f"Image file: {name}")
        parts.append(types.Part.from_bytes(data=data, mime_type=mime))
    return _json_call(parts, DESIGN_GUIDE, 0.2)


def plan_workbook(project_type, description, design=None, dataset_summary=None, theme_pref=None, images=None) -> dict:
    parts = [f"Project type: {project_type}\nUser description: {description}"]
    if theme_pref and theme_pref != "Let the AI choose":
        parts.append(f"Preferred theme: {theme_pref}")
    if design:
        parts.append("DESIGN SPEC extracted from the user's reference images (mirror it as closely as the spec format allows, "
                     "use its palette as explicit theme colours, and list unknowns under assumptions):\n" + json.dumps(design)[:12000])
    if dataset_summary:
        parts.append("The user imported real data. These sheets will be inserted automatically with their real rows; "
                     "do NOT invent rows for them. Use the exact sheet and column names in dashboard KPIs/charts and formulas:\n" + dataset_summary)
    return _json_call(parts, SPEC_GUIDE, 0.5)


def _compact(spec, keep_rows=3):
    s = copy.deepcopy(spec)
    for sh in s["sheets"]:
        if "rows" in sh:
            sh["rows_total"] = len(sh["rows"])
            sh["rows"] = sh["rows"][:keep_rows]
    return s


def ai_edit(spec, instruction, design=None) -> dict:
    parts = ["CURRENT WORKBOOK SPEC (rows truncated for brevity):\n" + json.dumps(_compact(spec))[:60000],
             "USER REQUEST: " + instruction]
    if design:
        parts.insert(1, "Reference design spec:\n" + json.dumps(design)[:6000])
    res = _json_call(parts, OPS_GUIDE, 0.3)
    if not isinstance(res, dict):
        raise AIError("The AI returned an unexpected format for the edit.")
    res.setdefault("operations", [])
    res.setdefault("summary", "")
    res.setdefault("clarification", None)
    return res


def transcribe(audio: bytes, mime="audio/wav") -> str:
    from google.genai import types
    text = _call_gemini([types.Part.from_bytes(data=audio, mime_type=mime),
                         "Transcribe this audio exactly. Return only the transcript text."],
                        "You are a precise speech transcriber.", json_mode=False, temperature=0.0)
    return text.strip()


def listing_copy(spec) -> dict:
    brief = {"name": spec["workbook_name"], "sheets": [{"name": s["name"], "purpose": s.get("purpose", "")} for s in spec["sheets"]]}
    return _json_call(["Write an Etsy listing for this Excel product, honest and without claims of guaranteed compatibility: "
                       + json.dumps(brief)],
                      'Return JSON {"title":"<=140 chars","description":"plain text","features":["..."],"tags":["13 tags, each <=20 chars"]}.', 0.7)


def test_connection() -> str:
    return _call_gemini(["Reply with the single word: ok"], "Be terse.", json_mode=False, temperature=0).strip()


# ------------------------------------------------------------ dataset import
def _infer_type(s):
    import pandas as pd
    s2 = s.dropna()
    if s2.empty:
        return "text"
    if pd.api.types.is_datetime64_any_dtype(s2):
        return "date"
    if pd.api.types.is_bool_dtype(s2):
        return "text"
    if pd.api.types.is_numeric_dtype(s2):
        return "number"
    if s2.map(lambda x: isinstance(x, (datetime, date))).all():
        return "date"
    if s2.map(lambda x: isinstance(x, str)).all() and re.match(r"^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}", str(s2.iloc[0])):
        try:
            if pd.to_datetime(s2, errors="coerce", format="mixed").notna().mean() >= 0.9:
                return "date"
        except (TypeError, ValueError):
            pass
    if s2.nunique() <= 10 and len(s2) >= 20:
        return "dropdown"
    return "text"


def df_to_sheet(name, df) -> dict:
    import pandas as pd
    df = df.copy().iloc[:MAX_ROWS, :40]
    df.columns = [str(c).strip() or f"Column {i + 1}" for i, c in enumerate(df.columns)]
    cols, rows = [], []
    for c in df.columns:
        t = _infer_type(df[c])
        col = {"name": c, "type": t}
        if t == "dropdown":
            col["options"] = [str(x) for x in df[c].dropna().unique()]
        cols.append(col)
    cols = _norm_columns(cols, [], name)
    for rec in df.itertuples(index=False):
        row = []
        for v, c in zip(rec, cols):
            if v is None or (not isinstance(v, str) and pd.isna(v)):
                row.append(None)
            elif isinstance(v, (pd.Timestamp, datetime, date)):
                row.append(pd.Timestamp(v).strftime("%Y-%m-%d"))
            elif hasattr(v, "item"):
                row.append(v.item())
            else:
                row.append(v)
        rows.append(row)
    return {"name": name, "type": "table", "title": name, "purpose": f"Imported data ({len(rows)} rows).",
            "notes": "Your imported data.", "sample_data": False, "extra_rows": 50, "columns": cols, "rows": rows}


def import_dataset(filename: str, data: bytes) -> tuple[list[dict], list[str]]:
    """Read CSV/XLSX into table-sheet specs. The original bytes are never modified."""
    import pandas as pd
    warns, sheets = [], []
    low = filename.lower()
    if low.endswith(".csv"):
        df = None
        for enc in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                df = pd.read_csv(io.BytesIO(data), encoding=enc)
                break
            except Exception:  # noqa: BLE001
                continue
        if df is None:
            raise SpecError("Could not parse the CSV file.")
        frames = {os.path.splitext(os.path.basename(filename))[0][:31] or "Data": df}
    elif low.endswith((".xlsx", ".xlsm")):
        try:
            z = zipfile.ZipFile(io.BytesIO(data))
            names = z.namelist()
        except zipfile.BadZipFile as e:
            raise SpecError("This is not a valid .xlsx file.") from e
        if any("vbaProject" in n for n in names) or low.endswith(".xlsm"):
            warns.append("This workbook contains macros. Macros are NOT carried over (only the data is imported).")
        if any(n.startswith("xl/externalLinks") for n in names):
            warns.append("External links were found; they are not preserved.")
        if any(n.startswith("xl/charts") for n in names) or any(n.startswith("xl/pivot") for n in names):
            warns.append("Existing charts/PivotTables are not imported; rebuild them via the editor.")
        if any(n.startswith("xl/drawings") or n.startswith("xl/media") for n in names):
            warns.append("Images/shapes are not imported.")
        try:
            wbf = load_workbook(io.BytesIO(data), data_only=False, read_only=True)
            nform = sum(1 for ws in wbf.worksheets for row in ws.iter_rows(max_row=500)
                        for c in row if isinstance(c.value, str) and c.value.startswith("="))
            if nform:
                warns.append("Formulas were converted to their last saved values (cell positions change when imported).")
        except Exception:  # noqa: BLE001
            pass
        try:
            frames = pd.read_excel(io.BytesIO(data), sheet_name=None)
        except Exception as e:  # noqa: BLE001
            raise SpecError(f"Could not read the workbook: {e}") from e
    else:
        raise SpecError("Only .xlsx and .csv files can be imported.")
    used = set()
    for nm, df in frames.items():
        if df.empty:
            continue
        if len(df) > MAX_ROWS:
            warns.append(f'"{nm}" was truncated to {MAX_ROWS} rows.')
        sheets.append(df_to_sheet(_clean_sheet_name(nm, used), df))
    if not sheets:
        raise SpecError("No data was found in the file.")
    return sheets, warns


def profile_sheet(sh) -> dict:
    import pandas as pd
    names = [c["name"] for c in sh["columns"]]
    df = pd.DataFrame(sh["rows"], columns=names)
    return {"rows": len(df), "duplicates": int(df.duplicated().sum()) if len(df) else 0,
            "missing": {c: int(df[c].isna().sum()) for c in names if df[c].isna().any()},
            "types": {c["name"]: c["type"] for c in sh["columns"]}}


def dataset_summary(sheets) -> str:
    out = []
    for sh in sheets:
        out.append(f'Sheet "{sh["name"]}" ({len(sh["rows"])} rows): ' +
                   ", ".join(f'{c["name"]} [{c["type"]}]' for c in sh["columns"]) +
                   "\nFirst rows: " + json.dumps(sh["rows"][:3], default=str))
    return "\n".join(out)


def merge_imported(spec, imported) -> dict:
    spec = copy.deepcopy(spec)
    for imp in imported:
        ex = next((s for s in spec["sheets"] if s["name"].lower() == imp["name"].lower() and s["type"] == "table"), None)
        if ex:
            keep = [c for c in ex["columns"] if c["type"] == "formula"
                    and c["name"].lower() not in {x["name"].lower() for x in imp["columns"]}]
            new = copy.deepcopy(imp)
            for c in keep:
                new["columns"].append(c)
                for r in new["rows"]:
                    r.append(None)
            new["purpose"] = ex.get("purpose") or new["purpose"]
            spec["sheets"][spec["sheets"].index(ex)] = new
        else:
            spec["sheets"].append(copy.deepcopy(imp))
    return spec


# ------------------------------------------------------- preview (Python-side)
_BIN = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
        ast.Div: operator.truediv, ast.Pow: operator.pow, ast.Mod: operator.mod}
_CMP = {ast.Eq: operator.eq, ast.NotEq: operator.ne, ast.Lt: operator.lt,
        ast.LtE: operator.le, ast.Gt: operator.gt, ast.GtE: operator.ge}


def _safe_eval(expr: str):
    """Evaluate a tiny arithmetic/IF subset without eval(). Returns None if unsupported."""
    py = expr.replace("<>", "!=").replace("^", "**")
    py = re.sub(r"(?<![<>!=])=(?!=)", "==", py)
    try:
        tree = ast.parse(py.strip(), mode="eval")
    except SyntaxError:
        return None

    def ev(n):
        if isinstance(n, ast.Expression):
            return ev(n.body)
        if isinstance(n, ast.Constant) and isinstance(n.value, (int, float, str, bool)):
            return n.value
        if isinstance(n, ast.BinOp) and type(n.op) in _BIN:
            a, b = ev(n.left), ev(n.right)
            if isinstance(n.op, ast.Pow) and abs(b) > 64:
                raise ValueError
            return _BIN[type(n.op)](a, b)
        if isinstance(n, ast.UnaryOp) and isinstance(n.op, (ast.USub, ast.UAdd, ast.Not)):
            v = ev(n.operand)
            return -v if isinstance(n.op, ast.USub) else (v if isinstance(n.op, ast.UAdd) else (not v))
        if isinstance(n, ast.Compare) and len(n.ops) == 1 and type(n.ops[0]) in _CMP:
            return _CMP[type(n.ops[0])](ev(n.left), ev(n.comparators[0]))
        if isinstance(n, ast.Call) and isinstance(n.func, ast.Name):
            f = n.func.id.upper()
            if f == "IF":
                return ev(n.args[1]) if ev(n.args[0]) else (ev(n.args[2]) if len(n.args) > 2 else False)
            if f == "IFERROR":
                try:
                    return ev(n.args[0])
                except Exception:  # noqa: BLE001
                    return ev(n.args[1])
            a = [ev(x) for x in n.args]
            fn = {"ROUND": lambda: round(a[0], int(a[1]) if len(a) > 1 else 0), "ABS": lambda: abs(a[0]),
                  "MAX": lambda: max(a), "MIN": lambda: min(a), "SUM": lambda: sum(a),
                  "AVERAGE": lambda: sum(a) / len(a), "AND": lambda: all(a), "OR": lambda: any(a),
                  "NOT": lambda: not a[0]}.get(f)
            if fn:
                return fn()
        raise ValueError

    try:
        v = ev(tree)
    except Exception:  # noqa: BLE001
        return None
    if isinstance(v, bool):
        return float(v)
    return v if isinstance(v, (int, float)) else None


def _eval_formula_row(expr, vals, prev):
    if "{" in expr:
        return None

    def sub(m):
        name, pv = m.group(1), False
        if name.endswith(":prev"):
            name, pv = name[:-5], True
        v = (prev or {}).get(name, 0) if pv else vals.get(name)
        if v is None or v == "":
            v = 0
        if isinstance(v, (datetime, date)):
            raise ValueError
        return json.dumps(v) if isinstance(v, str) else repr(float(v))

    try:
        return _safe_eval(re.sub(r"\[([^\[\]]+)\]", sub, expr.lstrip("=")))
    except ValueError:
        return None


def compute_preview(spec) -> dict:
    """Approximate values computed in Python from the spec (Excel recalculates the real file)."""
    import pandas as pd
    spec, _ = normalize_spec(spec)
    tables, sheets = {}, {s["name"]: s for s in spec["sheets"]}
    for sh in spec["sheets"]:
        if sh["type"] != "table":
            continue
        cols = sh["columns"]
        recs, prev = [], None
        for row in sh["rows"]:
            vals = {c["name"]: _coerce(v, c["type"]) for c, v in zip(cols, row)}
            for _ in range(2):
                for c in cols:
                    if c["type"] == "formula":
                        vals[c["name"]] = _eval_formula_row(c["formula"], vals, prev)
            recs.append(vals)
            prev = vals
        tables[sh["name"]] = pd.DataFrame(recs, columns=[c["name"] for c in cols])
    kpis, charts = [], []
    for sh in spec["sheets"]:
        if sh["type"] != "dashboard":
            continue
        vals = {}
        for k in sh["kpis"]:
            if "expr" in k:
                continue
            df = tables[k["sheet"]]
            d = df
            w = k.get("where")
            if w:
                d = df[df[w["column"]].map(lambda x: _where_match(x, w))]
            s = pd.to_numeric(d[k["column"]], errors="coerce")
            v = {"sum": float(s.sum()), "average": float(s.mean()) if s.notna().any() else 0.0,
                 "count": int(d[k["column"]].notna().sum()),
                 "max": float(s.max()) if s.notna().any() else 0.0,
                 "min": float(s.min()) if s.notna().any() else 0.0}[k["agg"]]
            vals[k["label"]] = v
        for k in sh["kpis"]:
            if "expr" in k:
                try:
                    e = re.sub(r"\{([^{}]+)\}", lambda m: repr(float(vals[m.group(1)])), k["expr"].lstrip("="))
                    v = _safe_eval(e)
                except Exception:  # noqa: BLE001
                    v = None
                vals[k["label"]] = v
            kpis.append({"label": k["label"], "value": vals.get(k["label"]), "format": k["format"],
                         "dashboard": sh["name"]})
        for ch in sh["charts"]:
            s, df = sheets[ch["sheet"]], tables[ch["sheet"]]
            if ch["type"] == "scatter":
                x, y = ch["value_columns"]
                pts = df[[x, y]].apply(pd.to_numeric, errors="coerce").dropna()
                charts.append({"title": ch["title"], "type": "scatter", "data": pts, "x": x, "y": y})
                continue
            cats = chart_categories(s, ch)
            w = ch.get("where")
            out = pd.DataFrame(index=[str(_key(c)) if isinstance(c, (datetime, date)) else str(c) for c in cats])
            skipped = []
            for v in (ch["value_columns"] or ["Count"]):
                if ch["value_columns"] and df[v].isna().all() and len(df):
                    skipped.append(v)
                    continue
                col = []
                for cat in cats:
                    d = df[df[ch["category_column"]].map(lambda x: x is not None and _key(x) == _key(cat))]
                    if w:
                        d = d[d[w["column"]].map(lambda x: _where_match(x, w))]
                    if ch["agg"] == "count" or not ch["value_columns"]:
                        col.append(float(len(d)))
                    else:
                        n = pd.to_numeric(d[v], errors="coerce")
                        col.append({"sum": float(n.sum()), "average": float(n.mean()) if n.notna().any() else 0.0,
                                    "max": float(n.max()) if n.notna().any() else 0.0,
                                    "min": float(n.min()) if n.notna().any() else 0.0}.get(ch["agg"], float(n.sum())))
                out[v] = col
            charts.append({"title": ch["title"], "type": ch["type"], "data": out, "skipped": skipped})
    return {"tables": tables, "kpis": kpis, "charts": charts}


def format_value(v, fmt, settings) -> str:
    if v is None:
        return "n/a"
    cur = settings.get("currency_symbol", "$")
    try:
        if fmt == "currency":
            return f"{cur}{v:,.2f}"
        if fmt == "percent":
            return f"{v * 100:.1f}%"
        if fmt == "integer":
            return f"{v:,.0f}"
        return f"{v:,.2f}".rstrip("0").rstrip(".")
    except (TypeError, ValueError):
        return str(v)


def read_back(xlsx: bytes, sheet: str, max_rows=60, max_cols=18):
    """Read the generated file back from disk-format bytes for display (formulas shown as text)."""
    import pandas as pd
    wb = load_workbook(io.BytesIO(xlsx))
    ws = wb[sheet]
    data = [[("" if c.value is None else str(c.value)) for c in row]
            for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, max_rows), max_col=min(ws.max_column, max_cols))]
    return pd.DataFrame(data)


def formula_summary(xlsx: bytes) -> list[dict]:
    wb = load_workbook(io.BytesIO(xlsx))
    out = []
    for ws in wb.worksheets:
        seen, n = [], 0
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    n += 1
                    pat = re.sub(r"\$?[A-Z]{1,3}\$?\d+", "<cell>", c.value)
                    if pat not in [s[1] for s in seen] and len(seen) < 4:
                        seen.append((c.coordinate, pat, c.value))
        out.append({"sheet": ws.title, "formulas": n, "examples": [f"{a}: {c}" for a, _, c in seen],
                    "charts": len(ws._charts), "validations": len(ws.data_validations.dataValidation),
                    "conditional_formats": len(ws.conditional_formatting)})
    return out


# ---------------------------------------------------------------- QA checks
def _chart_refs(chart):
    refs = []
    for s in chart.series:
        for attr in ("val", "cat", "xVal", "yVal"):
            o = getattr(s, attr, None)
            if o is None:
                continue
            for sub in ("numRef", "strRef"):
                r = getattr(o, sub, None)
                if r is not None and r.f:
                    refs.append(r.f)
    return refs


def _recalc_libreoffice(xlsx: bytes):
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        return None
    try:
        with tempfile.TemporaryDirectory() as d:
            src = os.path.join(d, "in.xlsx")
            with open(src, "wb") as f:
                f.write(xlsx)
            out = os.path.join(d, "out")
            os.makedirs(out)
            subprocess.run([exe, "--headless", "--convert-to", "xlsx", "--outdir", out, src],
                           capture_output=True, timeout=120, check=False)
            p = os.path.join(out, "in.xlsx")
            if not os.path.exists(p):
                return None
            wb = load_workbook(p, data_only=True)
            return [f"{ws.title}!{c.coordinate}: {c.value}" for ws in wb.worksheets for row in ws.iter_rows()
                    for c in row if isinstance(c.value, str) and re.match(r"^#(REF|NAME|VALUE|DIV/0|N/A|NUM)", c.value)]
    except Exception:  # noqa: BLE001
        return None


def check_workbook(xlsx: bytes, spec: dict) -> list[dict]:
    """Real checks on the saved file. Each item: {status: pass|warn|fail|info, check, detail}."""
    R: list[dict] = []

    def add(status, check, detail=""):
        R.append({"status": status, "check": check, "detail": detail})

    if not xlsx or xlsx[:2] != b"PK":
        add("fail", "Output file exists and is a .xlsx (zip) container", "File is empty or not a zip archive.")
        return R
    try:
        wb = load_workbook(io.BytesIO(xlsx))
    except Exception as e:  # noqa: BLE001
        add("fail", "Workbook can be opened", str(e)[:200])
        return R
    add("pass", "Workbook opens with openpyxl", f"{len(wb.sheetnames)} sheets, {len(xlsx) / 1024:.0f} KB")

    missing = [s["name"] for s in spec["sheets"] if s["name"] not in wb.sheetnames]
    add("fail" if missing else "pass", "All specified sheets are present", ", ".join(missing))
    low = [n.lower() for n in wb.sheetnames]
    bad = [n for n in wb.sheetnames if len(n) > 31 or re.search(r"[\[\]:*?/\\]", n)]
    add("fail" if bad or len(set(low)) != len(low) else "pass", "Sheet names valid and unique", ", ".join(bad))

    names = set(wb.sheetnames)
    nform, errs, refs_bad, unbalanced = 0, [], [], []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    nform += 1
                    if "#REF!" in v:
                        errs.append(f"{ws.title}!{c.coordinate}")
                    if v.count("(") != v.count(")"):
                        unbalanced.append(f"{ws.title}!{c.coordinate}")
                    for a, b in re.findall(r"(?:'((?:[^']|'')+)'|\b([A-Za-z_][\w]*))!", v):
                        nm = (a or b).replace("''", "'")
                        if nm not in names:
                            refs_bad.append(f"{ws.title}!{c.coordinate} -> {nm}")
    add("pass" if nform else "warn", "Formulas are populated", f"{nform} formula cells found")
    add("fail" if errs or unbalanced else "pass", "No #REF! / unbalanced formulas", ", ".join((errs + unbalanced)[:6]))
    add("fail" if refs_bad else "pass", "Cross-sheet references point to existing sheets", "; ".join(refs_bad[:5]))

    nch, chart_bad = 0, []
    for ws in wb.worksheets:
        for ch in ws._charts:
            nch += 1
            rs = _chart_refs(ch)
            if not rs:
                chart_bad.append(f"{ws.title}: chart without data references")
            for r in rs:
                m = re.match(r"^'?(.*?)'?!(.+)$", r)
                if not m or m.group(1).replace("''", "'") not in names:
                    chart_bad.append(f"{ws.title}: {r}")
    expected = sum(len(s.get("charts", [])) for s in spec["sheets"] if s["type"] == "dashboard")
    add("fail" if chart_bad else "pass", "Chart source ranges are valid", "; ".join(chart_bad[:4]) or f"{nch} native chart(s)")
    add("warn" if nch < expected else "pass", "All requested charts were created", f"{nch} of {expected}")

    ndv = sum(len(ws.data_validations.dataValidation) for ws in wb.worksheets)
    ncf = sum(len(ws.conditional_formatting) for ws in wb.worksheets)
    lists_bad = []
    for ws in wb.worksheets:
        for dv in ws.data_validations.dataValidation:
            if dv.type == "list" and dv.formula1 and dv.formula1.startswith("=Lists!") and "Lists" not in names:
                lists_bad.append(ws.title)
    add("fail" if lists_bad else "pass", "Dropdown source ranges exist", f"{ndv} validation rule(s), {ncf} conditional-format range(s)")

    for s in spec["sheets"]:
        if s["type"] == "table" and s["sample_data"] and s["rows"]:
            ws = wb[s["name"]] if s["name"] in names else None
            if ws is not None and "SAMPLE DATA" not in str(ws["A2"].value):
                add("warn", f'Sample data on "{s["name"]}" is labelled', "No sample-data notice found")
    if any(s["type"] == "table" and s["sample_data"] and s["rows"] for s in spec["sheets"]):
        add("info", "Sample data is clearly labelled", "Tables with sample rows carry a 'SAMPLE DATA' notice in row 2.")
    empty_tables = [s["name"] for s in spec["sheets"] if s["type"] == "table" and not s["columns"]]
    add("fail" if empty_tables else "pass", "Every table sheet has columns", ", ".join(empty_tables))

    try:
        pv = compute_preview(spec)
        evaluated = 0
        for sh in spec["sheets"]:
            if sh["type"] == "table":
                for c in sh["columns"]:
                    if c["type"] == "formula" and sh["rows"] and pv["tables"][sh["name"]][c["name"]].notna().any():
                        evaluated += 1
        add("info", "Python re-calculation of sample formulas", f"{evaluated} formula column(s) evaluated on sample rows; "
            f"{sum(1 for k in pv['kpis'] if k['value'] is not None)} KPI(s) computed. Functions beyond basic arithmetic/IF are left to Excel.")
    except Exception as e:  # noqa: BLE001
        add("warn", "Python re-calculation of sample formulas", f"Could not run: {e}")

    try:
        buf = io.BytesIO()
        wb.save(buf)
        load_workbook(io.BytesIO(buf.getvalue()))
        add("pass", "Workbook re-saves and re-opens", "")
    except Exception as e:  # noqa: BLE001
        add("fail", "Workbook re-saves and re-opens", str(e)[:200])

    lo = _recalc_libreoffice(xlsx)
    if lo is None:
        add("info", "Independent formula calculation (LibreOffice)",
            "Not available in this environment. Excel recalculates formulas when the file opens (fullCalcOnLoad is set), "
            "but formula results were not independently verified here.")
    else:
        add("fail" if lo else "pass", "LibreOffice recalculation found no formula errors", "; ".join(lo[:5]))
    return R


def etsy_checks(xlsx: bytes, spec: dict) -> list[dict]:
    R = check_workbook(xlsx, spec)
    wb = load_workbook(io.BytesIO(xlsx))
    has_ins = any(n.lower() == "instructions" for n in wb.sheetnames)
    R.append({"status": "pass" if has_ins else "fail", "check": "Instructions sheet included", "detail": ""})
    for s in spec["sheets"]:
        if s["type"] == "table" and not (s["title"] and (s["notes"] or s["purpose"])):
            R.append({"status": "warn", "check": f'"{s["name"]}" has a title and usage note', "detail": "Add a purpose/notes via the editor."})
    inputs = [s["name"] for s in spec["sheets"] if s["type"] == "table" and s["extra_rows"] == 0]
    if inputs:
        R.append({"status": "warn", "check": "Room for customer data", "detail": "No spare rows on: " + ", ".join(inputs)})
    R.append({"status": "info", "check": "Compatibility",
              "detail": "Built with standard features (formulas, validation, native charts). MAXIFS/MINIFS/IFERROR need Excel 2019+/365. "
                        "Not tested in desktop Excel, Excel for Mac, or Google Sheets from this app."})
    R.append({"status": "info", "check": "Listing preview images",
              "detail": "Not generated by this app. Open the workbook and take screenshots of each sheet for your listing."})
    R.append({"status": "info", "check": "Originality",
              "detail": "Use references for layout inspiration only; do not copy a competitor's branding, text or artwork."})
    return R


# ----------------------------------------------------------------- PDF guide
def build_pdf_guide(spec: dict, filename: str) -> bytes:
    from xml.sax.saxutils import escape
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    T = _T(spec)
    st = getSampleStyleSheet()
    h1 = ParagraphStyle("h1", parent=st["Title"], textColor=colors.HexColor("#" + T["primary"]))
    h2 = ParagraphStyle("h2", parent=st["Heading2"], textColor=colors.HexColor("#" + T["primary"]))
    body = st["BodyText"]
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm, topMargin=2 * cm, bottomMargin=2 * cm,
                            title=f"{spec['workbook_name']} - Setup Guide")
    P = lambda t, s=body: Paragraph(escape(t), s)  # noqa: E731
    story = [P(spec["workbook_name"], h1), P("Setup and usage guide", st["Heading3"]), Spacer(1, 12),
             P(f"File: {filename}"), Spacer(1, 6),
             P("Open the workbook in Microsoft Excel (2019 or Microsoft 365 recommended). If Excel opens it in Protected View, click "
               "Enable Editing. Formulas recalculate automatically; press F9 if values look out of date."), Spacer(1, 10),
             P("Quick start", h2)]
    for i, t in enumerate(["Start on the Dashboard sheet to see your summary.",
                           "Go to each data sheet and replace the sample rows with your own information.",
                           "Use the drop-down lists where an arrow appears in a cell.",
                           "Do not type over formula columns; they calculate for you.",
                           "The dashboard and charts update when your data changes."], 1):
        story.append(P(f"{i}. {t}"))
    story += [Spacer(1, 10), P("Worksheets", h2)]
    rows = [["Sheet", "Purpose"]] + [[Paragraph(escape(s["name"]), body), Paragraph(escape(s["purpose"] or s["title"]), body)]
                                      for s in spec["sheets"]]
    tb = Table(rows, colWidths=[4.5 * cm, 12.5 * cm], repeatRows=1)
    tb.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#" + T["primary"])),
                            ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#" + T["on_primary"])),
                            ("GRID", (0, 0), (-1, -1), 0.3, colors.grey), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(tb)
    for s in spec["sheets"]:
        if s["type"] != "table":
            continue
        story += [Spacer(1, 12), P(s["name"], h2)]
        if s["notes"]:
            story.append(P(s["notes"]))
        for c in s["columns"]:
            d = {"dropdown": "choose from the list: " + ", ".join(c.get("options", [])[:8]),
                 "formula": "calculated automatically (do not edit)", "date": "enter a date",
                 "currency": "enter an amount", "number": "enter a number", "percent": "enter a percentage", "text": "free text"}[c["type"]]
            story.append(P(f"- {c['name']}: {d}"))
    story += [Spacer(1, 14), P("Troubleshooting", h2),
              P("Formulas show 0 or blank: check that data is entered in the correct columns and that Excel calculation is set to Automatic."),
              P("A drop-down refuses my value: pick from the list, or edit the list source if you need new options."),
              P("Charts look empty: the chart data tables on the Dashboard (columns T onward) list the categories being plotted; add new category names there."),
              Spacer(1, 10), P("Sample data in this workbook is for demonstration only and is not real data.")]
    doc.build(story)
    return buf.getvalue()


# ---------------------------------------------------- templates and example
TEMPLATES = {
    "Personal productivity": {
        "All-in-one life planner": "An all-in-one life planner with dashboard, goals, habits, weekly tasks, and monthly overview.",
        "Annual planner": "An annual planner with a yearly overview, monthly goals and key dates.",
        "Monthly and weekly planner": "A monthly and weekly planner with a task list, priorities and a weekly schedule.",
        "Daily task tracker": "A daily task tracker with priority, status, due date and a completion dashboard.",
        "Goal tracker": "A goal tracker with target, progress percentage, deadline and status with progress bars.",
        "Habit tracker": "A habit tracker with daily yes/no check-ins per habit and a completion dashboard.",
        "Time-blocking planner": "A time-blocking planner with hourly blocks, categories and a time-by-category summary.",
        "Student planner": "A student planner with courses, assignments, due dates, grades and a dashboard.",
    },
    "Fitness and wellness": {
        "Annual workout tracker": "An annual workout tracker with exercise log, sets, reps, weight and monthly progress charts.",
        "Exercise log": "An exercise log with date, exercise, sets, reps, weight and volume calculation.",
        "Fitness progress dashboard": "A fitness progress dashboard with body measurements, workouts per week and trend charts.",
        "Weight tracker (kg)": "A weight tracker in kilograms with date, weight, change, goal and a trend chart.",
        "Weight tracker (lb)": "A weight tracker in pounds with date, weight, change, goal and a trend chart.",
        "Hydration tracker": "A hydration tracker with daily water intake against a goal and a weekly chart.",
        "Meal planner": "A weekly meal planner with breakfast, lunch, dinner, snacks and calories.",
        "Grocery list": "A grocery list with category, quantity, price, purchased status and a total cost.",
    },
    "Finance": {
        "Personal budget dashboard": "A personal budget dashboard with income, expenses by category, savings and monthly trend.",
        "Monthly expense tracker": "A monthly expense tracker with date, category, amount, payment method and charts.",
        "Savings tracker": "A savings tracker with goals, deposits, balance, progress percentage and a chart.",
        "Debt repayment tracker": "A debt repayment tracker with balances, interest rates, payments, remaining balance and payoff progress.",
        "Financial goal tracker": "A financial goal tracker with target amount, saved amount, deadline and progress.",
        "Income and expense dashboard": "An income and expense dashboard with monthly totals, categories and net savings.",
    },
    "Business": {
        "Small business profit and expense dashboard": "A small business profit and expense dashboard with revenue, costs, profit margin and monthly charts.",
        "Sales analytics dashboard": "A sales analytics dashboard with sales by product, region, month and top KPIs.",
        "Inventory management system": "An inventory management system with stock levels, reorder alerts, valuation and suppliers.",
        "Restaurant inventory and food cost calculator": "A restaurant inventory and food cost calculator with ingredients, recipes, cost per portion and margin.",
        "Product pricing calculator": "A product pricing calculator with cost, overhead, margin, price and profit per unit.",
        "Supplier tracker": "A supplier tracker with contacts, products, lead times, ratings and orders.",
        "Project and task management dashboard": "A project and task management dashboard with owners, status, deadlines and progress charts.",
    },
}

EXAMPLE_SPEC = {
    "workbook_name": "Personal Budget Dashboard",
    "theme": {"name": "Sage green"},
    "settings": {"currency_symbol": "$", "add_instructions": True},
    "assumptions": ["Offline example: all figures are sample data.", "Currency is US dollars."],
    "sheets": [
        {"name": "Dashboard", "type": "dashboard", "title": "Personal Budget Dashboard",
         "purpose": "Income, spending and savings at a glance. Sample data shown.",
         "kpis": [
             {"label": "Total Income", "sheet": "Transactions", "column": "Amount", "agg": "sum", "format": "currency",
              "where": {"column": "Type", "equals": "Income"}},
             {"label": "Total Expenses", "sheet": "Transactions", "column": "Amount", "agg": "sum", "format": "currency",
              "where": {"column": "Type", "equals": "Expense"}},
             {"label": "Net Savings", "expr": "{Total Income}-{Total Expenses}", "format": "currency"},
             {"label": "Savings Rate", "expr": "IFERROR(({Total Income}-{Total Expenses})/{Total Income},0)", "format": "percent"}],
         "charts": [
             {"type": "doughnut", "title": "Spending by Category", "sheet": "Transactions", "category_column": "Category",
              "value_columns": ["Amount"], "agg": "sum", "where": {"column": "Type", "equals": "Expense"}},
             {"type": "column", "title": "Expenses by Month", "sheet": "Transactions", "category_column": "Month",
              "value_columns": ["Amount"], "agg": "sum", "where": {"column": "Type", "equals": "Expense"}},
             {"type": "bar", "title": "Budget vs Spent", "sheet": "Budget", "category_column": "Category",
              "value_columns": ["Monthly Budget", "Spent"], "agg": "sum"}]},
        {"name": "Transactions", "type": "table", "title": "Transactions",
         "purpose": "Log every income and expense here.", "notes": "Enter one row per transaction. Choose Month, Type and Category from the lists.",
         "extra_rows": 200,
         "columns": [{"name": "Date", "type": "date"},
                     {"name": "Month", "type": "dropdown", "options": ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]},
                     {"name": "Type", "type": "dropdown", "options": ["Income", "Expense"]},
                     {"name": "Category", "type": "dropdown", "options": ["Salary", "Rent", "Groceries", "Transport", "Utilities", "Dining", "Health", "Fun"]},
                     {"name": "Description", "type": "text"},
                     {"name": "Amount", "type": "currency", "min": 0},
                     {"name": "Running Balance", "type": "formula", "format": "currency",
                      "formula": "[Running Balance:prev]+IF([Type]=\"Income\",[Amount],-[Amount])"}],
         "rows": [
             ["2025-01-01", "Jan", "Income", "Salary", "Monthly salary", 3200, None],
             ["2025-01-02", "Jan", "Expense", "Rent", "January rent", 1100, None],
             ["2025-01-05", "Jan", "Expense", "Groceries", "Weekly shop", 96.4, None],
             ["2025-01-09", "Jan", "Expense", "Transport", "Bus pass", 45, None],
             ["2025-01-14", "Jan", "Expense", "Utilities", "Electricity", 72.5, None],
             ["2025-01-20", "Jan", "Expense", "Dining", "Dinner out", 58, None],
             ["2025-02-01", "Feb", "Income", "Salary", "Monthly salary", 3200, None],
             ["2025-02-02", "Feb", "Expense", "Rent", "February rent", 1100, None],
             ["2025-02-06", "Feb", "Expense", "Groceries", "Weekly shop", 104.2, None],
             ["2025-02-11", "Feb", "Expense", "Health", "Pharmacy", 31.9, None],
             ["2025-02-16", "Feb", "Expense", "Fun", "Cinema", 24, None],
             ["2025-02-21", "Feb", "Expense", "Utilities", "Internet", 49, None],
             ["2025-03-01", "Mar", "Income", "Salary", "Monthly salary", 3200, None],
             ["2025-03-02", "Mar", "Expense", "Rent", "March rent", 1100, None],
             ["2025-03-07", "Mar", "Expense", "Groceries", "Weekly shop", 110.8, None],
             ["2025-03-12", "Mar", "Expense", "Transport", "Fuel", 60, None],
             ["2025-03-19", "Mar", "Expense", "Dining", "Lunch with friends", 42.5, None]],
         "conditional": [{"column": "Amount", "rule": "data_bar", "color": "blue"}]},
        {"name": "Budget", "type": "table", "title": "Monthly Budget by Category",
         "purpose": "Set a monthly budget per category and compare with what was spent.",
         "notes": "Edit the Monthly Budget column. Spent adds up all matching expense rows.", "extra_rows": 20,
         "columns": [{"name": "Category", "type": "text"},
                     {"name": "Monthly Budget", "type": "currency", "min": 0},
                     {"name": "Spent", "type": "formula", "format": "currency",
                      "formula": "SUMIFS({Transactions.Amount},{Transactions.Category},[Category],{Transactions.Type},\"Expense\")"},
                     {"name": "Remaining", "type": "formula", "format": "currency", "formula": "[Monthly Budget]-[Spent]"},
                     {"name": "Used", "type": "formula", "format": "percent", "formula": "IFERROR([Spent]/[Monthly Budget],0)"}],
         "rows": [["Rent", 1100, None, None, None], ["Groceries", 300, None, None, None], ["Transport", 150, None, None, None],
                  ["Utilities", 160, None, None, None], ["Dining", 120, None, None, None], ["Health", 80, None, None, None],
                  ["Fun", 90, None, None, None]],
         "conditional": [{"column": "Remaining", "rule": "less_than", "value": 0, "color": "red"},
                         {"column": "Used", "rule": "progress", "color": "green"}]},
    ],
}
