"""AI Excel Studio - Streamlit interface. Run with:  streamlit run app.py"""
from __future__ import annotations

import copy
import io
import json
import zipfile
from datetime import datetime

import altair as alt
import pandas as pd
import streamlit as st

import engine as E

st.set_page_config(page_title="AI Excel Studio", page_icon="📊", layout="wide")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
ICON = {"pass": "✅", "warn": "⚠️", "fail": "❌", "info": "ℹ️"}
PROJECT_TYPES = [
    "Business dashboard", "Financial dashboard", "Personal life planner", "Budget and expense tracker",
    "Goal and habit tracker", "Fitness and workout tracker", "Meal planner and grocery tracker",
    "Inventory and stock management", "Sales analytics", "Business operations", "Student planner",
    "Project management", "HR and employee tracker", "Custom Excel template"]
NAV = ["New Project", "Build Plan", "Workspace", "Templates", "Version History", "Settings & Help"]

ss = st.session_state
for k, v in {"project": None, "plan": None, "design": None, "imported": [], "plan_warnings": [],
             "pending": None, "goto": None, "prefill": "", "desc_ver": 0, "listing": None,
             "transcript": None, "clear_edit": False, "nav": "New Project"}.items():
    ss.setdefault(k, v)

st.markdown("""<style>
.block-container{padding-top:1.5rem;max-width:1250px}
div[data-testid="stMetricValue"]{font-size:1.7rem}
</style>""", unsafe_allow_html=True)


# ------------------------------------------------------------------ helpers
def api_ready() -> bool:
    return bool(E.get_api_key())


def cur() -> dict:
    return ss.project["versions"][-1]


@st.cache_data(show_spinner=False)
def _import_cached(name: str, data: bytes):
    return E.import_dataset(name, data)


def store_version(spec: dict, note: str) -> bool:
    """Build a REAL workbook, verify it, and save it as a new version. Returns True only if a file exists."""
    try:
        data, spec, warns = E.build_workbook(spec)
    except (E.SpecError, Exception) as e:  # noqa: BLE001
        st.error(f"Workbook generation failed: {e}")
        return False
    checks = E.check_workbook(data, spec)
    if checks and checks[0]["status"] == "fail":
        st.error("The generated file could not be opened, so it was not saved: " + checks[0]["detail"])
        return False
    p = ss.project or {"versions": [], "chat": [], "design": ss.design}
    p["versions"].append({"n": len(p["versions"]) + 1, "time": datetime.now().strftime("%H:%M:%S"),
                          "note": note, "spec": spec, "bytes": data, "warnings": warns, "checks": checks})
    ss.project = p
    return True


def feature_summary(spec: dict, checks: list) -> str:
    L = [f"# {spec['workbook_name']}", "", "## Worksheets"]
    for s in spec["sheets"]:
        L.append(f"- **{s['name']}** ({s['type']}): {s.get('purpose') or s.get('title', '')}")
        if s["type"] == "table":
            L.append("  - Columns: " + ", ".join(f"{c['name']} ({c['type']})" for c in s["columns"]))
        if s["type"] == "dashboard":
            L.append(f"  - {len(s['kpis'])} KPI card(s), {len(s['charts'])} native chart(s)")
    L += ["", f"Theme: {spec['theme']['name']} ({spec['theme']['font']})", "", "## Quality checks"]
    L += [f"- [{c['status'].upper()}] {c['check']}" + (f": {c['detail']}" if c["detail"] else "") for c in checks]
    if spec.get("assumptions"):
        L += ["", "## Assumptions"] + [f"- {a}" for a in spec["assumptions"]]
    return "\n".join(L)


def render_checks(checks):
    for c in checks:
        st.markdown(f"{ICON.get(c['status'], '•')} **{c['check']}**" + (f" — {c['detail']}" if c["detail"] else ""))


def render_chart(ch: dict):
    st.markdown(f"**{ch['title']}**")
    d = ch["data"]
    if ch["type"] == "scatter":
        st.scatter_chart(d, x=ch["x"], y=ch["y"])
        return
    if d.shape[1] == 0 or d.empty:
        st.info("Preview unavailable: the values come from formulas Excel must calculate.")
        return
    d = d.copy()
    d.index.name = "category"
    if ch["type"] in ("pie", "doughnut"):
        df = pd.DataFrame({"category": d.index.astype(str), "value": d.iloc[:, 0].values})
        st.altair_chart(alt.Chart(df).mark_arc(innerRadius=60 if ch["type"] == "doughnut" else 0)
                        .encode(theta="value:Q", color="category:N", tooltip=["category", "value"]),
                        use_container_width=True)
        return
    long = d.reset_index().melt(id_vars="category", var_name="series", value_name="value")
    base = alt.Chart(long)
    if ch["type"] == "line":
        c = base.mark_line(point=True).encode(x=alt.X("category:N", sort=None), y="value:Q", color="series:N")
    elif ch["type"] == "area":
        c = base.mark_area(opacity=0.6).encode(x=alt.X("category:N", sort=None), y="value:Q", color="series:N")
    elif ch["type"] == "bar":
        c = base.mark_bar().encode(y=alt.Y("category:N", sort=None), x="value:Q", color="series:N")
    else:
        c = base.mark_bar().encode(x=alt.X("category:N", sort=None), y="value:Q", color="series:N")
    st.altair_chart(c, use_container_width=True)
    if ch.get("skipped"):
        st.caption("Not shown in this preview (calculated by Excel): " + ", ".join(ch["skipped"]))


# --------------------------------------------------------------- edit engine
def commit_ops(ops, summary, instr, confirmed=False):
    p, v = ss.project, cur()
    new, log, pend, warns = E.apply_operations(v["spec"], ops, confirmed)
    if pend:
        ss.pending = {"ops": ops, "summary": summary, "instr": instr, "items": pend}
        p["chat"].append(("assistant", "This change would remove content, so I need your confirmation first:\n\n"
                          + "\n".join(f"- {x}" for x in pend)))
        return False
    if not any(not l.startswith("Skipped") for l in log):
        p["chat"].append(("assistant", "No change was applied:\n\n" + "\n".join(f"- {l}" for l in log)))
        return False
    if not store_version(new, summary or instr[:80]):
        p["chat"].append(("assistant", "The change produced an invalid workbook and was not saved. Your previous version is unchanged."))
        return False
    issues = E.preservation_report(v["spec"], cur()["spec"])
    msg = f"**{summary or 'Changes applied'}**\n\n" + "\n".join(f"- {l}" for l in log)
    if warns:
        msg += "\n\nNotes:\n" + "\n".join(f"- {w}" for w in warns)
    msg += ("\n\nExisting content that changed:\n" + "\n".join(f"- {i}" for i in issues)) if issues else \
        "\n\nNo existing sheets, columns or data rows were removed."
    msg += f"\n\nSaved as version {cur()['n']}."
    p["chat"].append(("assistant", msg))
    return True


def run_edit(instr: str):
    p = ss.project
    p["chat"].append(("user", instr))
    with st.spinner("Planning the change..."):
        try:
            res = E.ai_edit(cur()["spec"], instr, p.get("design"))
        except E.AIError as e:
            p["chat"].append(("assistant", f"⚠️ {e}"))
            return
    if res["clarification"]:
        p["chat"].append(("assistant", res["clarification"]))
    elif not res["operations"]:
        p["chat"].append(("assistant", "I couldn't turn that into a safe change. Could you rephrase or be more specific?"))
    else:
        commit_ops(res["operations"], res["summary"], instr)


# -------------------------------------------------------------------- pages
def page_new():
    st.title("📊 AI Excel Studio")
    st.caption("Describe an Excel product, add reference images or data, and get a real, editable .xlsx workbook.")
    if not api_ready():
        st.warning("No Gemini API key found. Add `GEMINI_API_KEY` to your Streamlit secrets (see README). "
                   "You can still build the offline example below.")
    with st.expander("Project instructions"):
        st.markdown("1. Pick a project type and describe what you need.\n2. Optionally add reference images (layout/colours) and/or a CSV/Excel dataset.\n"
                    "3. Review the AI build plan, change it if you like, then approve it.\n4. Edit the workbook by typing or dictating requests, then download the .xlsx.\n\n"
                    "Projects live in this browser session. Download the project file (.json) from the Download tab to continue later.")
    c1, c2 = st.columns([3, 2], gap="large")
    with c1:
        ptype = st.selectbox("Project type", PROJECT_TYPES)
        desc = st.text_area("Describe your Excel product", value=ss.prefill, key=f"desc_{ss.desc_ver}", height=150,
                            placeholder="e.g. A sage-green and cream life planner with a dashboard, monthly budget, goals and a habit tracker.")
        theme = st.selectbox("Design theme", ["Let the AI choose"] + list(E.THEMES), help="Reference images override this.")
    with c2:
        imgs = st.file_uploader("Reference images (PNG/JPG)", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
        data_file = st.file_uploader("Import a dataset or existing workbook (.xlsx / .csv)", type=["xlsx", "csv"])
        proj_file = st.file_uploader("Continue a previous project (.json)", type=["json"])
    if data_file is not None:
        try:
            sheets, warns = _import_cached(data_file.name, data_file.getvalue())
        except E.SpecError as e:
            st.error(str(e))
            sheets, warns = [], []
        for w in warns:
            st.warning(w)
        keep = []
        for sh in sheets:
            prof = E.profile_sheet(sh)
            with st.expander(f"Use '{sh['name']}' - {prof['rows']} rows, {len(sh['columns'])} columns", expanded=len(sheets) == 1):
                if st.checkbox("Include this sheet", value=True, key=f"inc_{sh['name']}"):
                    keep.append(sh)
                st.dataframe(pd.DataFrame(sh["rows"][:8], columns=[c["name"] for c in sh["columns"]]), use_container_width=True)
                st.caption(f"Missing values: {prof['missing'] or 'none'} | duplicate rows: {prof['duplicates']} | "
                           "types: " + ", ".join(f"{k}={v}" for k, v in prof["types"].items()))
        ss.imported = keep
        st.caption("Your original file is never modified; a new workbook is generated separately.")
    else:
        ss.imported = []

    b1, b2, b3 = st.columns(3)
    if b1.button("Create build plan", type="primary", disabled=not (desc.strip() and api_ready())):
        images = [(f.name, f.getvalue(), f.type or "image/png") for f in imgs]
        design = None
        with st.status("Working on your plan...", expanded=True) as status:
            try:
                if images:
                    status.write(f"Analysing {len(images)} reference image(s)...")
                    design = E.analyze_images(images, desc)
                status.write("Designing the workbook...")
                raw = E.plan_workbook(ptype, desc, design, E.dataset_summary(ss.imported) if ss.imported else None, theme, images)
                if ss.imported:
                    raw = E.merge_imported(raw, ss.imported)
                spec, warns = E.normalize_spec(raw)
                status.update(label="Plan ready", state="complete")
            except (E.AIError, E.SpecError) as e:
                status.update(label="Could not create the plan", state="error")
                st.error(str(e))
                spec = None
            except Exception as e:  # noqa: BLE001
                status.update(label="Could not create the plan", state="error")
                st.error(f"Unexpected problem while planning: {e}")
                spec = None
        if spec:
            ss.plan, ss.design, ss.plan_warnings = spec, design, warns
            ss.goto = "Build Plan"
            st.rerun()
    if b2.button("Build offline example (no AI)"):
        ss.project = None
        if store_version(copy.deepcopy(E.EXAMPLE_SPEC), "Offline example"):
            ss.goto = "Workspace"
            st.rerun()
    if proj_file is not None and b3.button("Load project file"):
        try:
            blob = json.loads(proj_file.getvalue())
            ss.project = None
            ss.design = blob.get("design")
            if store_version(blob["spec"], "Loaded from project file"):
                ss.project["design"] = ss.design
                ss.goto = "Workspace"
                st.rerun()
        except (json.JSONDecodeError, KeyError, TypeError) as e:
            st.error(f"Not a valid AI Excel Studio project file: {e}")


def page_plan():
    spec = ss.plan
    if not spec:
        st.info("No plan yet. Start on the New Project page.")
        return
    st.header(f"Build plan: {spec['workbook_name']}")
    for w in ss.plan_warnings:
        st.warning(w)
    c1, c2, c3 = st.columns(3)
    c1.metric("Worksheets", len(spec["sheets"]))
    c2.metric("Theme", spec["theme"]["name"])
    c3.metric("Charts", sum(len(s.get("charts", [])) for s in spec["sheets"]))
    st.dataframe(pd.DataFrame([{"Sheet": s["name"], "Type": s["type"], "Purpose": s.get("purpose", "")} for s in spec["sheets"]]),
                 use_container_width=True, hide_index=True)
    for s in spec["sheets"]:
        with st.expander(f"{s['name']} ({s['type']})"):
            if s["type"] == "table":
                st.dataframe(pd.DataFrame([{"Field": c["name"], "Type": c["type"],
                                            "Formula / options": c.get("formula") or ", ".join(c.get("options", []))}
                                           for c in s["columns"]]), use_container_width=True, hide_index=True)
                st.caption(f"{len(s['rows'])} {'sample ' if s['sample_data'] else ''}rows, {s['extra_rows']} spare rows. "
                           f"Conditional rules: {len(s['conditional'])}")
            elif s["type"] == "dashboard":
                st.write("KPIs: " + (", ".join(k["label"] for k in s["kpis"]) or "none"))
                st.write("Charts: " + (", ".join(f"{c['title']} ({c['type']})" for c in s["charts"]) or "none"))
            else:
                st.write("\n".join(s.get("lines", [])) or "Auto-generated usage guide.")
    if spec.get("assumptions"):
        st.subheader("Assumptions (editable via the box below)")
        st.markdown("\n".join(f"- {a}" for a in spec["assumptions"]))
    if ss.design:
        with st.expander("Design specification extracted from your reference images"):
            st.json(ss.design)
    st.divider()
    change = st.text_input("Change the plan", placeholder="e.g. Add a debt repayment sheet and use dark blue colours")
    if st.button("Update plan", disabled=not change.strip()):
        with st.spinner("Updating plan..."):
            try:
                res = E.ai_edit(spec, change, ss.design)
                if res["clarification"]:
                    st.info(res["clarification"])
                else:
                    new, log, _, warns = E.apply_operations(spec, res["operations"], confirmed=True)
                    ss.plan, ss.plan_warnings = new, warns
                    st.rerun()
            except (E.AIError, E.SpecError) as e:
                st.error(str(e))
    if st.button("✅ Approve and build workbook", type="primary"):
        ss.project = None
        with st.spinner("Building your workbook..."):
            ok = store_version(ss.plan, "Initial build")
        if ok:
            ss.project["design"] = ss.design
            ss.goto = "Workspace"
            st.rerun()


def tab_preview(v):
    spec = v["spec"]
    try:
        pv = E.compute_preview(spec)
    except Exception as e:  # noqa: BLE001
        pv = None
        st.error(f"Preview calculation failed: {e}")
    st.caption("KPI and chart previews are calculated in Python from the same spec used to build the file. "
               "Excel recalculates the real workbook when you open it, so tiny differences are possible.")
    if pv:
        for sh in [s for s in spec["sheets"] if s["type"] == "dashboard"]:
            st.subheader(sh["title"] or sh["name"])
            ks = [k for k in pv["kpis"] if k["dashboard"] == sh["name"]]
            if ks:
                cols = st.columns(min(4, len(ks)))
                for i, k in enumerate(ks):
                    cols[i % len(cols)].metric(k["label"], E.format_value(k["value"], k["format"], spec["settings"]))
            titles = {c["title"] for c in sh["charts"]}
            cs = [c for c in pv["charts"] if c["title"] in titles]
            for i in range(0, len(cs), 2):
                cc = st.columns(2)
                for col, ch in zip(cc, cs[i:i + 2]):
                    with col:
                        render_chart(ch)
    st.subheader("Generated file contents (read back from the .xlsx)")
    fs = E.formula_summary(v["bytes"])
    sheet = st.selectbox("Worksheet", [f["sheet"] for f in fs])
    st.dataframe(E.read_back(v["bytes"], sheet), use_container_width=True, hide_index=True, height=340)
    with st.expander("Formula and feature summary"):
        st.dataframe(pd.DataFrame([{"Sheet": f["sheet"], "Formulas": f["formulas"], "Charts": f["charts"],
                                    "Validations": f["validations"], "Conditional formats": f["conditional_formats"],
                                    "Example formulas": " | ".join(f["examples"])[:300]} for f in fs]),
                     use_container_width=True, hide_index=True)
    for w in v["warnings"]:
        st.warning(w)


def tab_edit(v):
    p = ss.project
    if ss.clear_edit or ss.transcript is not None:
        ss["edit_text"] = ss.transcript if ss.transcript is not None else ""
        ss.transcript, ss.clear_edit = None, False
    for role, text in p["chat"][-12:]:
        with st.chat_message(role):
            st.markdown(text)
    if ss.pending:
        st.warning("Confirmation needed:\n\n" + "\n".join(f"- {x}" for x in ss.pending["items"]))
        a, b = st.columns(2)
        if a.button("Yes, apply these changes", type="primary"):
            pend, ss.pending = ss.pending, None
            commit_ops(pend["ops"], pend["summary"], pend["instr"], confirmed=True)
            st.rerun()
        if b.button("Cancel"):
            ss.pending = None
            p["chat"].append(("assistant", "Cancelled. Nothing was changed."))
            st.rerun()
    txt = st.text_area("Edit your workbook", key="edit_text", height=100,
                       placeholder="Describe the changes you want to make to your Excel workbook...")
    c1, c2 = st.columns([1, 3])
    if c1.button("Apply changes", type="primary", disabled=not (txt.strip() and api_ready())):
        run_edit(txt.strip())
        ss.clear_edit = True
        st.rerun()
    c2.caption("Examples: change the theme to sage green and cream · add a monthly revenue chart · add a debt repayment sheet · "
               "add a dropdown for month · increase rows for future entries")
    with st.expander("🎤 Voice input"):
        if hasattr(st, "audio_input"):
            audio = st.audio_input("Record a command")
            if audio is not None and st.button("Transcribe recording"):
                try:
                    with st.spinner("Transcribing..."):
                        text = E.transcribe(audio.getvalue(), audio.type or "audio/wav")
                    ss.transcript = text
                    st.rerun()
                except E.AIError as e:
                    st.error(str(e))
            st.caption("The transcript appears in the box above so you can correct it before applying. Needs microphone permission "
                       "in your browser. If recording is unavailable, just type your request.")
        else:
            st.info("Voice recording is not supported by this Streamlit version or browser. Please type your request.")
    st.caption(f"Current version: {v['n']} - every edit is saved as a new version you can restore.")


def tab_quality(v):
    if st.button("Re-run checks"):
        v["checks"] = E.check_workbook(v["bytes"], v["spec"])
    render_checks(v["checks"])
    st.caption("These checks inspect the saved file. They do not prove every formula calculates correctly in every Excel version.")


def tab_theme(v):
    spec = v["spec"]
    t = spec["theme"]
    preset = st.selectbox("Preset", ["(current)"] + list(E.THEMES))
    src = t if preset == "(current)" else E.norm_theme(preset)
    cols = st.columns(5)
    colors = {k: cols[i].color_picker(k.title(), src[k], key=f"cp_{preset}_{k}")
              for i, k in enumerate(("primary", "accent", "background", "surface", "text"))}
    font = st.selectbox("Font", E.FONTS, index=E.FONTS.index(t["font"]) if t["font"] in E.FONTS else 0)
    if st.button("Apply theme", type="primary"):
        commit_ops([{"op": "set_theme", "theme": None if preset == "(current)" else preset, "colors": colors, "font": font}],
                   "Theme updated", "theme")
        st.rerun()
    st.caption("Applying a theme rebuilds the styling only; your data, formulas and charts are kept.")


def tab_data(v):
    spec = v["spec"]
    tables = [s for s in spec["sheets"] if s["type"] == "table"]
    if tables:
        name = st.selectbox("Table sheet", [s["name"] for s in tables])
        sh = next(s for s in tables if s["name"] == name)
        idx = [i for i, c in enumerate(sh["columns"]) if c["type"] != "formula"]
        names = [sh["columns"][i]["name"] for i in idx]
        df = pd.DataFrame([[r[i] for i in idx] for r in sh["rows"]], columns=names)
        edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, key=f"de_{name}_{v['n']}")
        own = st.checkbox("This is my own data (remove the 'sample data' notice)", value=not sh["sample_data"])
        st.caption("Dates as YYYY-MM-DD; percentages as fractions (0.25 = 25%). Formula columns calculate automatically in Excel.")
        if st.button("Save data changes as a new version", type="primary"):
            new = copy.deepcopy(spec)
            ns = next(s for s in new["sheets"] if s["name"] == name)
            rows = []
            for rec in edited.astype(object).where(edited.notna(), None).itertuples(index=False):
                full = [None] * len(sh["columns"])
                for i, val in zip(idx, rec):
                    full[i] = val.strftime("%Y-%m-%d") if hasattr(val, "strftime") else val
                if any(x not in (None, "") for x in full):
                    rows.append(full)
            ns["rows"], ns["sample_data"] = rows, not own
            if store_version(new, f"Edited data on {name}"):
                st.success("Saved.")
                st.rerun()
    st.divider()
    up = st.file_uploader("Add data from a CSV or Excel file", type=["csv", "xlsx"], key="add_data")
    if up is not None:
        try:
            sheets, warns = _import_cached(up.name, up.getvalue())
            for w in warns:
                st.warning(w)
            if st.button(f"Add {len(sheets)} sheet(s) to this workbook"):
                if store_version(E.merge_imported(spec, sheets), f"Imported {up.name}"):
                    st.rerun()
        except E.SpecError as e:
            st.error(str(e))


def render_versions():
    p = ss.project
    if not p:
        st.info("No project yet.")
        return
    for v in reversed(p["versions"]):
        c1, c2, c3 = st.columns([5, 1, 1])
        c1.write(f"**v{v['n']}** · {v['time']} · {v['note']}" + ("  ← current" if v is p["versions"][-1] else ""))
        c2.download_button("Download", v["bytes"], file_name=E.suggest_filename(v["spec"]).replace(".xlsx", f"_v{v['n']}.xlsx"),
                           mime=XLSX_MIME, key=f"dl_v{v['n']}")
        if v is not p["versions"][-1] and c3.button("Restore", key=f"rs_{v['n']}"):
            if store_version(copy.deepcopy(v["spec"]), f"Restored version {v['n']}"):
                st.rerun()


def tab_etsy(v):
    spec = v["spec"]
    st.write("Checks for a customer-ready product. They report what was actually tested and what was not.")
    R = E.etsy_checks(v["bytes"], spec)
    render_checks(R)
    fn = st.text_input("Product filename", E.suggest_filename(spec))
    a, b = st.columns(2)
    with a:
        if st.button("Generate PDF setup guide"):
            try:
                ss["pdf"] = E.build_pdf_guide(spec, fn)
            except Exception as e:  # noqa: BLE001
                st.error(f"PDF guide failed: {e}")
        if ss.get("pdf"):
            st.download_button("Download PDF guide", ss["pdf"], file_name=fn.replace(".xlsx", "_Guide.pdf"), mime="application/pdf")
    with b:
        if st.button("Draft listing copy with AI", disabled=not api_ready()):
            try:
                with st.spinner("Writing..."):
                    ss.listing = E.listing_copy(spec)
            except E.AIError as e:
                st.error(str(e))
    if ss.listing:
        L = ss.listing
        st.text_input("Title", L.get("title", ""))
        st.text_area("Description", L.get("description", ""), height=180)
        st.write("**Features:** " + " · ".join(L.get("features", [])))
        st.write("**Tags:** " + ", ".join(L.get("tags", [])))
        st.caption("Review and edit before publishing; do not claim compatibility you have not tested.")


def tab_download(v):
    spec = v["spec"]
    fn = E.suggest_filename(spec)
    st.download_button("⬇️ Download Excel workbook (.xlsx)", v["bytes"], file_name=fn, mime=XLSX_MIME, type="primary")
    st.caption(f"{len(v['bytes']) / 1024:.0f} KB · version {v['n']} · {len(spec['sheets'])} sheets")
    tables = [s for s in spec["sheets"] if s["type"] == "table"]
    if tables:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as z:
            for s in tables:
                d = pd.DataFrame([[x for x, c in zip(r, s["columns"]) if c["type"] != "formula"] for r in s["rows"]],
                                 columns=[c["name"] for c in s["columns"] if c["type"] != "formula"])
                z.writestr(f"{s['name']}.csv", d.to_csv(index=False))
        st.download_button("Sample data as CSV files (.zip)", buf.getvalue(), file_name="sample_data_csv.zip", mime="application/zip")
    summary = feature_summary(spec, v["checks"])
    st.download_button("Feature & quality summary (.md)", summary, file_name="summary.md", mime="text/markdown")
    st.download_button("Project file (.json) - to continue later",
                       json.dumps({"app": "ai-excel-studio", "version": 1, "spec": spec, "design": ss.project.get("design")}),
                       file_name=fn.replace(".xlsx", ".project.json"), mime="application/json")
    with st.expander("Included worksheets and features"):
        st.markdown(summary)


def page_workspace():
    if not ss.project:
        st.info("No workbook yet. Create a plan or build the offline example from New Project.")
        return
    v = cur()
    st.header(f"📗 {v['spec']['workbook_name']}")
    tabs = st.tabs(["Preview", "Edit with AI", "Quality check", "Theme", "Data", "Versions", "Prepare for Etsy", "Download"])
    for t, fn in zip(tabs, (tab_preview, tab_edit, tab_quality, tab_theme, tab_data,
                            lambda _v: render_versions(), tab_etsy, tab_download)):
        with t:
            fn(v)


def page_templates():
    st.header("Template library")
    st.caption("Templates are starting points. Pick one, then refine it with images, text or voice.")
    for cat, items in E.TEMPLATES.items():
        with st.expander(cat, expanded=cat == "Finance"):
            for name, prompt in items.items():
                c1, c2 = st.columns([4, 1])
                c1.write(f"**{name}**  \n{prompt}")
                if c2.button("Use", key=f"tpl_{name}"):
                    ss.prefill, ss.desc_ver = f"{name}: {prompt}", ss.desc_ver + 1
                    ss.goto = "New Project"
                    st.rerun()


def page_settings():
    st.header("Settings & Help")
    st.write(f"**Gemini API key:** {'configured ✅' if api_ready() else 'missing ❌ (add GEMINI_API_KEY to Streamlit secrets)'}")
    st.write(f"**Model:** `{E.get_model()}` (override with the `GEMINI_MODEL` secret)")
    if st.button("Test Gemini connection", disabled=not api_ready()):
        try:
            st.success(f"Connected. Model replied: {E.test_connection()}")
        except E.AIError as e:
            st.error(str(e))
    st.markdown("""
**Limits and honesty notes**
- The free Gemini tier has rate and daily limits; if you hit them, wait and retry.
- Charts, validation and conditional formatting are native Excel features. PivotTables and slicers are not generated.
- Formulas are saved as formulas; Excel calculates them on opening. Previews here are a Python approximation.
- Existing workbooks are imported as data only (macros, images, old charts and pivots are not carried over).
- Voice input records audio in the browser and transcribes it with Gemini; you review the text before applying.
- Projects are kept in this browser session. Use the project file (.json) to continue later.
""")


# ------------------------------------------------------------------ routing
if ss.goto:
    ss.nav, ss.goto = ss.goto, None
st.sidebar.title("📊 AI Excel Studio")
page = st.sidebar.radio("Navigate", NAV, key="nav")
if ss.project:
    st.sidebar.success(f"Project: {cur()['spec']['workbook_name']}  \nVersion {cur()['n']}")
st.sidebar.caption("Reference images and data import are on the New Project page.")
{"New Project": page_new, "Build Plan": page_plan, "Workspace": page_workspace, "Templates": page_templates,
 "Version History": render_versions, "Settings & Help": page_settings}[page]()
