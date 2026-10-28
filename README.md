# ai-excel-studio# AI Excel Studio

Streamlit app that turns reference images, text prompts, voice and CSV/Excel data into a **real, downloadable `.xlsx`**
(native charts, formulas, dropdowns, conditional formatting), then lets you keep editing the same workbook with plain-English requests.

## Files
| File | Purpose |
|---|---|
| `app.py` | Streamlit interface and workflow |
| `engine.py` | Gemini calls, workbook builder, safe edit operations, quality checks, PDF guide |
| `requirements.txt` | Python dependencies |
| `README.md` | This file |

## Run locally
```bash
pip install -r requirements.txt
export GEMINI_API_KEY="your-key"      # Windows PowerShell: $env:GEMINI_API_KEY="your-key"
streamlit run app.py
```
Get a free key at https://aistudio.google.com/apikey. Never commit it.

## Deploy (GitHub + Streamlit Community Cloud)
1. Create a **`.gitignore`** file containing these lines, then push all files to a GitHub repo:
   ```
   .streamlit/secrets.toml
   .env
   __pycache__/
   *.xlsx
   ```
2. On https://share.streamlit.io choose **New app**, pick the repo, branch and `app.py`.
3. Open **Advanced settings -> Secrets** and paste:
   ```toml
   GEMINI_API_KEY = "your-key"
   # GEMINI_MODEL = "gemini-2.5-flash"   # optional override
   ```
4. Deploy. The key lives only in Streamlit's secrets store, never in the repo or the UI.

## How it works (safety)
The AI never writes code. It returns a JSON *spec* (sheets, columns, KPIs, charts) or a list of whitelisted edit *operations*.
`engine.py` validates everything, blocks risky formula functions, and builds the file with openpyxl.
Destructive edits (deleting a sheet or column) always ask for confirmation first.

## Honest limitations
- Formulas are saved as formulas; Excel calculates them on opening (`fullCalcOnLoad`). In-app previews are a Python approximation.
  If LibreOffice is installed (add `libreoffice-calc` to a `packages.txt` on Streamlit Cloud) the quality check also recalculates the file; otherwise it reports that it could not.
- No PivotTables, slicers or macros. Importing an existing workbook brings in data only.
- Free Gemini tier has rate/daily limits. Voice input records audio and transcribes with Gemini (needs mic permission, works over HTTPS).
- Projects live in the browser session; use *Download -> Project file* to continue later.
- Not tested in desktop Excel, Excel for Mac or Google Sheets from this app.
