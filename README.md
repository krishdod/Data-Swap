## Sheet Updating (Source → Template Excel mapper)

This is a small local web app:

- Upload **Source.xlsx** (sheet with data)
- Upload **Template.xlsx** (sheet with only header row)
- Map **Source columns → Template columns** with a preview
- Download a **new filled Template.xlsx** (plus audit sheets)

### Run locally (Windows)

Open PowerShell in this folder and run:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Then open `http://127.0.0.1:8000`.

### Notes

- The app is **deterministic**: it never drops rows, and it never writes anything unless every template column is mapped (or explicitly set to blank/constant).
- “AI suggestions” are local heuristic suggestions (no internet / no external model by default). You can later plug in an LLM if you want.

