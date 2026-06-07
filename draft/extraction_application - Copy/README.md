# Extraction Application

This folder contains a minimal FastAPI server and frontend to upload a JSON file, preview rows, and generate a simple PowerPoint.

Quick start:

1. Create and activate a Python virtualenv and install dependencies:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

2. Run the server:

```powershell
python -m extraction_application.server
# or
python -m uvicorn extraction_application.server:app --reload --host 127.0.0.1 --port 8000
```

3. Open http://127.0.0.1:8000/ in your browser, upload `sample_input.json` or your JSON, then click "Generate PPT".

Notes:
- Outputs are saved to `extraction_application/outputs/` and served at `/outputs/`.
- The server expects JSON arrays or newline-delimited JSON objects.
