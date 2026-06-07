# GLPI Reporting Project

This repository contains a small reporting and automation project with:
- a Python reporting pipeline for GLPI data,
- a simple backend app,
- a frontend interface,
- notebooks for analysis,
- generated reports and slides.

## Project structure

- `reports/` — main reporting scripts, input files, and generated outputs
- `backend/` — backend application files
- `frontend/` — frontend HTML/JS files
- `notebooks/` — analysis notebooks
- `draft/` — temporary or non-essential files kept for review

---

## Requirements

Use Python 3.10+.

Install the Python dependencies:

```powershell
cd reports
python -m pip install -r requirements.txt
```

---

## Run the reporting pipeline

### 1) Generate charts and visuals

```powershell
cd reports
python process_and_visualize.py --output-dir slides
```

This creates PNG visualizations in `reports/slides/`.

### 2) Generate the PowerPoint report

```powershell
cd reports
python generate_pptx_from_pngs.py --images-dir slides --output SGLN_Monthly_Report_V1_V2_corrected.pptx
```

### 3) Run the main report script directly

```powershell
cd reports
python glpi_pipeline.py --input csvjson.json --output generated_report.pptx
```

---

## Run the backend

```powershell
cd backend
python server.py
```

The backend app is the API / processing service for the project.

---

## Run the frontend

You can open the frontend directly in a browser, or serve it locally with:

```powershell
cd frontend
python -m http.server 8000
```

Then open:

```text
http://localhost:8000
```

---

## Docker option

From the `reports` folder:

```powershell
cd reports
docker build -t glpi-report .
docker run --rm -v "%cd%":/app glpi-report
```

You can also use the existing helper script:

```powershell
cd reports
.
un_report.bat
```

---

## GitHub setup and push

If the repository is not initialized yet:

```powershell
git init
git branch -M main
git remote add origin https://github.com/your-username/your-repo-name.git
```

Add and commit changes:

```powershell
git add .
git commit -m "Update project files"
```

Push to GitHub:

```powershell
git push -u origin main
```

If the remote is already set, you can simply use:

```powershell
git add .
git commit -m "Your commit message"
git push origin main
```

---

## Notes

- The main report pipeline is in `reports/`.
- Keep the `draft/` folder only for review until you are sure the files are no longer needed.
- Generated PPTX and PNG files are useful outputs, but you may clean them later if you want a lighter repo.
