# GLPI Visualization and PowerPoint Report

## Overview
This solution processes GLPI incident data from `csvjson.json` and `Synthèse DC - Incident.json`, generates chart images, and creates a PowerPoint report.

It includes two main scripts:

- `process_and_visualize.py`: loads JSON data, normalizes French/encoded column names, computes SLA metrics, and creates visualization PNG files.
- `generate_pptx_from_pngs.py`: reads the generated PNG slides and builds a PowerPoint report with captions.

## What was fixed
- Added robust normalization for mis-encoded French column names such as `Priorit�` and `Temps de r�solution d�pass�`.
- Ensured `Priorité`, `SLA_Depasse`, `Date_Ouverture`, and `Date_Resolution` are mapped correctly.
- Verified the `csvjson.json` data and regenerated corrected visuals.

## Requirements
Install the required Python packages:

```powershell
python -m pip install -r requirements.txt
```

## Run the visualization pipeline
Generate the PNG visualizations from the JSON files and store them in the `slides/` folder:

```powershell
python process_and_visualize.py --output-dir slides
```

This will create files like:
- `slides/slide1_courbe_groupes_V1.png`
- `slides/slide1_source_V1.png`
- `slides/slide2_jauges_V1.png`
- `slides/slide2_tableau_V1.png`
- `slides/slide3_historique_V1.png`
- `slides/slide3_backlog_ttr_V1.png`
- `slides/slide1_courbe_groupes_V2.png`
- `slides/slide1_source_V2.png`
- `slides/slide2_jauges_V2.png`
- `slides/slide2_tableau_V2.png`
- `slides/slide3_historique_V2.png`
- `slides/slide3_backlog_ttr_V2.png`

## Build the PowerPoint report
Create the PowerPoint from the generated PNG images in `slides/`:

```powershell
python generate_pptx_from_pngs.py --images-dir slides --output SGLN_Monthly_Report_V1_V2_corrected.pptx
```

## Expected output
- `SGLN_Monthly_Report_V1_V2_corrected.pptx`
- Several PNG slide visuals in the working directory

## Notes
- The `csvjson.json` source contains only `Moyenne` priority tickets in the current dataset, so only the P3 SLA gauge is populated.
- If you want to use another JSON location, pass `--data-file-v1` and `--data-file-v2` to `generate_pptx_from_pngs.py`.

## Docker

Build the Docker image from the project root:

```powershell
docker build -t glpi-report .
```

Run the container to generate the visuals and PowerPoint in the current directory:

```powershell
docker run --rm -v "%cd%":/app glpi-report
```

For a smoother single command on Windows, run:

```powershell
.\run_report.bat
```

This script always rebuilds the Docker image with the latest code, then runs the container. It will create the PNG files in `slides/` and `SGLN_Monthly_Report_V1_V2_corrected.pptx` in the current folder.
