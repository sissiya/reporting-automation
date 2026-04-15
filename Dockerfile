# Dockerfile for GLPI visualization and PowerPoint report
FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt ./
RUN python -m pip install --no-cache-dir -r requirements.txt

COPY . ./

CMD ["/bin/sh", "-c", "python process_and_visualize.py --output-dir slides && python generate_pptx_from_pngs.py --images-dir slides --output SGLN_Monthly_Report_V1_V2_corrected.pptx"]
