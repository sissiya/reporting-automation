@echo off
set IMAGE_NAME=glpi-report

echo Building Docker image %IMAGE_NAME%...
docker build -t %IMAGE_NAME% .
if errorlevel 1 (
    echo Build failed.
    exit /b 1
)

echo Running GLPI report generation...
docker run --rm -v "%cd%":/app %IMAGE_NAME%
if errorlevel 1 (
    echo Docker run failed.
    exit /b 1
)

echo Done. Generated files will be in %cd%.
