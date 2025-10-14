### File Converter Web App

A Flask-based web application to convert various file formats:
- JPG to PNG
- JPG to PDF
- PNG to JPG
- PNG to PDF
- PDF to Word
- PDF to JPG
- PDF to PNG

## Deployment on Vercel

This app is optimized for Vercel serverless deployment with reduced dependencies to stay under the 250 MB limit.

## Installation

```bash
pip install -r requirements.txt
```

## Run Locally

```bash
python main.py
```

## Note

Word to PDF conversion is disabled on Vercel as it requires Windows-specific dependencies.
