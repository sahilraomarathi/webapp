# Gas Connection Document Generator

Fills DOCX templates and downloads them directly — no LibreOffice needed.
Users open the downloaded .docx on their phone (Word, Google Docs, WPS) and convert to PDF from there.

## Requirements
- Node.js v16+ (https://nodejs.org)
- Python 3 (https://python.org)

## One-time Setup
```bash
cd server
npm install
pip install python-docx
```

## Run the Server
```bash
node index.js
```
Then open: **http://localhost:3000**

## Deploy Free (Render.com — Recommended)
1. Push this folder to GitHub
2. Go to render.com → New → Web Service → connect your repo
3. Build Command: `npm install && pip install python-docx`
4. Start Command: `node index.js`
5. Done — you get a free `yourapp.onrender.com` URL

## How it works
1. Fill in the form fields
2. Click "Download Document" — a .docx file downloads
3. Open the .docx on your phone in Word / Google Docs / WPS
4. Export/Print as PDF from there
