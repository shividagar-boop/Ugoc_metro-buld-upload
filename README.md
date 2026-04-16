# Tender Processor Deploy Package

## What was broken
Your original setup had a static `index.html` plus a local Python script with hardcoded desktop paths. Netlify can host the HTML, but it cannot run that Python script as-is.

Also, your original frontend expected users to upload:
- PDFs
- Bulk_Load_Template.xlsx
- Address_Lookup.xlsx

This package fixes that by:
- keeping the frontend on **Netlify**
- moving the Python processing to **Render**
- bundling the template and address lookup on the backend so the user only uploads PDFs

## Folder structure
- `frontend/index.html` → deploy on Netlify
- `backend/app.py` → deploy on Render
- `backend/data/Address_Lookup.xlsx`
- `backend/data/Bulk_Load_Template.xlsx`
- `netlify.toml` → tells Netlify to publish the frontend folder

## Deploy the backend on Render
1. Push this whole package to GitHub.
2. In Render, create a new **Web Service** from that repo.
3. Use these settings if Render does not auto-detect them:
   - Root Directory: `backend`
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app`
4. After deploy, copy the backend URL, for example:
   - `https://your-backend-name.onrender.com`
5. Test this:
   - `https://your-backend-name.onrender.com/health`

## Connect the frontend
Open `frontend/index.html` and replace:
- `https://YOUR-RENDER-BACKEND.onrender.com/process-tenders`

with your real backend URL, for example:
- `https://your-backend-name.onrender.com/process-tenders`

Commit and push again.

## Deploy the frontend on Netlify
1. Create a new site from the same GitHub repo.
2. Netlify should read `netlify.toml` automatically.
3. If it asks manually:
   - Publish directory: `frontend`
   - Build command: leave blank
4. Deploy.

## Important truth bomb
If you want **everything** on Netlify, you would need to **rewrite the backend in JavaScript/TypeScript or Go** for Netlify Functions. Netlify’s current Functions support is for TypeScript, JavaScript, and Go, and the functions runtime is Node.js.

## Why this architecture is the sane one
- Netlify = static frontend, fast and easy
- Render = Python backend, good for `pdfplumber`, `pandas`, `openpyxl`
- No hardcoded desktop paths
- No need for users to upload the lookup/template every time
