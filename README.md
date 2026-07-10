# Finance Management System

Simple version: **one backend file** (`api/main.py`) + plain HTML/CSS/JS
frontend. Data is stored as an Excel file in a GitHub repo (no database).

## Project structure

```
.
├── vercel.json         Routes /api/* to api/main.py
├── requirements.txt    Python dependencies
├── api/
│   └── main.py          The ENTIRE backend - login, save/load entries,
│                         GitHub read/write, report calculations, all in one file
├── index.html           Login page
├── dashboard.html        Summary cards + Download Excel button
├── entry.html            Form to add/edit a quarter's numbers
├── reports.html          Report table (Amounts / QoQ% / YoY%)
├── style.css             All styling
└── app.js                Small shared helper (login check, fetch calls, toast messages)
```

## 1. Push to GitHub

```bash
git init
git add .
git commit -m "Finance management system"
git branch -M main
git remote add origin https://github.com/<you>/<repo>.git
git push -u origin main
```

## 2. Create a GitHub token

GitHub → Settings → Developer settings → Fine-grained tokens → generate one
scoped to this repo with **Contents: Read and write** permission.

## 3. Deploy on Vercel

1. [vercel.com/new](https://vercel.com/new) → Import this repo.
2. Framework preset: "Other" (it auto-detects static + Python function).
3. Add these Environment Variables:

   | Name | Value |
   |---|---|
   | `GITHUB_TOKEN` | your personal access token |
   | `GITHUB_REPO` | `your-username/your-repo` |
   | `GITHUB_BRANCH` | `main` |
   | `EXCEL_PATH` | `data/finance_data.xlsx` |
   | `ACCESS_CODE` | a shared password anyone uses to log in |

4. Deploy. Every `git push` to `main` auto-redeploys.

## 4. Use it

Visit your `*.vercel.app` URL → log in with any name + the `ACCESS_CODE` →
add quarterly entries → view Reports and Dashboard. The Excel file is
created automatically in your repo on the first save.

## Customizing line items

Edit `STANDARD_PARTICULARS` near the top of `api/main.py` to match your own
chart of accounts. Anything typed into the entry form's "Other" section is
picked up automatically in reports too.
