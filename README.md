# PhillipCapital Risk Management — Credit Worksheet Processor

Internal tool for automating the extraction and filing of FOCUS report data into credit worksheets and net capital tracking workbooks.

---

## What This Does

- Upload one or multiple FOCUS-report PDFs
- Extracts helper-code values automatically from each PDF
- Writes results into a combined Excel workbook per customer per year
  - Monthly credit worksheet tabs (`March 2026`, `June 2026`, …)
  - Accumulating `Net Capital YYYY` tab with one column per month
- Tracks quarterly filing status (March / June / September / December) per customer
- Customer portal with per-year financial summaries (Total Equity, Net Capital, Excess Net Capital)

---

## Requirements

| Requirement | Minimum Version | Notes |
|---|---|---|
| Python | 3.10+ | 3.11 recommended |
| pip | any recent | comes with Python |
| Git | any | to clone the repo |

No database, no Docker, no cloud account required.

---

## First-Time Setup

### 1. Clone the Repository

```bash
git clone https://github.com/riskteam01/riskdashboard.git
cd riskdashboard
```

### 2. Create a Virtual Environment

**Windows (PowerShell):**
```powershell
python -m venv venv
venv\Scripts\Activate.ps1
```

If you get a script execution error on Windows, run this first (once):
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**macOS / Linux:**
```bash
python3 -m venv venv
source venv/bin/activate
```

You should see `(venv)` at the start of your terminal prompt when the environment is active.

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

This installs:

| Package | Purpose |
|---|---|
| `fastapi` | Web framework |
| `uvicorn` | Runs the web server |
| `python-multipart` | Handles file uploads |
| `pymupdf` | Reads and parses PDF files |
| `openpyxl` | Reads and writes Excel files |

### 4. Add Your Excel Template

The application requires a Credit Worksheet Excel template to generate output files.

1. Place your template `.xlsx` file inside the `templates/` folder
   *(the folder is created automatically on first run — or create it manually now)*
2. The preferred filename is:
   ```
   templates/Buckler Excel Credit WS Template.xlsx
   ```
   If the file is named differently, the app will fall back to the first `.xlsx` it finds in that folder.

> **Important:** Without a template the app will start but PDF processing will fail. Add the template before running any batch jobs.

If you also have a Net Capital template, place it in `templates/` as well. The app detects it automatically by looking for a sheet whose name contains `Net Capital`.

### 5. Run the Server

```bash
python server.py
```

You should see:

```
INFO:     Uvicorn running on http://0.0.0.0:8000 (Press CTRL+C to quit)
INFO:     Application startup complete.
```

Open your browser and go to:

```
http://localhost:8000
```

> `0.0.0.0` in the terminal output is not a URL you can visit — always use `localhost:8000`.

---

## Logging In

On first launch the app creates a default admin account:

| Field | Value |
|---|---|
| Username | `admin` |
| Password | `admin` |

**Change this immediately** after logging in via **Settings → Users**.

---

## Directory Structure

These folders are created automatically on first startup. You do not need to create them manually (except `templates/` if you want to add your template before the first run).

```
riskdashboard/
├── templates/        ← Put your Excel template(s) here
├── uploads/          ← Temporary PDF uploads (auto-cleaned after 7 days)
├── outputs/          ← Legacy output files
├── net_capital/      ← Combined customer workbooks (one per customer per year)
├── audits/           ← Per-run audit text files
├── logs/             ← Debug logs (auto-cleaned after 30 days)
├── assets/           ← Logo and static files
├── users.json        ← User accounts (auto-created)
├── customers.json    ← Customer records (auto-created)
├── app_config.json   ← App settings (auto-created)
└── server.py         ← Entry point
```

---

## Everyday Use

### Starting the Server

Every time you open a new terminal session, activate the virtual environment first:

**Windows:**
```powershell
venv\Scripts\Activate.ps1
python server.py
```

**macOS / Linux:**
```bash
source venv/bin/activate
python server.py
```

### Stopping the Server

Press `Ctrl + C` in the terminal.

---

## Updating

To pull the latest version:

```bash
git pull origin main
pip install -r requirements.txt
python server.py
```

---

## Troubleshooting

### `venv\Scripts\Activate.ps1 cannot be loaded` (Windows)
Your PowerShell execution policy is blocking scripts. Run once:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```
Then activate again.

### `ModuleNotFoundError: No module named 'fastapi'` (or any other package)
Your virtual environment is not active. Run the activate command for your OS (step 2 above) and then re-run `python server.py`.

### `Port 8000 is already in use`
Something else is running on port 8000. Either stop that process, or change the port in `server.py`:
```python
uvicorn.run("server:app", host="0.0.0.0", port=8001, reload=True)
```
Then visit `http://localhost:8001`.

### PDF processing fails with "no template found"
The `templates/` folder is empty or contains no `.xlsx` file. Add your Credit Worksheet template (see step 4).

### `######` or `3E+08` in the Excel output
Column auto-sizing runs at save time. If you see this in an older file, re-process the PDF — newly saved workbooks will have auto-fitted columns.

### Pages load but the logo is missing
Place your logo file (`logo.png` or `logo.jpg`) inside the `assets/` folder. The app serves it as a static file.

---

## Browser Support

Any modern browser works. Tested on Chrome and Edge. No extensions required.

---

## Notes for Developers

- The server runs with `reload=True` by default — file changes are picked up automatically without restarting.
- All data is stored locally as JSON files and Excel workbooks. There is no external database.
- PDF text extraction uses PyMuPDF (`fitz`). Scanned PDFs without embedded text will produce blank or missing fields.
- The `net_capital/` directory holds the combined customer workbooks. Back this folder up regularly — it is the primary output store.
- `customers.json` maps customer IDs to their name and full report history. Do not delete it unless you intend to reset all customer records.
