# LPO Month & Amount Calculator

A full-stack web application built for the Planning and Procurement Department at **United Engineering Projects (UEP) LLC, Muscat, Oman**. The app automates the translation of procurement delivery and payment clause text into calendar-aligned project milestones and cash-flow projections — eliminating a previously manual, time-intensive process across thousands of LPO entries.

---

## Live Demo

**Frontend:** [lpo-month-amount-calculator.vercel.app](https://lpo-month-amount-calculator.vercel.app)  
**Backend API:** [lpo-calculator-backend.onrender.com](https://lpo-calculator-backend.onrender.com)

---

## What It Does

The application accepts a UEP procurement Excel workbook and processes the **LPO MAT - PO** sheet through two automated tasks:

### Task A — Delivery & Payment Month Calculation
- Scans every **OPEN** row in the procurement ledger
- Reads the **LPO Date** (column E) as a start date
- Parses the **Delivery Terms** (column J) and **Payment Terms** (column K) using an intelligent multi-layer term parser
- Calculates the resulting end month for both delivery and payment due
- Writes the **LPO Amount** into the correct month column under the **Delivery** and **Payment Due** section headers (Jan-2025 through Dec of the current year)
- Highlights rows in **pale yellow** where terms could not be resolved, for manual review
- Returns a detailed breakdown of which rows were unresolved and why

### Task B — LPO Amount Aggregation
- Allows the user to select one or multiple job numbers
- Sums all LPO amounts across every matching row for each selected job number
- Writes the total against the last matching row in column Y
- Displays a summary table showing job number, all row numbers found, total amount, and placement location

---

## Features

- Drag and drop Excel file upload
- Automatic month column generation (Jan-2025 → Dec of current year) for both Delivery and Payment Due sections
- Intelligent term parser handling 10+ delivery/payment clause formats
- Detailed post-processing stats panel showing:
  - Total job number rows
  - Open and closed row counts
  - Rows highlighted for manual review (with expandable row-by-row breakdown)
  - Delivery and payment cells concluded
  - Delivery and payment unresolved row counts
- Full **Rows Requiring Manual Review** table with issue badges (Delivery unresolved / Payment unresolved / Both unresolved)
- Task B modal popup showing job totals before download
- Download of fully processed Excel file with borders, headers, and all computed values

---

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Frontend | HTML, CSS, JavaScript |
| Backend | Python, FastAPI |
| Excel Processing | openpyxl |
| Date Calculation | Python datetime, dateutil |
| Frontend Hosting | Vercel |
| Backend Hosting | Render |
| Version Control | Git, GitHub |

---

## Project Structure

```
LPO_Webapp/
├── backend/
│   ├── main.py                  # FastAPI app and /process endpoint
│   ├── requirements.txt
│   ├── Procfile                 # Render deployment config
│   └── services/
│       ├── lpo_mat_processor.py # Core Excel processing engine
│       └── term_parser.py       # Delivery/payment term parsing logic
└── frontend/
    ├── index.html
    └── src/
        ├── app.js               # Frontend logic and API integration
        ├── styles.css
        └── config.js            # API URL configuration
```

---

## How the Term Parser Works

The term parser (`term_parser.py`) interprets plain-English delivery and payment clauses through a priority chain:

1. **Immediate terms** — EX-STOCK, in advance, at sight, cash on delivery → end date = LPO date
2. **Explicit month names** — "Material by July/Aug-2026" → extracts month(s) directly
3. **Duration calculation** — "90 Days PDC", "4-6 Weeks", "3 Months" → adds to LPO date
4. **Within N days** — "within 30 days" → adds days to LPO date
5. **Unresolvable** → row flagged for manual review and highlighted yellow in Excel

Supports ranges (takes the maximum: "4-6 weeks" → 6 weeks), dual month references ("Jul/Aug-2026" → writes amount to both columns), and multiple date string formats including `DD/MM/YYYY`, `DD-MM-YYYY`, `YYYY-MM-DD`, and `MM-DD-YYYY`.

---

## API

### `POST /process`

Accepts a multipart form with:
- `file` — the `.xlsx` Excel workbook
- `selected_job_numbers` — JSON array of job number strings for Task B

Returns:
- The processed `.xlsx` file as a streaming response
- `X-Job-Numbers` header — JSON array of all unique job numbers found
- `X-Summary` header — JSON object with processing statistics

---

## Key Implementation Details

- **Formula-safe reading** — workbook is loaded twice: once for formatting/writing, once with `data_only=True` to read computed cell values from formula cells
- **Silent miss detection** — if the parser resolves a month but that month falls outside the column range, the row is still flagged for review rather than silently skipped
- **Decimal precision** — all LPO amount arithmetic uses Python's `Decimal` type to avoid floating point rounding errors
- **Case-insensitive remarks** — OPEN/Open/open and CLOSED/Closed/closed are all handled correctly
- **Border application** — borders are applied last, after all data writes, to ensure full coverage across all generated columns

---

## Local Development

### Backend

```bash
cd backend
python -m venv .venv
.venv\Scripts\activate      # Windows
pip install -r requirements.txt
uvicorn main:app --reload
```

### Frontend

Open `frontend/index.html` directly in a browser, or serve with any static file server. Update `frontend/src/config.js` to point to your local backend URL if needed.

---

## Built By

**Nekha Sudheer**  
Engineering Intern — Planning & Procurement Department  
United Engineering Projects LLC, Muscat, Oman  
May 2026
