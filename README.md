# Excel Checker MVP

Browser-based spreadsheet validator that lets analysts upload Excel workbooks, auto-detect column data types, highlight invalid cells, resolve duplicate rows, and export a focused error report. The app intentionally avoids persistence—refreshing the page clears every upload.

## Tech Stack

- **Frontend:** Vue 3 + Vite (modern glassmorphism-inspired UI, inline grid editing, column override controls)
- **Backend:** FastAPI + Pandas + OpenPyXL for type inference, validation, and Excel export

## Getting Started

### Backend

```bash
cd backend
python -m venv .venv
.venv\Scripts\activate  # Windows PowerShell
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
```

### Frontend

```bash
cd frontend
npm install
npm run dev -- --host
```

Vite proxies `/api/*` calls to `http://localhost:8000`, so keep the FastAPI server running while developing the UI.

## Docker Deployment

For an on-prem or self-hosted setup you can run both services through Docker:

```bash
# from the project root
docker compose build
docker compose up -d
```

- The **frontend** is served by Nginx on [http://localhost:5173](http://localhost:5173) and automatically proxies `/api/*` requests to the backend container.
- The **backend** FastAPI service listens on container port `8000`, exposed on `http://localhost:9533`.
- To customize the API base path during the frontend image build, pass `--build-arg VITE_API_BASE=<value>` to the `docker compose build` command.
- Resource reservations/limits are included in `docker-compose.yml` to guarantee the backend gets at least 2 vCPUs/2 GB RAM (up to 4/4 GB) and the frontend gets 1 vCPU/1 GB (up to 2/2 GB). Adjust those values if your on-prem host has different capacity.

Stop the stack with `docker compose down`. These containers are stateless; restarting the stack clears every uploaded session exactly like the browser app.

## Core Features

1. **Excel ingestion:** Upload `.xls/.xlsx/.xlsm/.csv` files. Columns are auto-detected as string, integer, float, boolean, or date.
2. **Validation & fixing:** Invalid cells are highlighted in red, duplicates in purple. Edit values inline, override column data types, and re-run validation without re-uploading.
3. **Duplicate management:** Duplicate row groups are listed with one-click removal (keeps the first row and deletes the rest, re-validating afterward).
4. **Reporting:** Download a fresh Excel workbook with two extra tabs (`Errors`, `Duplicates`) summarizing every issue.
5. **Cleaning tips:** Sidebar lists additional one-click ideas (normalize casing, trim whitespace, split multi-value cells, validate against dictionaries) as future enhancements.

## API Overview

| Endpoint | Method | Description |
| --- | --- | --- |
| `/api/upload` | `POST` (multipart) | Accepts file upload, detects types, validates, returns session + grid payload. |
| `/api/validate` | `POST` | Re-validates with edited rows or manual column type overrides. |
| `/api/duplicates/remove` | `POST` | Removes given row IDs (keeps the first one in each selected group) and re-validates. |
| `/api/report/{sessionId}` | `GET` | Streams an Excel report with invalid cells highlighted and summaries. |

All endpoints keep session data in-memory only.

## Testing

Backend validation helpers include lightweight pytest coverage:

```bash
cd backend
pytest
```

## Manual QA Checklist

- Upload a sample file, confirm detected types match expectation.
- Edit a cell to violate the inferred type, run **Revalidate**, ensure the cell turns red.
- Override a column type (e.g., force boolean) and re-validate to see recalculated errors.
- Use the duplicate controls to drop redundant rows, confirm grid + stats update.
- Download an error report and open it in Excel to verify the `Errors` tab references the same cells highlighted in the UI.

Future enhancements include whitespace trimming, reference-list validation, and collaborative editing.

