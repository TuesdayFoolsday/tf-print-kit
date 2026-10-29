# print kit

Renders a bid JSON to a PDF.

    pip install -r requirements.txt   # Playwright + Chromium must already be available
    python3 render/render.py samples/sample.json

Output goes to `out/`.

## Price book

    python3 pricebook/price_book.py append quotes.json "Price Book/"           # supplier quotes -> monthly files
    python3 pricebook/price_book.py latest "Price Book/" --item brick
    python3 pricebook/price_book.py movers "Price Book/" --days 90

Input format and item vocabulary: docstring at the top of `pricebook/price_book.py`.

## Dashboard

    python3 pricebook/price_book.py export "Price Book/" "Price Book Latest.csv"   # Friday: also writes Price Movers.csv
    python3 dashboard/build_dashboard.py --bid-log "Bid Log 2026.csv" --bid-metrics "Bid Metrics 2026.csv" \
        --prices "Price Movers.csv" --job-cost "Job Cost History.csv" --out Dashboard.xlsx --data "Dashboard Data.txt"

`Dashboard.xlsx` (Excel tables + live formulas) is placed in SharePoint once. Each night only `Dashboard Data.txt`
(JSON, text) is uploaded; a Power Automate flow runs `dashboard/load_dashboard_data.ts` (Office Script) to load it
into the workbook's tables. The build recalculates a copy in LibreOffice, scans every formula and cross-checks the
headline numbers before saying "Upload: yes".

## Job cost history

    python3 jobcost/job_cost.py template <dir>          # blank intake files
    python3 jobcost/job_cost.py add actuals.csv "Bid Metrics 2026.csv" "Job Cost History.csv" [--splits splits.csv]

Intake columns and rules: docstring at the top of `jobcost/job_cost.py`.
