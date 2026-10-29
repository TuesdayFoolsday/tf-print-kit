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

    python3 pricebook/price_book.py export "Price Book/" "Price Book Latest.csv"   # Friday: latest prices + changes
    python3 dashboard/build_dashboard.py --bid-log "Bid Log 2026.csv" --bid-metrics "Bid Metrics 2026.csv" \
        --prices "Price Book Latest.csv" [--job-cost "Job Cost History.csv"] --out Dashboard.xlsx

Builds the workbook with live formulas, recalculates it in LibreOffice, checks every formula and re-checks the
headline numbers in Python. Last line says whether to upload.

## Job cost history

    python3 jobcost/job_cost.py template <dir>          # blank intake files
    python3 jobcost/job_cost.py add actuals.csv "Bid Metrics 2026.csv" "Job Cost History.csv" [--splits splits.csv]

Intake columns and rules: docstring at the top of `jobcost/job_cost.py`.
