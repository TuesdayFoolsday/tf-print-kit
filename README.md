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

## Pending files (unattended runs never rewrite a large file)

    python3 bidlog/bid_log.py patch "Bid Log 2026.csv" <out_dir> --bid-id <id> --set status=lost --by "hourly 11:52"
    python3 bidlog/bid_log.py merge "Bid Log 2026.csv" <pending_dir> --out "Merged Bid Log.csv"   # readers merge first
    python3 pricebook/price_book.py pending quotes.json <out_dir>                                  # one small file per quote
    python3 pricebook/price_book.py latest "Price Book/" --pending "Price Book/Pending" --item brick
    python3 pricebook/price_book.py merge "Price Book/" "Price Book/Pending"                       # consolidation, in chat

Runs upload each pending file as a new small file; nothing existing is retyped. Consolidation folds them into the
base files and moves them to `Merged/`.
