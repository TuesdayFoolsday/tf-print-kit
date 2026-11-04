# print kit

Renders a bid JSON to a PDF (the only proposal format; the .docx was dropped Nov 1 2026).

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
        --prices "Price Movers.csv" --job-cost "Job Cost History.csv" --out Dashboard.xlsx --data "Dashboard Data.txt" [--previous <dir>]

In production this runs in the private courier repo's `dashboard` Action every evening (7:20pm CST), not in a
Claude run: it downloads the CSVs and pending files byte for byte, merges, builds, and uploads the changed parts.

`Dashboard.xlsx` (Excel tables + live formulas) is placed in SharePoint once. Each night the data goes up as four
text parts, `Dashboard Data 1.txt` … `4.txt` (each ≤ 18 KB; unchanged parts are skipped with `--previous <dir>`);
Power Automate (Flow 3) passes all four to `dashboard/load_dashboard_data.ts` (Office Script), which joins them,
refuses a mix of nights, and loads the workbook's tables. The build recalculates a copy in LibreOffice, scans every formula and cross-checks the
headline numbers before saying "Upload: yes".

## Job cost history

    python3 jobcost/job_cost.py template <dir>          # blank intake files
    python3 jobcost/job_cost.py add actuals.csv "Bid Metrics 2026.csv" "Job Cost History.csv" [--splits splits.csv]

Intake columns and rules: docstring at the top of `jobcost/job_cost.py`.

## Bid Metrics (estimates)

    python3 metrics/bid_check.py <compact.json> "Cost Map.csv" "Merged Bid Metrics.csv" --project "<name>" \
        --units "Merged Bid Metrics Units.csv" [--grout ... | --grout-walls ...] --pending-out <dir> --units-pending-out <dir>
    python3 metrics/india_metrics.py <workbook.xlsx> "Cost Map.csv" --tabs <tabs> --project "<name>" --metrics-id in-<slug> --status final
    python3 metrics/metrics_log.py merge "Bid Metrics 2026.csv" <pending_dir> --out "Merged Bid Metrics.csv"

New Bid Check rows are saved as `draft`; a bid's row becomes `final` at submission (System Rules C6, A15).

## Pending files (unattended runs never rewrite a large file)

    python3 bidlog/bid_log.py patch "Bid Log 2026.csv" <out_dir> --bid-id <id> --set status=lost --by "hourly 11:52"
    python3 bidlog/bid_log.py merge "Bid Log 2026.csv" <pending_dir> --out "Merged Bid Log.csv"   # readers merge first
    python3 pricebook/price_book.py pending quotes.json <out_dir>                                  # one small file per quote
    python3 pricebook/price_book.py latest "Price Book/" --pending "Price Book/Pending" --item brick
    python3 pricebook/price_book.py merge "Price Book/" "Price Book/Pending"                       # consolidation, in chat

Runs upload each pending file as a new small file; nothing existing is retyped. Bid Log, Bid Metrics and Units are
consolidated by the courier repo's `dashboard` Action (push `trigger/dashboard.txt` with `merge`); the Price Book in
chat. Applied files move to `Merged/`.

## Courier (binary files into SharePoint without typing their bytes)

    python3 courier/send.py send --repo <courier clone> --drive <driveId> --folder <folderItemId> [--replace] <pdf>=<name>
    python3 courier/send.py wait --repo <courier clone> <id> [--timeout 300]     # 0 ok · 1 failed · 2 no result yet
    python3 courier/send.py directory --repo <courier clone> contacts.csv          # new contacts -> Directory.csv (B3)

The file is sealed (`courier/box.py`: X25519 + AES-GCM) into the private courier repo's `inbox/`; that repo's
GitHub Action (`courier/workflow.yml`) runs `courier/receive.py`: opens it, uploads through Microsoft Graph,
downloads it back, compares SHA-256, writes `results/<id>.json` (sizes, fingerprints, item ids — no names) and
empties the inbox. Keys: `send.py keygen` (private key → the repo secret COURIER_PRIVATE_KEY).
