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

`Dashboard.xlsx` (Excel tables + live formulas) is placed in SharePoint once. Each night the data goes up as four
text parts, `Dashboard Data 1.txt` … `4.txt` (each ≤ 18 KB; unchanged parts are skipped with `--previous <dir>`);
Power Automate (Flow 3) passes all four to `dashboard/load_dashboard_data.ts` (Office Script), which joins them,
refuses a mix of nights, and loads the workbook's tables. The build recalculates a copy in LibreOffice, scans every formula and cross-checks the
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

## Courier (binary files into SharePoint without typing their bytes)

    python3 courier/send.py send --repo <courier clone> --drive <driveId> --folder <folderItemId> [--replace] <pdf>=<name>
    python3 courier/send.py wait --repo <courier clone> <id> [--timeout 300]     # 0 ok · 1 failed · 2 no result yet

The file is sealed (`courier/box.py`: X25519 + AES-GCM) into the private courier repo's `inbox/`; that repo's
GitHub Action (`courier/workflow.yml`) runs `courier/receive.py`: opens it, uploads through Microsoft Graph,
downloads it back, compares SHA-256, writes `results/<id>.json` (sizes, fingerprints, item ids — no names) and
empties the inbox. Keys: `send.py keygen` (private key → the repo secret COURIER_PRIVATE_KEY).

**Fetch (exact copies out of SharePoint, for editing in chat):**

    python3 courier/send.py fetch --repo <courier clone> --drive <driveId> --item <itemId> --out <local file>
    python3 courier/send.py fetch --repo <courier clone> --drive <driveId> --path "_Knowledge/Run Notes.md" --out <local file>
    python3 courier/send.py send --repo <courier clone> --drive <driveId> --folder <folderId> --replace \
        --if-unchanged <sha256 from fetch> <local file>=<name>                     # send the edit back

`fetch` makes a one-time key and pushes a sealed request to the courier repo's `fetch/`; its `fetch` Action
(`courier/fetch-workflow.yml` → `courier/fetch.py`) downloads the file twice, checks both copies agree, seals it with
the one-time key into `pickup/` and writes `results/fetch-<id>.json` (size, fingerprint, item id — no names).
`fetch` opens it, checks the fingerprint, saves the exact bytes, deletes the pickup and throws the key away (so the
sealed copy left in git history can't be opened by anyone), then prints the `send` command for the way back.
`--if-unchanged` makes the courier refuse the replace if the file in SharePoint changed since the fetch. Uncollected
pickups are deleted after 24 hours.
