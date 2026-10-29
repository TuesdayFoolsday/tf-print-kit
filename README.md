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
