# Bulk "residents count" correction for a billing web app (Selenium)

A Python + Selenium script that automates a repetitive back-office task in a
water-utility billing system (WG Craft). A bulk change had reset the number of
residents on many subscriber accounts to 1, and the correct values had to be
restored account by account. Doing this by hand takes hours for ~2,000
accounts; the script does it unattended.

> Built for an internal system I work with. **No real subscriber data, URLs or
> credentials are included** – the demo workbook contains fictional accounts.
> Use automation like this only on systems you are authorised to operate.

## What it does

For every account in an Excel list:

1. Opens *Subscribers → Individuals* and searches by account number.
2. Skips the account if the current value is not the expected one (safety check).
3. Creates a registration document (number `1`, date = today) and confirms.
4. Opens the residents dialog, sets the new count, fills the end date with
   today only if it is empty, and confirms.
5. Re-reads the account card to verify the new value and logs the result.

## Features

- **Reads only the visible rows of a filtered Excel sheet** (rows hidden by an
  autofilter are ignored), restores account numbers that lost a leading zero.
- **Resumable:** every result is appended to `progress.csv`; re-running skips
  finished accounts, so a crash, a lost connection or a shutdown costs nothing.
- **Safe by default:** running without arguments processes only 3 accounts;
  `--all` is required for a full run. `--dry-run` needs no browser.
- **Duplicate protection:** if an error happens after the registration
  document was created, the account is marked `error_partial` and never retried
  automatically.
- **Debuggable:** on any error a screenshot and the page HTML are saved to
  `errors/`; the run stops after several consecutive errors.
- **Slow-site tolerant:** explicit waits instead of fixed sleeps.
- **Secrets stay out of the code:** credentials and base URL come from `.env`.

## Setup

```bash
pip install -r requirements.txt
cp .env.example .env        # then fill in WG_BASE_URL, WG_LOGIN, WG_PASSWORD
```

Requires Python 3.9+ and Google Chrome (Selenium Manager downloads the driver).
Set `BROWSER = "edge"` in the script to use Edge instead.

## Usage

```bash
python set_residents.py --excel accounts_demo.xlsx --dry-run   # preview only
python set_residents.py            # test run: first 3 accounts
python set_residents.py --all      # full run, resumes from progress.csv
python set_residents.py --all --limit 300   # work in batches
```

Settings (Excel column, first data row, value pattern, timeouts) are at the top
of `set_residents.py`. The new value is taken from a repeating `PATTERN`
(`(value, count)` blocks) – replace `value_for()` to read it from Excel or a
database instead.

## Output

`progress.csv` has one row per account: `done`, `unverified`, `skipped_not_1`,
`not_found`, `error` or `error_partial`, plus a message and timestamp.

## Notes

- Element lookup is based on the visible (Russian-language) labels of the web
  app, so it needs adapting for another UI language or product.
- Tech: Python, Selenium, openpyxl, python-dotenv.
