# Daily banking importer — Station Hotel Aberdeen

## Browser interface — Windows and Mac
The browser interface runs locally and has PDF/Excel file pickers, the three ledger checks, an entry review and Excel/CSV download buttons. No hosted account is needed.

**Windows:** Install Python 3.12 from python.org, extract the repository ZIP, open `banking-importer`, run **Install Windows.bat** once, then **Run Web Windows.bat** each day.

**Mac:** Install Python 3.12 from python.org. In Terminal, enter `cd ` (with a trailing space), drag the extracted `banking-importer` folder into the window and press Return. Then run:
```sh
bash "Install Mac.command"
bash "Run Web Mac.command"
```
After installation you only need the second command each day. These Terminal commands also work if a downloaded ZIP loses executable permissions or macOS prevents opening `.command` files directly.

The launcher opens your browser automatically. If it does not, enter `http://127.0.0.1:8765` in your own browser. Keep the Terminal window open while using the app; press Ctrl+C to stop it. Download the workbook before stopping. Files are processed locally; temporary uploaded copies are removed after each import, and download data is discarded when the app stops. Use the updated workbook as the next day's input. Recalculate and save it in Excel first so its opening balance caches are current.

For an existing installation, rerun the installation command to add the web dependencies. The app is intended for local use on your own computer, and listens only on its loopback address.

## Windows file-picker alternative
1. Install Python 3 (3.12 recommended) from https://www.python.org/downloads/windows/, including the Python launcher and Tcl/Tk.
2. Extract this ZIP into a folder. Double-click **Install Windows.bat** once. This downloads the PDF-reading dependencies.
3. Each day, close the workbook in Excel, double-click **Run Windows.bat**, select the daily PDF, then the monthly workbook.
4. A new `Banking import ...` folder beside your workbook contains an updated workbook copy and a review CSV. The original is never overwritten.
5. Open the copy in desktop Excel. Use Formulas > Calculate Now (F9), or Ctrl+Alt+F9 for a full recalculation. Compare guest, AR and deposit ledger closing balances with the review CSV. Use this updated copy as the input on the next day.

## What it imports
The report date selects the week and daily column. Revenue, service charge, VAT, occupied rooms, sleepers, guest payments, deposit payments and deposits transferred are imported. Guest and deposit payments are kept separate. Mapping rules are in `mapping.json`, based on the supplied report and the grouping confirmed in chat. The supplied posting catalogue contains 243 codes; 192 have mappings. `posting-code-coverage.csv` lists every code, its mapping or its need for review. Package profit/loss, generic conference packages (4930), internal transfers, cheques and other ambiguous items remain blocked until their banking treatment is confirmed. Unmapped codes stop the import with their code and name.

Accommodation includes Nor1 upgrades. Restaurant dinner beverage includes soft drinks; bar dinner beverage includes night beverage. Pay By Link and Web Payment are grouped together within the ledger where they appear. Luggage storage maps to Other Misc Income; parking maps to Car Parking. Payment signs are reversed to match the banking sheet. Review these mappings before regular use.

Rows without source transactions remain blank. Existing totals, date formulas, ledger formulas, formatting, validation extensions and other workbook contents are preserved. Actual cash banked, PDQ totals, covers, change orders, corrections and any figures absent from the PDF still require another source or manual entry. Available rooms remains the existing workbook formula (115), even though the report has one out-of-order room.

## Checks and limits
The importer checks displayed revenue, non-revenue and payment totals, guest/deposit payment totals, payment splits and mapped totals. Revenue differences up to 2p are allowed with a visible rounding notice; larger differences stop the import. Other totals must match exactly. It checks the three closing ledger balances against the report using the workbook formulas and cached prior-day balances. If all opening balances match and the guest closing discrepancy is at most 2p, it writes the signed rounding correction to Guest / Differences (row 300). A positive entry reduces the guest balance; a negative entry increases it. Any remaining ledger mismatch blocks workbook creation. The supplied 12 October report has a 1p revenue difference. With matching opening balances, the importer offsets this 1p through row 300. The supplied workbook has a separate £236.33 prior-day guest balance mismatch, so the revised importer deliberately refuses to produce a balanced workbook from it. Correct/recalculate the prior day first; do not use a large Differences entry to hide the mismatch.

It refuses occupied destination cells, formulas, missing dates, different hotel names and changed report layouts. It supports text-based PDFs containing trial-balance pages followed by one manager flash report, including extra trial-balance pages for busier days. Scanned PDFs need OCR first; this version does not perform OCR. Date formulas must have cached results: open/save a new monthly template in Excel first. Excel itself recalculates formulas; the script does not.

PDF processing and workbook updates stay on your computer. No report is uploaded by the program. Only the initial dependency installation needs Internet access.

## Command line / cloud
```sh
python -m pip install -r requirements.txt
python banking_import.py report.pdf banking.xlsx --output-dir review
python banking_import.py report.pdf banking.xlsx --output-dir imported --apply
```
Choose separate output directories for separate runs. Without `--apply`, only a CSV preview is produced. No workbook is changed. Windows file-picker interactions require local desktop Python and have not been tested on Windows in this Linux environment.

## Department grouping
Regular room-service food/drinks and Nor1 meal extras use Restaurant rows by meal (30–38); untimed Nor1 wine uses Restaurant Beverage-General (34). Night postings join Dinner. Accommodation extras and their adjustments join Accommodation (13), preserving the report sign.

Ballroom, banquet, conference and wedding postings all use C&B rows 40–54: room hire 40; food 43–47; beverage 48–52; miscellaneous 53; equipment/AV hire 54. Untimed food/drinks and tea/coffee use General rows. Corkage, entertainment and external supplies use C&B Misc Income. These defaults preserve totals while making the category choices explicit; adjust the mapping if your hotel's accounting treatment differs.
