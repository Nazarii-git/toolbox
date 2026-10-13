"""Import Station Hotel night audit reports without rebuilding the Excel workbook."""
import argparse
import csv
from datetime import datetime
from decimal import Decimal
import json
from pathlib import Path
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
import pdfplumber

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
MONEY = r'-?\s*\d[\d,]*\.\d{2}'

def amount(value):
    return Decimal(value.replace(',', '').replace(' ', ''))

def require(condition, message):
    if not condition:
        raise ValueError(message)

ROUNDING_NOTES = []

def reconciles(actual, expected, label):
    difference = actual - expected
    require(abs(difference) <= Decimal('0.02'), f'{label}: difference {difference} exceeds two-pence display-rounding tolerance')
    if difference:
        ROUNDING_NOTES.append(f'{label}: displayed lines differ from report total by {difference}; source amounts preserved; see Differences entry and ledger checks for any rounding correction.')
    return True

def field(text, label):
    matches = re.findall(r'^' + re.escape(label) + r'\s+(' + MONEY + r')\s*$', text, re.M)
    require(len(matches) == 1, f'Missing or ambiguous report field: {label}')
    return amount(matches[0])

def transactions(text):
    result = {}
    for code, label, value in re.findall(r'^(\d{4})\s+(.+?)\s+(' + MONEY + r')\s*$', text, re.M):
        require(code not in result, f'Duplicate transaction code {code}')
        result[code] = (label, amount(value))
    return result

def parse_report(path):
    with pdfplumber.open(path) as pdf:
        pages = [p.extract_text() or '' for p in pdf.pages]
    require(len(pages) >= 2, 'Expected a trial balance followed by a manager flash report.')
    require(all('Station Hotel Aberdeen' in p for p in pages), 'Report hotel does not match Station Hotel Aberdeen.')
    require(all('Night Audit Trial Balance' in p for p in pages[:-1]) and 'Manager Flash Report' in pages[-1], 'Unexpected report layout.')
    dates = [re.findall(r'\b\d{2}/\d{2}/\d{2}\b', p) for p in pages]
    require(all(d for d in dates), 'Missing report date.')
    require(len(set(d for page in dates for d in page)) == 1, 'Report pages have different dates.')
    date = datetime.strptime(dates[0][0], '%d/%m/%y').date()
    trial = '\n'.join(pages[:-1])
    require('\nGuest Ledger\n' in trial, 'Missing guest ledger section.')
    first, second = trial.split('\nGuest Ledger\n', 1)
    manager = pages[-1]
    daily = transactions(first)
    # Sections, not numeric ranges, define accounting classification: e.g.
    # levy 7102 and package profit 7400 are revenue despite their code numbers.
    revenue_text = first.split('\nRevenue\n', 1)[1].split('\nRevenue Total ', 1)[0]
    nonrevenue_text = first.split('\nNon Revenue\n', 1)[1].split('\nNon Revenue Total ', 1)[0]
    payment_text = first.split('\nPayment\n', 1)[1].split('\nPayment Total ', 1)[0]
    revenue_codes, nonrevenue_codes, payment_codes = map(transactions, (revenue_text, nonrevenue_text, payment_text))
    require(set(daily) == set(revenue_codes) | set(nonrevenue_codes) | set(payment_codes), 'Transaction outside recognised daily report sections.')
    require('AR Ledger' in second and 'Deposit Ledger' in second and 'Package Ledger' in second, 'Missing ledger sections.')
    guest_text, remainder = second.split('\nAR Ledger\n', 1)
    ar_text, remainder = remainder.split('\nDeposit Ledger\n', 1)
    deposit_text, package_text = remainder.split('\nPackage Ledger\n', 1)
    groups = {'daily': daily, 'guest': transactions(guest_text), 'ar': transactions(ar_text), 'deposit': transactions(deposit_text)}
    revenue = field(first, 'Revenue Total')
    nonrevenue = field(first, 'Non Revenue Total')
    payments = field(first, 'Payment Total')
    reconciles(sum((v for c, (_, v) in daily.items() if c in revenue_codes), Decimal(0)), revenue, 'Revenue lines')
    require(sum((v for c, (_, v) in daily.items() if c in nonrevenue_codes), Decimal(0)) == nonrevenue, 'Non-revenue lines do not reconcile.')
    require(sum((v for c, (_, v) in daily.items() if c in payment_codes), Decimal(0)) == payments, 'Payment lines do not reconcile.')
    for code, (_, value) in daily.items():
        if code in payment_codes:
            split = sum((g.get(code, ('', Decimal(0)))[1] for g in (groups['guest'], groups['ar'], groups['deposit'])), Decimal(0))
            require(value == split, f'Guest/AR/deposit breakdown does not reconcile for {code}.')
    require(field(ar_text, 'Payments') == 0 or bool(groups['ar']), 'Nonzero AR payments lack transaction detail; manual review required.')
    for name, text in [('guest', guest_text), ('deposit', deposit_text)]:
        require(sum((v for _, v in groups[name].values()), Decimal(0)) == field(text, 'Payments'), f'{name} payment total does not reconcile.')
    def stat(label):
        match = re.search(r'^' + re.escape(label) + r'\s+(\d+)\s', manager, re.M)
        require(match is not None, f'Missing manager statistic {label}')
        return Decimal(match[1])
    # Manager rows have six columns. Verify today's revenue independently.
    match = re.search(r'^Total Revenue\s+(' + MONEY + r')\s', manager, re.M)
    require(match is not None and amount(match[1]) == revenue, 'Manager and trial balance revenue differ.')
    return {'date': date, 'groups': groups, 'stats': {7: stat('Rooms Occupied'), 9: stat('Total In-House Persons')},
            'transfer': field(deposit_text, 'Deposits Transfered at Check-In'),
            'opening': {320: field(first, 'Balance Brought Forward'), 321: field(ar_text, 'Balance Yesterday'), 322: field(deposit_text, 'Balance Yesterday')},
            'balances': {320: field(guest_text, 'Balance Today'), 321: field(ar_text, 'Balance Today'), 322: field(deposit_text, 'Balance Today')},
            'revenue': revenue, 'nonrevenue': nonrevenue, 'payments': payments, 'payment_codes': set(payment_codes)}

def workbook_target(path, date):
    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        targets = {r.attrib['Id']: r.attrib['Target'] for r in rels}
        matches = []
        for s in wb.find('s:sheets', NS):
            if not re.fullmatch(r'Week [1-6]', s.attrib['name']):
                continue
            relid = s.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
            target = targets[relid]
            target = target.lstrip('/') if target.startswith('/') else 'xl/' + target
            root = ET.fromstring(z.read(target))
            cells = {c.attrib['r']: c for c in root.findall('.//s:sheetData/s:row/s:c', NS)}
            # Resolve shared formulas in memory only; the original XML is preserved on save.
            from openpyxl.formula.translate import Translator
            masters = {}
            for ref, cell in cells.items():
                f = cell.find('s:f', NS)
                if f is not None and f.attrib.get('t') == 'shared' and f.text:
                    masters[f.attrib['si']] = (ref, f.text)
            for ref, cell in cells.items():
                f = cell.find('s:f', NS)
                if f is not None and f.attrib.get('t') == 'shared' and not f.text:
                    origin, formula = masters[f.attrib['si']]
                    f.text = Translator('=' + formula, origin=origin).translate_formula(ref)[1:]
            for col in 'FGHIJKL':
                cell = cells.get(col + '5')
                if cell is None: continue
                value = cell.find('s:v', NS)
                if value is None or value.text is None: continue
                # Excel serial dates (the workbook uses the 1900 date system).
                from datetime import timedelta
                props = wb.find('s:workbookPr', NS)
                base = datetime(1904, 1, 1) if props is not None and props.attrib.get('date1904') in ('1', 'true') else datetime(1899, 12, 30)
                if (base + timedelta(days=float(value.text))).date() == date:
                    matches.append((s.attrib['name'], target, col, cells))
        require(len(matches) == 1, 'Could not identify one date column. Open and save the workbook in Excel to refresh date formulas, and check the month.')
        return matches[0]

def plan(report, config):
    rows = {r: (v, 'Manager DAY statistic') for r, v in report['stats'].items()}
    seen = set()
    for rule in config['rules']:
        group, codes = rule['group'], rule['codes']
        values = []
        for code in codes:
            require((group, code) not in seen, f'Duplicate mapping for {group}/{code}')
            seen.add((group, code))
            if code in report['groups'][group]:
                label, value = report['groups'][group][code]
                values.append((code, label, value))
        if not values: continue
        value = sum((v for _, _, v in values), Decimal(0)) * Decimal(rule.get('sign', 1))
        row = rule['row']
        require(row not in rows, f'Duplicate destination row {row}')
        rows[row] = (value, ' + '.join(f'{c} {l} ({v})' for c, l, v in values))
    for group, entries in report['groups'].items():
        for code in entries:
            if group == 'daily' and code in report['payment_codes']: continue
            name = config.get('catalogue', {}).get(code, {}).get('name', entries[code][0])
            require((group, code) in seen, f'Unmapped transaction {group}/{code} ({name}); confirm its banking destination before importing.')
    rows[313] = (report['transfer'], 'Deposit ledger: Deposits Transfered at Check-In')
    reconciles(sum((v for r, (v, _) in rows.items() if 13 <= r <= 278), Decimal(0)), report['revenue'], 'Mapped revenue')
    require(sum((v for r, (v, _) in rows.items() if 282 <= r <= 286), Decimal(0)) == report['nonrevenue'], 'Mapped non-revenue does not reconcile.')
    require(sum((v for r, (v, _) in rows.items() if 290 <= r <= 312), Decimal(0)) == -report['payments'], 'Mapped payments do not reconcile.')
    return rows

def check_ledgers(report, col, cells, rows):
    """Evaluate the template's three ledger formulas from cached opening balances."""
    previous = chr(ord(col) - 1)
    errors, checks, opening = [], [], {}
    penny = Decimal('0.01')
    for row in (320, 321, 322):
        cell = cells.get(f'{previous}{row}')
        value = cell.find('s:v', NS) if cell is not None else None
        require(value is not None and value.text is not None, f'Missing cached opening balance {previous}{row}; recalculate and save in Excel.')
        opening[row] = Decimal(value.text).quantize(penny)
        if opening[row] != report['opening'][row]:
            errors.append(f'{previous}{row} opening balance {opening[row]} differs from report {report["opening"][row]}.')
    expected_formulas = {
        320: f'={previous}320+{col}288-SUM({col}290:{col}300)-{col}313+{col}314',
        321: f'={previous}321+{col}299-SUM({col}301:{col}306)',
        322: f'=+{previous}322+{col}313-SUM({col}307:{col}312)',
    }
    for row, expected in expected_formulas.items():
        f = cells[f'{col}{row}'].find('s:f', NS)
        require(f is not None and f.text == expected[1:], f'Changed ledger formula {col}{row}; review template.')
    def total(start, end):
        return sum((v for r, (v, _) in rows.items() if start <= r <= end), Decimal(0))
    gross = total(13, 278) + total(282, 286)
    calculated = {
        320: opening[320] + gross - total(290, 300) - total(313, 313) + total(314, 314),
        321: opening[321] + total(299, 299) - total(301, 306),
        322: opening[322] + total(313, 313) - total(307, 312),
    }
    difference = calculated[320] - report['balances'][320]
    # A positive Differences entry reduces the guest ledger closing balance.
    if not errors and difference and abs(difference) <= Decimal('0.02'):
        rows[300] = (difference, 'Rounding correction: calculated guest closing balance minus report closing balance')
        calculated[320] -= difference
    for row, actual in calculated.items():
        expected = report['balances'][row]
        status = 'PASS' if actual == expected else 'FAIL'
        checks.append((row, status, actual, expected))
        if status == 'FAIL':
            errors.append(f'{col}{row} calculated {actual}; report {expected}; difference {actual - expected}.')
    return checks, errors

def _write_copy(source, output, target, col, rows):
    require(source.resolve() != output.resolve(), 'Output must be a separate workbook.')
    require(not output.exists(), f'Output already exists: {output}')
    with zipfile.ZipFile(source) as src, zipfile.ZipFile(output, 'x') as dst:
        for entry in src.infolist():
            data = src.read(entry.filename)
            if entry.filename == target:
                text = data.decode('utf-8')
                for row, (value, _) in rows.items():
                    ref = f'{col}{row}'
                    pattern = r'<c\b[^>]*\br="' + ref + r'"[^>]*?(?:/>|>.*?</c>)'
                    match = re.search(pattern, text, re.S)
                    require(match is not None, f'Missing template cell {ref}; changed workbook layout.')
                    old = match[0]
                    require('<f' not in old, f'Refusing to replace formula {ref}')
                    opening = old.split('>')[0].rstrip('/') + '>'
                    opening = re.sub(r'\s+t="[^"]*"', '', opening)
                    replacement = opening + '<v>' + str(value) + '</v></c>'
                    text = text[:match.start()] + replacement + text[match.end():]
                data = text.encode('utf-8')
            dst.writestr(entry, data)

def write_copy(source, output, target, col, rows):
    # Prepare an entire copy before exposing the final filename.
    import tempfile
    require(source.resolve() != output.resolve(), 'Output must be a separate workbook.')
    require(not output.exists(), f'Output already exists: {output}')
    with tempfile.TemporaryDirectory(dir=output.parent) as staging:
        temp = Path(staging) / 'prepared.xlsx'
        _write_copy(source, temp, target, col, rows)
        # Exclusive creation prevents accidental replacement of an existing output.
        try:
            with output.open('xb') as final, temp.open('rb') as prepared:
                import shutil
                shutil.copyfileobj(prepared, final)
        except FileExistsError:
            raise ValueError(f'Output already exists: {output}')
        except OSError:
            output.unlink(missing_ok=True)
            raise

def run():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('report', type=Path)
    ap.add_argument('workbook', type=Path)
    ap.add_argument('--mapping', type=Path, default=Path(__file__).with_name('mapping.json'))
    ap.add_argument('--output-dir', type=Path, default=Path.cwd())
    ap.add_argument('--apply', action='store_true', help='Write a separate workbook after confirming mapping.json')
    args = ap.parse_args()
    config = json.loads(args.mapping.read_text())
    report = parse_report(args.report)
    sheet, target, col, cells = workbook_target(args.workbook, report['date'])
    rows = plan(report, config)
    ledger_checks, ledger_errors = check_ledgers(report, col, cells, rows)
    for row in rows:
        c = cells.get(f'{col}{row}')
        require(c is not None, f'Missing destination cell {col}{row}')
        require(c.find('s:f', NS) is None and c.find('s:is', NS) is None and c.attrib.get('t') not in ('s','inlineStr'), f'Destination {col}{row} has a formula or text; review required.')
        v = c.find('s:v', NS)
        require(v is None or v.text in (None, ''), f'Destination {col}{row} already contains a value; refusing to overwrite.')
    args.output_dir.mkdir(parents=True, exist_ok=True)
    review = args.output_dir / f'review-{report["date"]}.csv'
    require(not review.exists(), f'Review already exists: {review}; choose a new output directory.')
    with review.open('x', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(['Status', 'Date', 'Sheet', 'Cell', 'Amount', 'Source'])
        for row, (value, source) in sorted(rows.items()):
            writer.writerow(['CONFIRMED' if config.get('confirmed') else 'PROPOSED - confirm grouping', report['date'], sheet, f'{col}{row}', str(value), source])
        for note in ROUNDING_NOTES:
            writer.writerow(['ROUNDING NOTICE', report['date'], sheet, '', '', note])
        for row, status, actual, expected in ledger_checks:
            writer.writerow([status + ' LEDGER BALANCE', report['date'], sheet, f'{col}{row}', str(actual), f'Report closing balance {expected}; difference {actual - expected}'])
        for error in ledger_errors:
            writer.writerow(['BLOCKER', report['date'], sheet, '', '', error])
    print(f'Review created: {review}\nDate: {report["date"]}; destination: {sheet}, column {col}\nRevenue, non-revenue, payments and ledger breakdown checks passed.')
    for note in ROUNDING_NOTES:
        print('Rounding notice: ' + note)
    for row, status, actual, expected in ledger_checks:
        print(f'{status}: {col}{row} calculated {actual}, report {expected}')
    if args.apply:
        require(not ledger_errors, 'Workbook not written: ' + ' '.join(ledger_errors))
        require(config.get('confirmed') is True, 'Grouping is not confirmed. Review mapping.json with your banking rules before setting confirmed to true.')
        output = args.output_dir / f'{args.workbook.stem}-updated-{report["date"]}.xlsx'
        write_copy(args.workbook, output, target, col, rows)
        print(f'Workbook copy created: {output}\nOpen in Excel, calculate the workbook, and compare closing balances in the review CSV.')
    else:
        print('Preview only: no workbook changed.')

if __name__ == '__main__':
    try:
        run()
    except (ValueError, OSError, KeyError, zipfile.BadZipFile) as e:
        print(f'Import stopped: {e}', file=sys.stderr)
        sys.exit(1)
