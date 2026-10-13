"""Regression checks for posting classification and safe mapping."""
import json
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import patch
import banking_import as importer

DAILY = '''Station Hotel Aberdeen 12/10/26
Night Audit Trial Balance
Revenue
7102 Aberdeen Levy Fee - 20% VAT 10.00
Revenue Total 10.00
Non Revenue
7001 VAT@5% 0.50
Non Revenue Total 0.50
Payment
9006 Maestro - 10.50
Payment Total - 10.50
Balance Brought Forward 0.00
Guest Ledger
'''
LEDGERS = '''Station Hotel Aberdeen 12/10/26
Night Audit Trial Balance
Payments - 10.50
9006 Maestro - 10.50
Balance Today 0.00
AR Ledger
Balance Yesterday 0.00
Payments 0.00
Balance Today 0.00
Deposit Ledger
Balance Yesterday 0.00
Deposits Transfered at Check-In 0.00
Payments 0.00
Balance Today 0.00
Package Ledger
'''
MANAGER = '''Station Hotel Aberdeen 12/10/26
Manager Flash Report
Rooms Occupied 1 1 1
Total In-House Persons 2 2 2
Total Revenue 10.00 10.00 10.00
'''

class ImportTests(unittest.TestCase):
    def setUp(self):
        self.config = json.loads(Path(__file__).with_name('mapping.json').read_text())
        importer.ROUNDING_NOTES.clear()

    def report(self, texts):
        pdf = SimpleNamespace(pages=[SimpleNamespace(extract_text=lambda t=t: t) for t in texts])
        with patch.object(importer.pdfplumber, 'open') as mocked:
            mocked.return_value.__enter__.return_value = pdf
            return importer.parse_report('synthetic.pdf')

    def test_high_numbered_revenue_and_additional_card_type(self):
        report = self.report([DAILY, LEDGERS, MANAGER])
        rows = importer.plan(report, self.config)
        self.assertEqual(rows[277][0], importer.Decimal('10.00'))
        self.assertEqual(rows[286][0], importer.Decimal('0.50'))
        self.assertEqual(rows[292][0], importer.Decimal('10.50'))

    def test_extra_trial_balance_page(self):
        report = self.report([DAILY, 'Station Hotel Aberdeen 12/10/26\nNight Audit Trial Balance\n', LEDGERS, MANAGER])
        self.assertEqual(report['revenue'], importer.Decimal('10.00'))

    def test_catalogued_but_unconfirmed_code_stops(self):
        report = self.report([DAILY.replace('7102 Aberdeen Levy Fee - 20% VAT', '7400 Package Profit'), LEDGERS, MANAGER])
        with self.assertRaisesRegex(ValueError, r'7400 \(Package Profit\)'):
            importer.plan(report, self.config)

    def test_department_rules_and_signed_adjustments(self):
        rules = {(r['group'], code): r['row'] for r in self.config['rules'] for code in r['codes']}
        for code, row in [('3101', 23), ('3500', 23), ('4700', 53), ('4930', 53), ('5001', 39), ('4000', 275), ('1066', 13), ('4051', 13), ('2020', 30), ('3023', 37), ('3054', 34), ('4900', 40), ('2700', 43), ('2033', 46), ('3810', 51), ('4950', 54)]:
            self.assertEqual(rules['daily', code], row)
        report = self.report([DAILY.replace('7102 Aberdeen Levy Fee - 20% VAT 10.00', '3053 Nor1 Bottle of White Wine 15.00\n3054 Nor1 ADJ Bottle of White Wine - 5.00'), LEDGERS, MANAGER])
        rows = importer.plan(report, self.config)
        self.assertEqual(rows[34][0], importer.Decimal('10.00'))

    def test_equal_opposite_correction_and_unrelated_blocker(self):
        import xml.etree.ElementTree as ET
        report = self.report([DAILY, LEDGERS, MANAGER])
        cells = {}
        for row in (320, 321, 322):
            cells[f'H{row}'] = ET.fromstring('<c xmlns="' + importer.NS['s'] + '"><v>0</v></c>')
        formulas = {320: 'H320+I288-SUM(I290:I300)-I313+I314', 321: 'H321+I299-SUM(I301:I306)', 322: '+H322+I313-SUM(I307:I312)'}
        for row, formula in formulas.items():
            cells[f'I{row}'] = ET.fromstring('<c xmlns="' + importer.NS['s'] + '"><f>' + formula + '</f></c>')
        for sign in (1, -1):
            report['balances'][320] = importer.Decimal(10 * sign)
            report['balances'][321] = importer.Decimal(-10 * sign)
            rows = importer.plan(report, self.config)
            checks, errors = importer.check_ledgers(report, 'I', cells, rows)
            self.assertFalse(errors)
            self.assertEqual(rows[302][0], importer.Decimal(10 * sign))
            self.assertEqual(rows[292][0], importer.Decimal('10.50') - 10 * sign)
            self.assertEqual(report['corrections']['rows'], [292, 302])
        for transfer, remainder in [(importer.Decimal('334.00'), importer.Decimal('0.02')),
                                     (importer.Decimal('-334.00'), importer.Decimal('-0.02'))]:
            # Zero calculated balances produce the same discrepancies as the error.
            report['balances'][320] = transfer - remainder
            report['balances'][321] = -transfer
            report['split_exceptions'] = ['Ledger Payment adjustment']
            rows = importer.plan(report, self.config)
            checks, errors = importer.check_ledgers(report, 'I', cells, rows)
            self.assertFalse(errors)
            self.assertEqual(rows[292][0], importer.Decimal('10.50') - transfer)
            self.assertEqual(rows[302][0], transfer)
            self.assertEqual(rows[300][0], remainder)
            self.assertTrue(all(status == 'PASS' for _, status, _, _ in checks))
        report['balances'][320] = importer.Decimal('333.97')
        report['balances'][321] = importer.Decimal('-334')
        rows = importer.plan(report, self.config)
        checks, errors = importer.check_ledgers(report, 'I', cells, rows)
        self.assertTrue(errors)  # 3p is outside the authorised rounding tolerance.
        self.assertIsNone(report['corrections'])
        report['balances'][322] = importer.Decimal('1')
        rows = importer.plan(report, self.config)
        checks, errors = importer.check_ledgers(report, 'I', cells, rows)
        self.assertTrue(errors)
        self.assertIsNone(report['corrections'])

    def test_explanation_appends_existing_text(self):
        import io, zipfile
        data = io.BytesIO()
        with zipfile.ZipFile(data, 'w') as z:
            z.writestr('xl/sharedStrings.xml', '<sst xmlns="' + importer.NS['s'] + '"><si><t>Existing explanation &amp; details</t></si></sst>')
        with zipfile.ZipFile(data) as z:
            text = '<row><c r="L336" s="1" t="s"><v>0</v></c></row>'
            result = importer.append_explanation(z, text, 'L336', 'New correction')
            self.assertIn('Existing explanation &amp; details\nNew correction', result)

    def test_export_requests_excel_recalculation(self):
        import xml.etree.ElementTree as ET
        for calc in ('<calcPr calcId="191029" calcMode="manual"/>', ''):
            raw = ('<workbook xmlns="' + importer.NS['s'] + '"><sheets/>' + calc + '<extLst/></workbook>').encode()
            result = ET.fromstring(importer.request_recalculation(raw))
            props = result.find('s:calcPr', importer.NS)
            self.assertEqual(props.attrib['calcMode'], 'auto')
            self.assertEqual(props.attrib['fullCalcOnLoad'], '1')
            self.assertEqual(props.attrib['forceFullCalc'], '1')

    def test_mapping_is_unique_and_codes_exist(self):
        seen = set()
        for rule in self.config['rules']:
            for code in rule['codes']:
                self.assertIn(code, self.config['catalogue'])
                key = rule['group'], code
                self.assertNotIn(key, seen)
                seen.add(key)
        self.assertEqual(len(self.config['catalogue']), 243)

if __name__ == '__main__':
    unittest.main()
