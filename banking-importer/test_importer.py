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
