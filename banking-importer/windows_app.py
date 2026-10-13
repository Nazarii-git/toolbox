"""Windows file-picker launcher; all processing stays on your computer."""
from pathlib import Path
import subprocess
import sys
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

root = tk.Tk()
root.withdraw()
report = filedialog.askopenfilename(title='Select daily night audit PDF', filetypes=[('PDF report', '*.pdf')])
if not report: sys.exit(0)
workbook = filedialog.askopenfilename(title='Select monthly banking workbook', filetypes=[('Excel workbook', '*.xlsx')])
if not workbook: sys.exit(0)
folder = Path(workbook).parent / ('Banking import ' + datetime.now().strftime('%Y-%m-%d %H%M%S'))
result = subprocess.run([sys.executable, str(Path(__file__).with_name('banking_import.py')), report, workbook, '--output-dir', str(folder), '--apply'], capture_output=True, text=True)
messagebox.showinfo('Import complete' if result.returncode == 0 else 'Import stopped', result.stdout + '\n' + result.stderr)
root.destroy()
