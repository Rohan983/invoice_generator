import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    # filename = filepath.split('\\')[-1]   #This also works but using pathlib library is better
    filename = Path(filepath).stem
    invoice_no, date = filename.split("-")

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=False, margin=0)
    pdf.add_page()

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, ln=1, txt=f"Invoice No.: {invoice_no}")
    pdf.cell(w=50, h=8, ln=1, txt=f"Date: {date}")
    pdf.ln(1)

    for index, rows in df.iterrows():
        print(rows)

    pdf.output(f"PDFs/{filename}.pdf")
