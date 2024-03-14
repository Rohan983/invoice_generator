import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
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

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    pdf.set_font(family="Times", size=10, style="B")
    column_names = list(df.columns)
    col_size = [30, 70, 32, 30, 28]
    for size, column_name in zip(col_size, column_names):
        pdf.cell(w=size, h=8, txt=column_name.replace("_", " ").title(), border=1)
    pdf.ln(8)

    for index, rows in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(rows["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(rows["product_name"]), border=1)
        pdf.cell(w=32, h=8, txt=str(rows["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(rows["price_per_unit"]), border=1)
        pdf.cell(w=28, h=8, txt=str(rows["total_price"]), border=1, ln=1)

    pdf.output(f"PDFs/{filename}.pdf")
