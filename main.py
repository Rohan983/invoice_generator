import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Getting all xlsx files
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    # Extracting filename from filepath
    # filename = filepath.split('\\')[-1]   # This also works but using pathlib library is better
    filename = Path(filepath).stem
    invoice_no, date = filename.split("-")

    # Setting up PDF page format
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=False, margin=0)

    pdf.add_page()

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, ln=1, txt=f"Invoice No.: {invoice_no}")
    pdf.cell(w=50, h=8, ln=1, txt=f"Date: {date}")
    pdf.ln(1)

    # Reading xlsx file
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Adding Column names
    pdf.set_font(family="Times", size=10, style="B")
    column_names = list(df.columns)
    col_size = [30, 70, 32, 30, 28]
    for size, column_name in zip(col_size, column_names):
        pdf.cell(w=size, h=8, txt=column_name.replace("_", " ").title(), border=1)
    pdf.ln(8)

    # Adding Rows
    for index, rows in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        for size, column_name in zip(col_size, column_names):
            pdf.cell(w=size, h=8, txt=str(rows[column_name]), border=1)
        pdf.ln(8)

    # Calculating and printing total price
    tot_sum = str(df["total_price"].sum())
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=132, h=8, txt="")
    pdf.cell(w=30, h=8, txt="Total Bill", border=1)
    pdf.cell(w=28, h=8, txt=tot_sum, border=1, ln=1)

    # Printing Company name and logo
    pdf.ln(2)
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=60, h=8, txt="The Rohan More Company")
    pdf.image("companylogo.png", w=10)

    # Generating PDF
    pdf.output(f"PDFs/{filename}.pdf")
