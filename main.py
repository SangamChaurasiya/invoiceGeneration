import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    filename = Path(filepath).stem
    invoice_no, date = filename.split('-')

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_no}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date {date}", ln=1)
    pdf.cell(w=0, h=8, ln=1)

    amount = 0
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add a header
    columns = list(df.columns)

    pdf.set_font(family="Times", style="B", size=12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=8, txt=f"{columns[0].replace('_', ' ').title()}", border=1)
    pdf.cell(w=70, h=8, txt=f"{columns[1].replace('_', ' ').title()}", border=1)
    pdf.cell(w=40, h=8, txt=f"{columns[2].replace('_', ' ').title()}", border=1)
    pdf.cell(w=30, h=8, txt=f"{columns[3].replace('_', ' ').title()}", border=1)
    pdf.cell(w=22, h=8, txt=f"{columns[4].replace('_', ' ').title()}", border=1, ln=1)

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=12)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(w=30, h=8, txt=f"{row['product_id']}", border=1)
        pdf.cell(w=70, h=8, txt=f"{row['product_name']}", border=1)
        pdf.cell(w=40, h=8, txt=f"{row['amount_purchased']}", border=1)
        pdf.cell(w=30, h=8, txt=f"{row['price_per_unit']}", border=1)
        pdf.cell(w=22, h=8, txt=f"{row['total_price']}", border=1, ln=1)
        amount += row['total_price']

    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=70, h=8, border=1)
    pdf.cell(w=40, h=8, border=1)
    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=22, h=8, txt=f"{amount}", border=1, ln=1)

    pdf.set_font(family="Times", style="B", size=12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=8, ln=1)
    pdf.cell(w=100, h=8, txt=f"The total due amount is {amount} Euros.", ln=1)
    pdf.cell(w=22, h=8, txt="PythonHow")
    pdf.image('pythonhow.png', w=10, h=8)
    pdf.output(f"PDFs/{filename}.pdf")
