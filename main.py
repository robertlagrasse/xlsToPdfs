# ingest multiple xls file
# build a pdf invoice for each

import pandas
import glob
from fpdf import FPDF
from pathlib import Path

paths = glob.glob('invoices/*.xlsx')

for path in paths:
    df = pandas.read_excel(path, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="Letter")
    pdf.add_page()
    filename = Path(path).stem
    invoice, date = filename.split('-')
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice Number:{invoice}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date:{date}", ln=1)

    # Build the header
    columns = df.columns
    columns = [item.replace('_',' ').title() for item in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10, style="")
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['amount_purchased']), border=1, align='R')
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1, align='R')
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, align='R', ln=1)

    total = df['total_price'].sum()
    pdf.set_font(family="Times", size=10, style="")
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1, align='R')
    pdf.cell(w=30, h=8, txt="", border=1, align='R')
    pdf.cell(w=30, h=8, txt=str(total), border=1, align='R', ln=1)

    pdf.set_font(family="Times", size=18, style="")
    pdf.cell(w=0, h=18, txt=f"Total Amount Due: ${total}", border=0, ln=1)
    pdf.cell(w=30, h=30, txt=f"Pay up", border=0)
    pdf.image("images/17.jpg", w=30)


    pdf.output(f"PDFs/{filename}.pdf")
