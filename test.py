import fpdf
import pandas as pd

pdf = fpdf.FPDF(orientation='P', unit='mm', format='A4')
pdf.add_page()

# Table header
xls = pd.read_excel("10001-2023.1.18.xlsx")
for col_name in xls.columns:
    pdf.set_font('Times', style='B', size=13)
    pdf.cell(w=39, h=12, txt=str(col_name), border=1)
pdf.ln()

# Table rows
for index, row in xls.iterrows():
    for value in row:
        pdf.set_font('Times', size=13)
        pdf.cell(w=39, h=12, txt=str(value), border=1)
    pdf.ln(12)

pdf.output('Invoice.pdf')
