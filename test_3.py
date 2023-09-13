import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('Excel_orders/*.xlsx')

for filepath in filepaths:
    filename = Path(filepath).stem
    inv_nbr, inv_date = filename.split('-')

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    # Header
    pdf.set_font('Times', style='B', size=18)
    pdf.cell(w=0, h=12, txt=f"Invoice n° {inv_nbr}", align='L', ln=1)
    pdf.cell(w=0, h=10, txt=f"Date {inv_date}", align='L', ln=1)
    pdf.ln(10)

    # Computing the max widths
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    # Changing the font to ensure the max_widths get calculated correctly
    pdf.set_font('Times', style='B', size=12)
    # Applying the correct modifications to header values for the same reason
    max_widths = [pdf.get_string_width(column.replace('_purchased', '')
                                       .replace('_', ' ').title())
                  for column in df.columns]
    for index_1, rows in df.iterrows():
        for index_2, value in enumerate(rows):
            width = pdf.get_string_width(str(value))
            if width > max_widths[index_2]:
                max_widths[index_2] = width

    # Table header
    pdf.set_font('Times', style='B', size=12)
    pdf.cell(w=max_widths[0] + 8, h=8, txt=df.columns[0], border=1)
    pdf.cell(w=max_widths[1] + 8, h=8, txt=df.columns[1], border=1)
    pdf.cell(w=max_widths[2] + 8, h=8, txt=df.columns[2], border=1)
    pdf.cell(w=max_widths[3] + 8, h=8, txt=df.columns[3], border=1)
    pdf.cell(w=max_widths[4] + 8, h=8, txt=df.columns[4], border=1)
    pdf.ln()

    # Table rows
    for index_1, row in df.iterrows():
        for index_2, value in enumerate(row):
            pdf.set_font('Times', size=11)
            pdf.cell(w=max_widths[index_2] + 8, h=10, txt=str(value), border=1)
        pdf.ln()

    # Additional row for the total
    total_iterations = len(max_widths)
    total_invoice = sum(df['total_price'])
    for index, max_width in enumerate(max_widths):
        if index != total_iterations - 1:
            pdf.cell(w=max_width + 8, h=10, txt='', border=1)
        else:
            pdf.cell(w=max_width + 8, h=10, txt=str(total_invoice), border=1)
    pdf.ln(20)

    # Total sentence
    pdf.set_font('Times', style='B', size=13)
    pdf.cell(w=0, h=12, txt=f"The total due amount is {total_invoice} euros.",
             align='L')

    pdf.output(f"PDF_Invoices/Invoice_{filename}.pdf")
