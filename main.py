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

    # Computing the max width
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    # Changing the font to ensure the max_width get calculated correctly
    pdf.set_font('Times', style='B', size=12)
    max_widths = [pdf.get_string_width(column.replace('_purchased', '')
                                       .replace('_', ' ').title())
                  for column in df.columns]
    for index_1, rows in df.iterrows():
        for index_2, value in enumerate(rows):
            width = pdf.get_string_width(str(value))
            if width > max_widths[index_2]:
                max_widths[index_2] = width

    # Table header
    for col_name, max_width in zip(df.columns, max_widths):
        col_name_formatted = (col_name.replace('_purchased', '')
                              .replace('_', ' ').title())
        pdf.set_font('Times', style='B', size=12)
        pdf.cell(w=max_width + 8, h=10, txt=str(col_name_formatted), border=1)
    pdf.ln()

    # Table rows
    for index_1, row in df.iterrows():
        for index_2, value in enumerate(row):
            pdf.set_font('Times', size=11)
            pdf.cell(w=max_widths[index_2] + 8, h=10, txt=str(value), border=1)
        pdf.ln()

    # Additional row for the total
    total_iterations = len(df.columns)
    total_invoice = sum(df['total_price'])
    for index, (col_name, max_width) in enumerate(zip(df.columns, max_widths)):
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
