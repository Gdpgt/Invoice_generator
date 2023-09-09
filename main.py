import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('Excel_orders/*.xlsx')

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")  # the sheet_name isn't necessary if there's only one sheet
    filename = Path(filepath).stem
    inv_nbr = filename[:5]  # Or filename.split('-')[0]
    inv_date = filename[6:15]  # Or filename.split('-')[1]
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    # Header
    pdf.set_font('Times', style='B', size=18)
    pdf.cell(w=0, h=12, txt=f"Invoice n° {inv_nbr}", align='L', ln=1)
    pdf.cell(w=0, h=10, txt=f"Date {inv_date}", align='L', ln=1)
    pdf.ln(10)

    # Computing the max width
    max_widths = [pdf.get_string_width(column) for column in df.columns]
    for index_1, rows in df.iterrows():
        for index_2, value in enumerate(rows):
            width = pdf.get_string_width(str(value))
            if width > max_widths[index_2]:
                max_widths[index_2] = width


    # Table header
    for col_name, max_width in zip(df.columns, max_widths):
        col_name_formatted = col_name.replace('_purchased', '').replace('_', ' ').title()
        pdf.set_font('Times', style='B', size=13)
        pdf.cell(w=max_width - 6, h=10, txt=str(col_name_formatted), border=1)
    pdf.ln()

    # Table rows
    for index_1, row in df.iterrows():
        for index_2, value in enumerate(row):
            pdf.set_font('Times', size=12)
            pdf.cell(w=max_widths[index_2] - 6, h=10, txt=str(value), border=1)
        pdf.ln()

    # Additional row for the total
    total_iterations = len(df.columns)
    total_invoice = sum(df['total_price'])
    for index, (col_name, max_width) in enumerate(zip(df.columns, max_widths)):
        if index != total_iterations - 1:
            pdf.cell(w=max_width - 6, h=10, txt='', border=1)
        else:
            pdf.cell(w=max_width - 6, h=10, txt=str(total_invoice), border=1)
    pdf.ln(20)

    # Total sentence
    pdf.set_font('Times', style='B', size=13)
    pdf.cell(w=0, h=12, txt=f"The total due amount is {total_invoice} euros.", align='L')

    pdf.output(f"PDF_Invoices/Invoice_{filename}.pdf")
