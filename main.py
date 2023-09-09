import pandas as pd
import glob
import fpdf

files = glob.glob('*.xlsx')

inv_nbr = files[0][:5]
inv_date = files[0][6:15]
pdf = fpdf.FPDF(orientation='P', unit='mm', format='A4')
pdf.add_page()

# Header
pdf.set_font('Times', style='B', size=18)
pdf.cell(w=0, h=12, txt=f"Invoice n°{inv_nbr}", align='L', ln=1)
pdf.cell(w=0, h=10, txt=f"Date {inv_date}", align='L', ln=1)
pdf.ln(10)

# Computing the max width
xls = pd.read_excel(files[0])
max_widths = [pdf.get_string_width(column) for column in xls.columns]
for index_1, rows in xls.iterrows():
    for index_2, value in enumerate(rows):
        width = pdf.get_string_width(str(value))
        if width > max_widths[index_2]:
            max_widths[index_2] = width


# Table header
for col_name, max_width in zip(xls.columns, max_widths):
    col_name_formatted = col_name.replace('_purchased', '').replace('_', ' ').title()
    pdf.set_font('Times', style='B', size=13)
    pdf.cell(w=max_width - 6, h=10, txt=str(col_name_formatted), border=1)
pdf.ln()

# Table rows
for index_1, row in xls.iterrows():
    for index_2, value in enumerate(row):
        pdf.set_font('Times', size=12)
        pdf.cell(w=max_widths[index_2] - 6, h=10, txt=str(value), border=1)
    pdf.ln()

# Additional row for the total
total_iterations = len(xls.columns)
total_invoice = sum(xls['total_price'])
for index, (col_name, max_width) in enumerate(zip(xls.columns, max_widths)):
    if index != total_iterations - 1:
        pdf.cell(w=max_width - 6, h=10, txt='', border=1)
    else:
        pdf.cell(w=max_width - 6, h=10, txt=str(total_invoice), border=1)
pdf.ln(20)

# Total sentence
pdf.set_font('Times', style='B', size=13)
pdf.cell(w=0, h=12, txt=f"The total due amount is {total_invoice} euros.", align='L')

pdf.output('Invoice.pdf')
