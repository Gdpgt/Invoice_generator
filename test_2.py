import fpdf
import pandas as pd

pdf = fpdf.FPDF(orientation='P', unit='mm', format='A4')
pdf.add_page()

# Table
xls = pd.read_excel("10001-2023.1.18.xlsx")
pdf.set_font('Times', style='B', size=13)

# Calculate maximum cell widths based on header and data
max_widths = [pdf.get_string_width(str(col_name)) for col_name in xls.columns] # [21.660202777777773, 28.53478333333333, 37.711591666666656, 29.043841666666662, 21.14197222222222]
for index, row in xls.iterrows():
    for col_index, value in enumerate(row):
        width = pdf.get_string_width(str(value))
        if width > max_widths[col_index]:
            max_widths[col_index] = width

# Output table header with calculated cell widths
for col_name, max_width in zip(xls.columns, max_widths):
    pdf.cell(w=max_width + 6, h=12, txt=str(col_name), border=1)
pdf.ln()

# Set font for the data rows
pdf.set_font('Times', size=13)

# Output table rows with calculated cell widths
for index, row in xls.iterrows():
    for col_index, value in enumerate(row):
        pdf.cell(w=max_widths[col_index] + 6, h=12, txt=str(value), border=1)
    pdf.ln(12)

pdf.output('Invoice.pdf')