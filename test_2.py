import pandas as pd
from fpdf import FPDF

pdf = FPDF(orientation='P', unit='mm', format='A4')
pdf.add_page()

df = pd.read_excel('Excel_orders/10001-2023.1.18.xlsx', sheet_name='Sheet 1')
pdf.set_font('Times', style='B', size=18)

max_widths = [pdf.get_string_width(column.replace('_purchased', '')
                                   .replace('_', ' ').title())
              for column in df.columns]
print(max_widths)