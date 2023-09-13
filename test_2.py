columns = list(df.columns)
pdf.set_font('Times', style='B', size=12)
pdf.cell(w=30, h=8, txt=columns[0], border=1)
pdf.cell(w=30, h=8, txt=columns[1], border=1)