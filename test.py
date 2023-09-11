from fpdf import FPDF
from glob import glob
from pathlib import Path

filepaths = glob('*.txt')
pdf = FPDF(orientation='P', unit='mm', format='A4')

for filepath in filepaths:
    pdf.add_page()
    title = Path(filepath).stem.capitalize()
    with open(filepath, 'r') as body:
        body = body.read()

    # Title
    pdf.set_font('Times', style='B', size=15)
    pdf.cell(w=0, h=15, txt=title, ln=1)

    # Body
    pdf.set_font('Times', size=12)
    pdf.multi_cell(w=0, h=10, txt=body)

pdf.output(f"Animal_descriptions.pdf")