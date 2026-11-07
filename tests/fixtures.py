from pathlib import Path
from docx import Document
from docx.shared import Inches
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.image import Image
from PIL import Image as PILImage, ImageDraw
import zipfile

out=Path('test-results');out.mkdir(exist_ok=True)
im=PILImage.new('RGBA',(500,130),(255,255,255,0)); d=ImageDraw.Draw(im);d.line([(25,90),(85,20),(110,110),(175,35),(225,85),(400,40)],fill='#123456',width=7);im.save(out/'signature.png')
doc=Document();doc.add_heading('Document de contrôle',0);doc.add_paragraph('Bonjour, voici le document à signer.');table=doc.add_table(rows=1,cols=2);table.cell(0,0).text='Visa du directeur';table.cell(0,1).text='Signature ici';doc.add_paragraph('Fin du document.');doc.sections[0].header.paragraphs[0].text='En-tête conservé';doc.save(out/'source.docx')
w=Workbook();s=w.active;s.title='Budget';s.append(['Poste','Montant']);s.append(['A',100]);s.append(['B',200]);s['B4']='=SUM(B2:B3)';s['B4'].number_format='#,##0.00';s.column_dimensions['A'].width=28;s.merge_cells('A6:C6');s['A6']='Zone de signature';s.print_area='A1:H30';s.print_options.horizontalCentered=True
chart=BarChart();chart.add_data(Reference(s,min_col=2,min_row=1,max_row=3),titles_from_data=True);s.add_chart(chart,'J2');s.add_image(Image(str(out/'signature.png')),'J20');s.row_dimensions[40].hidden=True
hidden=w.create_sheet('Archives');hidden.sheet_state='hidden';protected=w.create_sheet('Protégée');protected.protection.sheet=True
w.save(out/'source.xlsx')
# La protection et les signatures existantes ne doivent pas être contournées.
with zipfile.ZipFile(out/'source.docx') as z,zipfile.ZipFile(out/'signed-already.docx','w') as dst:
    for n in z.namelist():dst.writestr(n,z.read(n))
    dst.writestr('_xmlsignatures/sig1.xml','<Signature/>')
