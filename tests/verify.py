from pathlib import Path
from docx import Document
from openpyxl import load_workbook
from lxml import etree
import zipfile
p=Path('test-results')
d=Document(p/'signed.docx')
assert len(d.inline_shapes)==2
assert d.tables[0].cell(0,1).text.startswith('Signature ici')
assert d.sections[0].header.paragraphs[0].text=='En-tête conservé'
w=load_workbook(p/'signed.xlsx');s=w['Budget']
assert s['B4'].value=='=SUM(B2:B3)'
assert s['B4'].number_format=='#,##0.00'
assert len(s._charts)==1 and len(s._images)==3
assert s.column_dimensions['A'].width==28
assert w['Archives'].sheet_state=='hidden'
assert 'A1:H30' in str(s.print_area).replace('$','')
assert w['Protégée'].protection.sheet
for ext, allowed in [('docx',{'word/document.xml','word/_rels/document.xml.rels','[Content_Types].xml'}),('xlsx',{'xl/worksheets/sheet1.xml','xl/worksheets/_rels/sheet1.xml.rels','xl/drawings/drawing1.xml','xl/drawings/_rels/drawing1.xml.rels','[Content_Types].xml'})]:
    with zipfile.ZipFile(p/f'source.{ext}') as before,zipfile.ZipFile(p/f'signed.{ext}') as after:
        for name in before.namelist():
            if name not in allowed:assert before.read(name)==after.read(name),f'Unexpected modification: {name}'
        for name in after.namelist():
            if name.endswith(('.xml','.rels')):etree.fromstring(after.read(name))
print('DOCX/XLSX reopened: signatures, formulas, charts, hidden sheets, print areas and untouched parts verified.')
