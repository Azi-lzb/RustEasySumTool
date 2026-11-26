"""Independent format retention and repeated-log verification using openpyxl."""
from copy import copy
import os
import shutil
from hashlib import sha256
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parent))
from parity import aggregate
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

MAIN = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

def archive(path):
    with ZipFile(path) as file:
        assert file.testzip() is None
        return {name: file.read(name) for name in file.namelist()}

def run():
    with TemporaryDirectory(prefix='rust_format_') as directory:
        root = Path(directory)
        template, source, output = [root / name for name in ['template.xlsx','source.xlsx','result.xlsx']]
        book = Workbook()
        sheet = book.active
        sheet.title = '业务表'
        for cell, text in {'K4':'MAX(A1)', 'K5':'SUMPRODUCT(A1,B1)', 'K6':'RATIO(A1,B1)',
                           'K7':'WAVG(A1:A3,B1:B3)', 'K8':'SUM(,,)'}.items():
            sheet[cell] = text
            sheet[cell].font = Font(name='微软雅黑', size=13, color='FF135724', bold=True)
            sheet[cell].fill = PatternFill('solid',fgColor='FFCCEEFF')
            sheet[cell].border = Border(bottom=Side(style='thin',color='FF123456'))
            sheet[cell].alignment = Alignment(horizontal='right',indent=2,wrap_text=True,readingOrder=2)
            sheet[cell].number_format = '#,##0.0000'
            sheet[cell].protection = Protection(locked=False,hidden=True)
        sheet['K4'].comment = Comment('原批注𠮷\n保留作者','原作者')
        sheet['K4'].comment.width = 800
        sheet['K4'].comment.height = 400
        sheet['D8'] = '普通批注单元格'
        sheet['D8'].comment = Comment('保持不变','另一作者')
        sheet['H2'] = '链接'
        sheet['H2'].hyperlink = 'https://example.com'
        sheet.merge_cells('B2:C3')
        sheet.column_dimensions['K'].width = 25
        sheet.row_dimensions[4].height = 37
        sheet.print_title_rows = '1:3'
        sheet.print_title_cols = 'A:B'
        sheet.print_area = 'A1:K20'
        sheet.page_setup.orientation = 'landscape'
        sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth = 1
        sheet.page_margins.left = 0.3
        sheet.freeze_panes = 'D4'
        sheet.auto_filter.ref = 'A1:K20'
        sheet.protection.sheet = True
        sheet.protection.set_password('test')
        validation = DataValidation(type='whole',operator='between',formula1='1',formula2='10')
        validation.promptTitle = '提示'
        validation.prompt = '填写1到10'
        sheet.add_data_validation(validation)
        validation.add('E5:E10')
        sheet.conditional_formatting.add('F5:F10',CellIsRule(operator='greaterThan',formula=['3'],fill=PatternFill('solid',fgColor='FFFF0000')))
        for row in [('名称','金额'),('甲',1),('乙',2)]:
            sheet.append(row)
        # Keep the table away from instruction and merged cells.
        for row, values in enumerate([('名称','金额'),('甲',1),('乙',2)],25):
            for column, value in enumerate(values,1): sheet.cell(row,column,value)
        table=Table(displayName='BusinessTable',ref='A25:B27')
        table.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True)
        sheet.add_table(table)
        book.create_sheet('附页')['B2']='固定说明'
        book.save(template)
        source_book=Workbook();source_book.active.title='业务表'
        for cell,value in {'A1':10,'A2':20,'A3':30,'B1':1,'B2':2,'B3':3,'K8':2}.items(): source_book.active[cell]=value
        source_book.save(source)
        before={path:sha256(path.read_bytes()).digest() for path in [template,source]}
        result=aggregate(template,[source],output)
        if os.environ.get('RUST_FORMAT_COPY'):
            shutil.copy2(output,os.environ['RUST_FORMAT_COPY'])
        assert result.function_counts == {'MAX':1,'SUMPRODUCT':1,'RATIO':1,'WAVG':1,'SUM':1}
        old,new=load_workbook(template),load_workbook(output)
        a,b=old['业务表'],new['业务表']
        for cell in ['K4','K5','K6','K7','K8']:
            for name in ['font','fill','border','alignment','number_format','protection']:
                assert copy(getattr(a[cell],name))==copy(getattr(b[cell],name)),(cell,name)
        for name in ['merged_cells','print_area','print_title_rows','print_title_cols','page_setup','page_margins','freeze_panes','auto_filter','data_validations','protection']:
            assert str(getattr(a,name))==str(getattr(b,name)),name
        assert list(a.conditional_formatting)==list(b.conditional_formatting)
        assert list(a.tables)==list(b.tables)
        assert b['K4'].comment.text.startswith(a['K4'].comment.text+'\n\n')
        assert b['K4'].comment.author=='原作者'
        assert b['D8'].comment.text==a['D8'].comment.text and b['D8'].comment.author==a['D8'].comment.author
        assert b['H2'].hyperlink.target==a['H2'].hyperlink.target
        assert b['K8'].value==6
        assert a.row_dimensions[4].height==b.row_dimensions[4].height
        assert a.column_dimensions['K'].width==b.column_dimensions['K'].width
        old_xml,new_xml=archive(template),archive(output)
        vml=[content for name,content in new_xml.items() if name.startswith('xl/rust-summary/') and name.endswith('.vml')]
        assert vml
        dimensions=[]
        for content in vml:
            tree=ET.fromstring(content)
            for shape in tree.iter('{urn:schemas-microsoft-com:vml}shape'):
                client=shape.find('{urn:schemas-microsoft-com:office:excel}ClientData')
                if client is not None and client.findtext('{urn:schemas-microsoft-com:office:excel}Row')=='3' and client.findtext('{urn:schemas-microsoft-com:office:excel}Column')=='10':
                    dimensions.append(dict(piece.strip().split(':',1) for piece in shape.attrib['style'].split(';') if ':' in piece))
        assert dimensions,vml
        assert float(dimensions[0]['width'].removesuffix('pt'))>=600
        assert float(dimensions[0]['height'].removesuffix('pt'))>=300
        allowed={'xl/worksheets/sheet1.xml','xl/worksheets/_rels/sheet1.xml.rels','xl/styles.xml','xl/workbook.xml','xl/_rels/workbook.xml.rels','[Content_Types].xml'}
        for name,content in old_xml.items():
            if name not in allowed: assert new_xml[name]==content, name
        style_old,style_new=[ET.fromstring(x['xl/styles.xml']) for x in [old_xml,new_xml]]
        for group in style_old:
            other=style_new.find(group.tag)
            assert other is not None
            for index,entry in enumerate(group): assert ET.tostring(entry)==ET.tostring(other[index]),group.tag
        # Existing history must survive a second run and the log remains last.
        next_template=root/'next-template.xlsx'
        new['业务表']['K4']='MAX(A1)'
        new.save(next_template)
        old.close();new.close()
        saved=load_workbook(next_template)
        history=[tuple(cell.value for cell in row) for row in saved['运行日志']]
        saved.close()
        next_output=root/'second.xlsx'
        aggregate(next_template,[source],next_output)
        repeated=load_workbook(next_output)
        assert repeated.sheetnames==['业务表','附页','运行日志']
        repeated_history=[tuple(cell.value for cell in row) for row in repeated['运行日志'].iter_rows(max_row=len(history))]
        assert repeated_history==history,[(i,a,b) for i,(a,b) in enumerate(zip(history,repeated_history)) if a!=b][:3]
        repeated.close()
        assert before=={path:sha256(path.read_bytes()).digest() for path in before}
        print('Format retention passed: styles, indent, merges, dimensions, printing, validation formulas, conditional formatting, tables, hyperlinks, protection, comments and appended history')

if __name__=='__main__':run()
