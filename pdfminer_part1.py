from pdfminer.layout import LAParams
from pdfminer.converter import PDFPageAggregator, HTMLConverter
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LTTextBox
from string import punctuation
import sys
from io import BytesIO
from string import punctuation
import re
from openpyxl import load_workbook

# pdf path
document = open('/Users/anan/Desktop/pdf/Part1.pdf', 'rb')
# crate pdf manager
rsrcmgr = PDFResourceManager()

retster = BytesIO()
# Set parameters for analysis.
laparams = LAParams()
laparams.char_margin = 4  # 5 4best
laparams.line_margin = 3  # test:3 part:5
laparams.word_margin = 4
laparams.boxes_flow = 0.5

# Create a PDF page aggregator object
device = HTMLConverter(rsrcmgr, retster, codec='utf-8', laparams=laparams)
# create pdf interpreter
interpreter = PDFPageInterpreter(rsrcmgr, device)

content = ""
for page in PDFPage.get_pages(document):
    interpreter.process_page(page)
    layout = retster.getvalue()
    str = print(layout)

    for element in layout:
        if isinstance(element, LTTextBox):
            obj = element.get_text()
            content += obj
            print(obj)

content.lstrip("Contact Information")  # 刪掉

# Excel
# wb = load_workbook("test.xlsx")
# sheet = wb["General Information"]

# sheet['B2'] = content[content.find(
#     "Description") + len("Description"): content.find("Company Status")]

# sheet['C2'] = content[content.find(
#     "Company Status") + len("Company Status"): content.find("Website")]

# sheet['D2'] = content[content.find(
#     "Website") + len("Website"): content.find("Entity Types")]

# sheet['E2'] = content[content.find(
#     "Entity Types") + len("Entity Types"): content.find("Formerly Known As")]

# sheet['F2'] = content[content.find(
#     "Formerly Known As") + len("Formerly Known As"): content.find("Legal Name")]

# sheet['G2'] = content[content.find(
#     "Legal Name") + len("Legal Name"): content.find("Business Status")]

# sheet['F2'] = content[content.find(
#     "Business Status") + len("Business Status"): content.find("Ownership Status")]

# sheet['I2'] = content[content.find(
#     "Ownership Status") + len("Ownership Status"): content.find("Financing Status")]

# sheet['J2'] = content[content.find(
#     "Financing Status") + len("Financing Status"): content.find("Year Founded")]

# sheet['K2'] = content[content.find(
#     "Year Founded") + len("Year Founded"): content.find("Universe")]

# sheet['L2'] = content[content.find(
#     "Universe") + len("Universe"): content.find("SIC Code")]

# sheet['M2'] = content[content.find(
#     "SIC Code") + len("SIC Code"): content.find("Primary Industry")]

# sheet['N2'] = content[content.find(
#     "Primary Industry") + len("Primary Industry"): content.find("Other Industries")]

# sheet['O2'] = content[content.find(
#     "Other Industries") + len("Other Industries"): content.find("Verticals")]

# sheet['P2'] = content[content.find(
#     "Verticals") + len("Verticals"): content.find("Employees")]

# sheet['Q2'] = content[content.find(
#     "Employees") + len("Employees"): content.find("View Employee History")]

# sheet['R2'] = content[content.find(
#     "Primary Office") + len("Primary Office"): content.find("Phone", content.find(
#         "Primary Office"))]

# wb.save('test.xlsx')
