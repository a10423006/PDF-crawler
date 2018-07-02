from pdfminer.layout import LAParams
from pdfminer.converter import HTMLConverter
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
import string
from string import punctuation
from io import BytesIO
from openpyxl import load_workbook
from bs4 import BeautifulSoup

# pdf path
document = open('/Users/anan/Desktop/pdf/Part1_test.pdf', 'rb')
# crate pdf manager
rsrcmgr = PDFResourceManager()

retster = BytesIO()
# Set parameters for analysis.
laparams = LAParams()
laparams.char_margin = 4  # 5 4best
laparams.line_margin = 5  # test:3 part:5
laparams.word_margin = 4
laparams.boxes_flow = 0.5

# Create a PDF page TMLConverter object
device = HTMLConverter(rsrcmgr, retster, codec='utf-8', laparams=laparams)
# create pdf interpreter
interpreter = PDFPageInterpreter(rsrcmgr, device)

for page in PDFPage.get_pages(document):
    interpreter.process_page(page)

content = retster.getvalue().decode()

# 以 Beautiful Soup 解析 HTML 程式碼
soup = BeautifulSoup(content, 'html.parser')

# 去 br標籤
invalid_tags = ['br']
for tag in invalid_tags:
    for match in soup.findAll(tag):
        match.replaceWithChildren()

f = open('/Users/anan/Desktop/pdf/test.html', 'a', encoding='utf8')
# f.write(soup.prettify())

# 找span標籤
span_tags = soup.find_all('span')

for tag in span_tags:
    str1 = tag.get('style')
    print(tag.string)
    # 找出所有字型大小為12的標籤
    # if str1.find('font-size:12px') != -1:
    #     print()

    #     # 找出所有字型大小為11的標籤
    # if str1.find('font-size:11px') != -1:
    #     print(tag.string)

document.close()
device.close()
retster.close()

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
