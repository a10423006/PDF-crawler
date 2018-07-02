from pdfminer.layout import LAParams
from pdfminer.converter import PDFPageAggregator
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LTTextBoxVertical
from string import punctuation 
import sys
import io
from string import punctuation 
import re


# pdf path
document = open('C:/Users/Vincent/Dropbox/writeProgram/python/20180509.pdf', 'rb')
#crate pdf manager
rsrcmgr = PDFResourceManager()
# Set parameters for analysis.
laparams = LAParams()
# Create a PDF page aggregator object
device = PDFPageAggregator(rsrcmgr, laparams=laparams)
#create pdf interpreter
interpreter = PDFPageInterpreter(rsrcmgr , device)

checkPoint = 0  #初始狀態 檢查各個段落
partOne = "General Information"
partTwo = "Deal History"
partThree = "Investors ("
# 这里layout是一个LTPage对象 里面存放着 这个page解析出的各种对象 一般包括LTTextBox, LTFigure, LTImage, 
# LTTextBoxHorizontal 等等 想要获取文本就获得对象的text属性，  
for page in PDFPage.get_pages(document):
    interpreter.process_page(page)
    layout = device.get_result()
    for element in layout:
        if isinstance(element, LTTextBoxVertical):
            obj = element._objs[0]
            # print("text: ", obj.get_text().replace('\n','')) #將換行取代成空增加讀取速度
            testResult = (obj.get_text().replace('\n',''))
            print ("text: ", testResult)
            print ("checkPoint:",(checkPoint))
            if partOne in testResult:         
                checkPoint = 1
            elif partTwo in testResult:             
                checkPoint = 2 
            elif partThree in testResult:             
                checkPoint = 3 
            # results = results.encode("utf8")
            # results = results.decode("utf8")
            # print(results)
            # results.encode(sys.stdin.encoding, "replace").decode(sys.stdin.encoding)
            
            if checkPoint == 1:
                print("text: ", obj.get_text().replace('\n','')) #將換行取代成空增加讀取速度
                with open('C:/Users/Vincent/Dropbox/writeProgram/python/1.txt', 'a',encoding = 'utf8') as f: #encoding = 'utf8' 超級重要!!!
                    results = obj.get_text()
                    f.write(results) #f objeocct open txt         
            elif checkPoint == 2:
                print("text: ", obj.get_text().replace('\n','')) #將換行取代成空增加讀取速度
                with open('C:/Users/Vincent/Dropbox/writeProgram/python/2.txt', 'a',encoding = 'utf8') as d: #encoding = 'utf8' 超級重要!!!
                    results = obj.get_text()
                    d.write(results) #f objeocct open txt      
            elif checkPoint == 3:
                print("text: ", obj.get_text().replace('\n','')) #將換行取代成空增加讀取速度
                with open('C:/Users/Vincent/Dropbox/writeProgram/python/3.txt', 'a',encoding = 'utf8') as i: #encoding = 'utf8' 超級重要!!!
                    results = obj.get_text()
                    i.write(results) #f objeocct open txt    
                        # print(re.sub('\W', '', results)) #將符號替代成空白的寫法
                        # results = results.decode("utf8")
                        # string = re.sub("[†\s+\.\!\/_,$%^*(+\"\']+|[+——！，。？、~@#￥%……&*（）]+".decode("utf8"), "".decode("utf8"),results) 
                        # results.encode("utf8").decode("cp950", "ignore")
                        # .decode("cp950", "ignore")
                        # results.encode("utf8").decode("\u2020", "ignore")
                        # results.replace("†","  ")
                        
            # print("--------------------")