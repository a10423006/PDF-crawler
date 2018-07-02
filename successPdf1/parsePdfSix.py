from pdfminer.layout import LAParams
from pdfminer.converter import PDFPageAggregator
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LTTextBox
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
laparams.char_margin = 3.2   #5 4best 
laparams.line_margin = 8   #5-8
laparams.word_margin = 10
laparams.boxes_flow = 0.5
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
    for obj in layout:
        if isinstance(obj, LTTextBox):
            # obj = element.get_text
            # obj = element._objs[0]
            print("text: ", obj.get_text()) #將換行取代成空增加讀取速度
            # print ("text: ", obj)
            print ("x0: ", obj.x0)
            print ("yo: ", obj.y0)
            print ("x1: ", obj.x1)
            print ("y1: ", obj.y1)
            print ("checkPoint:",(checkPoint))
            if partOne in str(obj):         
                checkPoint = 1

            elif partTwo in str(obj):             
                checkPoint = 2 

            elif partThree in str(obj):             
                checkPoint = 3 
            # results = results.encode("utf8")
            # results = results.decode("utf8")
            # print(results)
            # results.encode(sys.stdin.encoding, "replace").decode(sys.stdin.encoding)
            
            if checkPoint == 1:
                print("text: ", obj.get_text()) #將換行取代成空增加讀取速度
                with open('C:/Users/Vincent/Dropbox/writeProgram/python/1.txt', 'a',encoding = 'utf8') as f: #encoding = 'utf8' 超級重要!!!
                    results = obj.get_text()
                    f.write('\n')
                    f.write(str(results)) #f objeocct open txt
                    f.write('\n')
                    f.write("x0:  ") #f objeocct open txt
                    f.write(str(obj.x0)) #f elementeocct open txt  
                    f.write('\t') 
                    f.write("y0:  ") #f elementeocct open txt 
                    f.write(str(obj.y0)) #f elementeocct open txt  
                    f.write('\t') 
                    f.write("x1:  ") #f elementeocct open txt  
                    f.write(str(obj.x1)) #f elementeocct open txt  
                    f.write('\t')   
                    f.write("y1:  ") #f elementeocct open txt
                    f.write(str(obj.y1)) #f objeocct open txt     
                    f.write('\n')
                         
            elif checkPoint == 2:
                print("text: ", obj.get_text()) #將換行取代成空增加讀取速度
                with open('C:/Users/Vincent/Dropbox/writeProgram/python/2.txt', 'a',encoding = 'utf8') as d: #encoding = 'utf8' 超級重要!!!
                    results = obj.get_text()
                    d.write('\n')
                    d.write(str(results)) #d objeocct open txt
                    d.write('\n')
                    d.write("x0:  ") #d objeocct open txt
                    d.write(str(obj.x0)) #d elementeocct open txt  
                    d.write('\t') 
                    d.write("y0:  ") #d elementeocct open txt 
                    d.write(str(obj.y0)) #d elementeocct open txt  
                    d.write('\t') 
                    d.write("x1:  ") #d elementeocct open txt  
                    d.write(str(obj.x1)) #d elementeocct open txt  
                    d.write('\t')   
                    d.write("y1:  ") #d elementeocct open txt
                    d.write(str(obj.y1)) #d elementeocct open txt     
                    d.write('\n')      
            elif checkPoint == 3:
                print("text: ", obj.get_text()) #將換行取代成空增加讀取速度
                with open('C:/Users/Vincent/Dropbox/writeProgram/python/3.txt', 'a',encoding = 'utf8') as i: #encoding = 'utf8' 超級重要!!!
                    results = obj.get_text()
                    i.write('\n')
                    i.write(str(results)) #i objeocct open txt
                    i.write('\n')
                    i.write("x0:  ") #i objeocct open txt
                    i.write(str(obj.x0)) #i elementeocct open txt  
                    i.write('\t') 
                    i.write("y0:  ") #i elementeocct open txt 
                    i.write(str(obj.y0)) #i elementeocct open txt  
                    i.write('\t') 
                    i.write("x1:  ") #i elementeocct open txt  
                    i.write(str(obj.x1)) #i elementeocct open txt  
                    i.write('\t')   
                    i.write("y1:  ") #i elementeocct open txt
                    i.write(str(obj.y1)) #i objeocct open txt     
                    i.write('\n')  
                        # print(re.sub('\W', '', results)) #將符號替代成空白的寫法
                        # results = results.decode("utf8")
                        # string = re.sub("[†\s+\.\!\/_,$%^*(+\"\']+|[+——！，。？、~@#￥%……&*（）]+".decode("utf8"), "".decode("utf8"),results) 
                        # results.encode("utf8").decode("cp950", "ignore")
                        # .decode("cp950", "ignore")
                        # results.encode("utf8").decode("\u2020", "ignore")
                        # results.replace("†","  ")
            
            # print("--------------------")