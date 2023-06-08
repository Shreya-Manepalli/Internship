
import pandas as pd
import numpy as np
import PyPDF2
import textract
import re

filename ='sample3.pdf' 

pdfFileObj = open(filename,'rb')               
pdfReader = PyPDF2.PdfReader(pdfFileObj)   
num_pages = len(pdfReader.pages)                


count = 0
text = ""
                                                            
while count < num_pages:                       
    pageObj = pdfReader.pages[count]
    count +=1
    text += pageObj.extract_text()


if text != "":
    text = text


else:
    text = textract.process('http://bit.ly/epo_keyword_extraction_document', method='tesseract', language='eng')
text = text.encode('ascii','ignore') 
print(text)