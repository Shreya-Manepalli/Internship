import fitz
import pandas as pd
doc = fitz.open('Sample2.pdf')
page1 = doc[0]
words = page1.get_text("words")
print(words[0])

def make_text(words):

    line_dict = {} 
    words.sort(key=lambda w: w[0])

    for w in words:  

        y1 = round(w[3], 1)  
        word = w[4] 
        line = line_dict.get(y1, [])  
        line.append(word)  
        line_dict[y1] = line  
    lines = list(line_dict.items())
    lines.sort()  

    return "n".join([" ".join(line[1]) for line in lines])

first_annots= []
rec=page1.first_annot.rect
mywords = [w for w in words if fitz.Rect(w[:4]) in rec]
ann= make_text(mywords)
first_annots.append(ann)
print(first_annots)
