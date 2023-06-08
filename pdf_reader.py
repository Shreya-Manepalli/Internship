import PyPDF2

rd = PyPDF2.PdfReader('sample1.pdf')
pg = rd.pages[0]
print(pg.extract_text())


