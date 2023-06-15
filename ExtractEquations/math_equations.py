import PyPDF2
import re

def extract_equations_from_pdf(pdf_path):
    equations = []
    
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        num_pages = len(pdf_reader.pages)
        
        for page_number in range(num_pages):
            page = pdf_reader.pages[page_number]
            text = page.extract_text()
            equation_pattern = r'[^\n=]+=[^\n]+'
            matches = re.findall(equation_pattern, text)
            equations.extend(matches)
    
    return equations
pdf_file_path = input("Enter the path or URL of the PDF file: ")
equations = extract_equations_from_pdf(pdf_file_path)
print("The equations in the following document are:")
for equation in equations:
    print(equation)


























