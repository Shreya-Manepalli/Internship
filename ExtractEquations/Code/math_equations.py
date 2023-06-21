import PyPDF2
import docx2txt
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

def extract_equations_from_docx(docx_path):
    equations = []

    text = docx2txt.process(docx_path)
    equation_pattern = r'[^\n=]+=[^\n]+'
    matches = re.findall(equation_pattern, text)
    equations.extend(matches)
    
    return equations

file_path = input("Enter the path of the file (PDF or DOCX): ")
equations = []
if file_path.lower().endswith('.pdf'):
    equations = extract_equations_from_pdf(file_path)
    print("The equations in the PDF document are:")
elif file_path.lower().endswith('.docx'):
    equations = extract_equations_from_docx(file_path)
    print("The equations in the Word document are:")
else:
    print("Unsupported file format.")

for equation in equations:
    print(equation)
