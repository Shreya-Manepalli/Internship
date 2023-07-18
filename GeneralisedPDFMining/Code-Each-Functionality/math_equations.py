import PyPDF2
import docx2txt
import re
from docx import Document

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

def save_equations_to_docx(equations, output_path):
    doc = Document()
    
    for equation in equations:
        doc.add_paragraph(equation)
    
    doc.save(output_path)

file_path = input("Enter the path of the file (PDF or DOCX): ")
equations = []

if file_path.lower().endswith('.pdf'):
    equations = extract_equations_from_pdf(file_path)
elif file_path.lower().endswith('.docx'):
    equations = extract_equations_from_docx(file_path)
else:
    print("Unsupported file format.")
output_path = "equations.docx"
save_equations_to_docx(equations, output_path)
