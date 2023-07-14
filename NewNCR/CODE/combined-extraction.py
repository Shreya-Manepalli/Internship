
import re
import PyPDF2

def extract_text_between_phrases(pdf_file):
    text_between_phrases1 = []
    text_between_phrases2 = []
    text_between_phrases3 = []
    
    # Open the PDF file
    with open(pdf_file, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        
        # Iterate over each page in the PDF
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            page_text = page.extract_text()
            
            # Find matches between the phrases using regex
            matches1 = re.findall(r"Description of Nonconformance(.*?)Cause of Nonconformance", page_text, re.DOTALL)
            matches2 = re.findall(r"Charact.No(.*?)Requirement Location", page_text, re.DOTALL)
            matches3 = re.findall(r"Defect type(.*?)Charact.No", page_text, re.DOTALL)
            
            # Add the matches to the result list
            text_between_phrases1.extend(matches1)
            text_between_phrases2.extend(matches2)
            text_between_phrases3.extend(matches3)
    
    return text_between_phrases1, text_between_phrases2, text_between_phrases3

# Usage example
pdf_file_path = 'NCR-dummy.pdf'
result1, result2, result3 = extract_text_between_phrases(pdf_file_path)

# Print the extracted text or numbers
if result1:
    print('Description of NC:', end=' ')
    print(', '.join(result1), end='')
print()

if result2:
    print('Charact.No:', end=' ')
    print(', '.join(result2), end='')
print()

if result3:
    print('Defect Type:', end=' ')
    print(', '.join(result3), end='')