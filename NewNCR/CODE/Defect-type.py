import re
import PyPDF2

def extract_text_between_phrases(pdf_file):
    text_between_phrases = []
    
    # Open the PDF file
    with open(pdf_file, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        
        # Iterate over each page in the PDF
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            page_text = page.extract_text()
            
            # Find matches between the phrases using regex
            matches = re.findall(r"Defect type(.*?)Charact.No", page_text, re.DOTALL)
            
            # Add the matches to the result list
            text_between_phrases.extend(matches)
    
    return text_between_phrases

# Usage example
pdf_file_path = 'NCR-dummy.pdf'
result = extract_text_between_phrases(pdf_file_path)

# Print the extracted text or numbers
for i, text in enumerate(result):
    print(text.strip(), end='')
    if i < len(result) - 1:
        print(',', end='')