import re
import docx
from PyPDF2 import PdfReader

def extract_figure_sentences_from_pdf(pdf_path):
    # Read the PDF file
    with open(pdf_path, 'rb') as file:
        pdf = PdfReader(file)
        num_pages = len(pdf.pages)

        figure_sentences = []
        for page_num in range(num_pages):
            page = pdf.pages[page_num]
            text = page.extract_text()

            # Extract sentences starting with "Fig" or "Figure"
            #sentences = re.findall(r'(?i)(fig.*?)(?=\n)', text)
            sentences = re.findall(r'(?i)((?:fig|figure).*?)(?=\n)', text)
            figure_sentences.extend(sentences)

    return figure_sentences

def save_sentences_to_word(sentences, output_path):
    doc = docx.Document()

    # Add each sentence as a paragraph in the Word document
    for sentence in sentences:
        doc.add_paragraph(sentence)

    # Save the Word document
    doc.save(output_path)

# Provide the path to your input PDF file
pdf_path = 'C:\Local\Project\StructuredTables\Pdf\SampleDocs\Michigan.pdf'

# Provide the path to save the output Word file
output_path = 'output.docx'

# Extract sentences starting with "Figure" from the PDF
figure_sentences = extract_figure_sentences_from_pdf(pdf_path)

# Save the sentences in a Word file
save_sentences_to_word(figure_sentences, output_path)
