import PyPDF2
import docx2txt
import re

def extract_sentences_with_word_from_pdf(filename, word):
    sentences = []
    with open(filename, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        num_pages = len(reader.pages)

        for page_num in range(num_pages):
            page = reader.pages[page_num]
            text = page.extract_text()
            found_sentences = re.findall(r"([^.!?]*{}[^.!?]*[.!?])".format(word), text)
            
            for i, sentence in enumerate(found_sentences):
                if i > 0:
                    sentences.append(found_sentences[i-1]) 
                sentences.append(sentence) 
                if i < len(found_sentences) - 1:
                    sentences.append(found_sentences[i+1]) 

    return sentences

def extract_sentences_with_word_from_docx(filename, word):
    sentences = []
    text = docx2txt.process(filename)
    found_sentences = re.findall(r"([^.!?]*{}[^.!?]*[.!?])".format(word), text)
    
    for i, sentence in enumerate(found_sentences):
        if i > 0:
            sentences.append(found_sentences[i-1]) 
        sentences.append(sentence) 
        if i < len(found_sentences) - 1:
            sentences.append(found_sentences[i+1]) 

    return sentences

search_word = input("Enter the word to search for: ")
file_path = input("Enter the path of the file (PDF or DOCX): ")
sentences = []

if file_path.lower().endswith('.pdf'):
    sentences = extract_sentences_with_word_from_pdf(file_path, search_word)
elif file_path.lower().endswith('.docx'):
    sentences = extract_sentences_with_word_from_docx(file_path, search_word)
else:
    print("Unsupported file format.")

if sentences:
    print("Text relevant to the word '{}':".format(search_word))
    for sentence in sentences:
        print(sentence)
else:
    print("No text found relevant to the word '{}'. Please try another word".format(search_word))
