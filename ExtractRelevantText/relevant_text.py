import PyPDF2
import re

def extract_sentences_with_word(filename, word):
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

search_word = input("Enter the word to search for: ")
pdf_filename = input("Enter the path or URL of the PDF file: ")
sentences = extract_sentences_with_word(pdf_filename, search_word)
if sentences:
    print("Text relevant to the word '{}':".format(search_word))
    for sentence in sentences:
        print(sentence)
else:
    print("No text found relevant to '{}'. Please try another word".format(search_word))
