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
            sentences += re.findall(r"[^.!?]*{}[^.!?]*[.!?]".format(word), text)

    return sentences

# Prompt the user for the input word
search_word = input("Enter the word to search for: ")

# Provide the filename of the PDF file
pdf_filename = input("Enter the path or URL of the PDF file: ")

# Extract the sentences containing the word from the PDF
sentences = extract_sentences_with_word(pdf_filename, search_word)

# Print the extracted sentences
if sentences:
    print("Sentences containing '{}':".format(search_word))
    for sentence in sentences:
        print(sentence)
else:
    print("No sentences found containing '{}'.".format(search_word))
