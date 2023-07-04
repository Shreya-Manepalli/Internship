'''
import fitz
from PIL import Image
import os

input_folder = 'C:\Local\Project\PDFFiles'  
output_folder = 'C:\Local\Project\output_images'  
dpi = 900


if not os.path.exists(output_folder):
    os.makedirs(output_folder)


pdf_files = [f for f in os.listdir(input_folder) if f.endswith('.pdf')]

for pdf_file in pdf_files:
    pdf_path = os.path.join(input_folder, pdf_file)
    doc = fitz.open(pdf_path)

    
    page = doc[0]
    zoom = dpi / 72.0
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    output_path = os.path.join(output_folder, f'{os.path.splitext(pdf_file)[0]}.jpg')
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    img.save(output_path, dpi=(dpi, dpi))

    doc.close()
'''


### SAME STRUCTURE AS INPUT FOLDER ALONG WITH SUBFOLDERS FOR EACH FILE CONSISTING THE IMAGE


import fitz
from PIL import Image
import os
import shutil

def convert_pdf_to_images(input_folder, output_folder, dpi=900):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for root, dirs, files in os.walk(input_folder):
        for filename in files:
            if filename.endswith('.pdf'):
                pdf_path = os.path.join(root, filename)
                relative_path = os.path.relpath(pdf_path, input_folder)
                output_subfolder = os.path.dirname(relative_path)
                output_subfolder_path = os.path.join(output_folder, output_subfolder)

                if not os.path.exists(output_subfolder_path):
                    os.makedirs(output_subfolder_path)

                doc = fitz.open(pdf_path)

                for i, page in enumerate(doc):
                    zoom = dpi / 72.0
                    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                    page_output_subfolder = os.path.splitext(filename)[0]
                    page_output_subfolder_path = os.path.join(output_subfolder_path, page_output_subfolder)

                    if not os.path.exists(page_output_subfolder_path):
                        os.makedirs(page_output_subfolder_path)

                    output_path = os.path.join(page_output_subfolder_path, f'page_{i+1}.jpg')
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    img.save(output_path, dpi=(dpi, dpi))

                doc.close()

        for dir_name in dirs:
            input_subfolder = os.path.join(root, dir_name)
            output_subfolder = os.path.join(output_folder, os.path.relpath(input_subfolder, input_folder))
            if not os.path.exists(output_subfolder):
                os.makedirs(output_subfolder)

# Example usage
input_folder = 'C:\Local\Project\File_container'
output_folder = 'C:\Local\Project\File_container_output'

convert_pdf_to_images(input_folder, output_folder)
