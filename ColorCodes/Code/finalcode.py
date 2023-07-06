import fitz
from PIL import Image
import os
import shutil
from PIL import Image
import cv2
import numpy as np
import os
import shutil
import easyocr
import os
import re
import pandas as pd
import easyocr
import os
import re
import pandas as pd
import shutil
import concurrent.futures


Image.MAX_IMAGE_PIXELS = None

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
input_folder = input("Enter the path of the parent input folder (which contains the pdf files): ")
output_folder = input("Enter the path of the output folder (chunks and images will be stored): ")

convert_pdf_to_images(input_folder, output_folder)

def has_data(chunk):
    chunk_gray = chunk.convert('L')
    edges = cv2.Canny(np.array(chunk_gray), 50, 150)

    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    min_area = chunk.width * chunk.height * 0.001
    max_area = chunk.width * chunk.height * 0.9
    valid_contours = []
    for contour in contours:
        area = cv2.contourArea(contour)
        x, y, w, h = cv2.boundingRect(contour)
        aspect_ratio = w / h if h != 0 else 0
        if area > min_area and area < max_area and aspect_ratio > 0.1 and aspect_ratio < 10:
            valid_contours.append(contour)

    if valid_contours:
        return True

    return False

def crop_image_into_chunks(image_path):
    image = Image.open(image_path)
    width, height = image.size
    chunk_x = width // 15
    chunk_y = height // 15
    padding_left = 400
    padding_right = 400
    padding_top = 300
    padding_bottom = 300
    chunks = []
    for i in range(15):
        for j in range(15):
            left = j * chunk_x
            upper = i * chunk_y
            right = left + chunk_x
            lower = upper + chunk_y
            left_padding = padding_left if j > 0 else 0
            right_padding = padding_right if j < 2 else 0
            top_padding = padding_top if i > 0 else 0
            bottom_padding = padding_bottom if i < 2 else 0
            left -= left_padding
            right += right_padding
            upper -= top_padding
            lower += bottom_padding
            chunk = image.crop((left, upper, right, lower))
            if has_data(chunk):
                chunks.append(chunk)
    return chunks


input_folder1 = output_folder

for root, dirs, files in os.walk(input_folder1):
    for file in files:
        if file.endswith('.jpg'):
            input_path = os.path.join(root, file)

            cropped_chunks = crop_image_into_chunks(input_path)
            for i, chunk in enumerate(cropped_chunks):
                chunk.save(os.path.join(root, f'chunk_{i+1}.jpg'))

            os.remove(input_path)

model_path = 'new_model.pth'
reader = easyocr.Reader(['en'], model_storage_directory=model_path)

#input_folder_path = input("Enter the path of the input folder: ")
input_folder_path = input_folder1
output_folder_path = input("Enter the path of the final output folder where the excel files will be stored: ")
colors_to_extract = ['GREEN', 'PINK', 'PURPLE', 'BLUE']

def process_image(image_path):
    result = reader.readtext(image_path, detail=0)

    color = None
    numbers = []
    color_number_dict = {}
    for text in result:
        if re.match(r'^P(?!URPLE)', text):
            color = 'PINK'
        elif text.isalpha() and text in colors_to_extract:
            if color and numbers:
                if color not in color_number_dict:
                    color_number_dict[color] = []
                color_number_dict[color].extend(numbers)
            color = text
            numbers = []
        elif re.match(r'^\d{1,3}(?:\s*-\s*\d{1,3})*$', text):
            numbers.append(text)

    return color, numbers


def process_subfolder(input_subfolder_path, output_subfolder_path):
    files = os.listdir(input_subfolder_path)

    color_number_dict = {}

    with concurrent.futures.ThreadPoolExecutor() as executor:
        image_paths = [os.path.join(input_subfolder_path, file) for file in files if file.endswith(('.jpg', '.png', '.jpeg', '.tiff'))]
        results = executor.map(process_image, image_paths)

        for color, numbers in results:
            if color and numbers:
                if color not in color_number_dict:
                    color_number_dict[color] = []
                color_number_dict[color].extend(numbers)

    if not color_number_dict:
        print(f"No relevant information found in {input_subfolder_path}.")
        return

    max_numbers = max(len(numbers_list) for numbers_list in color_number_dict.values())

    for color, numbers_list in color_number_dict.items():
        if len(numbers_list) < max_numbers:
            numbers_list.extend([None] * (max_numbers - len(numbers_list)))

    df = pd.DataFrame(color_number_dict)

    output_file = os.path.join(output_subfolder_path, 'output.xlsx')
    with pd.ExcelWriter(output_file) as writer:
        df.to_excel(writer, sheet_name='Data', index=False)

    print(f"Output saved to {output_file} successfully.")


def process_folder(input_folder, output_folder):
    for root, dirs, files in os.walk(input_folder):
        for dir_name in dirs:
            input_subfolder_path = os.path.join(root, dir_name)
            output_subfolder_path = input_subfolder_path.replace(input_folder, output_folder)
            os.makedirs(output_subfolder_path, exist_ok=True)
            process_subfolder(input_subfolder_path, output_subfolder_path)


process_folder(input_folder_path, output_folder_path)

shutil.rmtree(input_folder_path)