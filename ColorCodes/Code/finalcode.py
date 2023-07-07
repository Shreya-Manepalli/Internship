#IMPORTING THE LIBRARIES
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


######################################################  CODE TO CONVERT PDF TO IMAGES  ###############################################################


Image.MAX_IMAGE_PIXELS = None

#This function takes the main PARENT FOLDER as INPUT FOLDER and then walks through all the subfolders in it and converts each pdf file to an image to store in the OUTPUT FOLDER
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
                #Coverts the pdf files into images and stores them in the respective locations
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
                os.makedirs(output_subfolder)  #Replicates the same input folder structure to store images in the place of the respective pdf files


input_folder = input("Enter the path of the parent input folder (which contains the pdf files): ") #Takes the main parent folder as the input
output_folder = input("Enter the path of the intermediate output folder (chunks and images will be stored): ")  #Intermediate output folder which stores all the Images of the pdf files

#Function to convert pdf files to images is being called here
convert_pdf_to_images(input_folder, output_folder)


######################################################  CODE TO CHUNK THE IMAGES  ###############################################################


#This function checks if the chunk has any data and deletes all the irrelevant chunks which don't have any significant information
def has_data(chunk):
    chunk_gray = chunk.convert('L')
    edges = cv2.Canny(np.array(chunk_gray), 50, 150)

    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE) #Checks if the chunks have any contours or lines

    min_area = chunk.width * chunk.height * 0.001
    max_area = chunk.width * chunk.height * 0.9
    valid_contours = []
    for contour in contours:
        area = cv2.contourArea(contour) #Checks for the contour area
        x, y, w, h = cv2.boundingRect(contour)
        aspect_ratio = w / h if h != 0 else 0
        if area > min_area and area < max_area and aspect_ratio > 0.1 and aspect_ratio < 10:
            valid_contours.append(contour)

    if valid_contours:
        return True

    return False

#Chunk Width and Height respectively

m=15
n=15

#This function crops the images into chunks for readability purposes
def crop_image_into_chunks(image_path):
    image = Image.open(image_path)
    width, height = image.size
    chunk_x = width // m
    chunk_y = height // n
    padding_left = 400     #Left padding added to each chunk
    padding_right = 400    #Right padding added to each chunk
    padding_top = 300      #Top padding added to each chunk
    padding_bottom = 300   #Bottom padding added to each chunk
    chunks = []
    for i in range(m):
        for j in range(n):
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

#This takes the folder containing the images as the input
input_folder1 = output_folder

for root, dirs, files in os.walk(input_folder1):
    for file in files:
        if file.endswith('.jpg'):
            input_path = os.path.join(root, file)

            cropped_chunks = crop_image_into_chunks(input_path)
            for i, chunk in enumerate(cropped_chunks):
                chunk.save(os.path.join(root, f'chunk_{i+1}.jpg'))

            os.remove(input_path)  #Removes the images after chunking is done


######################################################  CODE TO GENERATE EXCEL FILES  ###############################################################


model_path = 'new_model.pth'
reader = easyocr.Reader(['en'], model_storage_directory=model_path)
input_folder_path = input_folder1  #Takes the folder containing the chunks of the images as input
output_folder_path = input("Enter the path of the final output folder where the excel files will be stored: ") #Final output folder where the excel files corresponding to each pdf file will stored

#Dictionary containing the colors (CAN ADD MORE COLORS TO THIS)
colors_to_extract = ['GREEN', 'PINK', 'PURPLE', 'BLUE']

#Function to read each image and extract the colors and the numbers
def process_image(image_path):
    result = reader.readtext(image_path, detail=0)

    color = None
    numbers = []
    color_number_dict = {}
    for text in result:
        if re.match(r'^P(?!URPLE)', text):
            color = 'PINK'      #Hardcoded PINK Color due to PDF Format issues (Any text starting with P and is not PURPLE will be considered as PINK)
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

#Threading in order to proccess multiple chunks at once
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

#Checks if any colors and numbers are present in the chunk
    if not color_number_dict:
        return

    max_numbers = max(len(numbers_list) for numbers_list in color_number_dict.values())

    for color, numbers_list in color_number_dict.items():
        if len(numbers_list) < max_numbers:
            numbers_list.extend([None] * (max_numbers - len(numbers_list)))

    df = pd.DataFrame(color_number_dict)  #Defines the dataframe with colors as column headers and the numbers as the values 

    pdf_file_name = os.path.basename(os.path.dirname(input_subfolder_path))

    output_file = os.path.join(output_subfolder_path, f"{pdf_file_name}.xlsx") #Output excel file
    with pd.ExcelWriter(output_file) as writer:
        df.to_excel(writer, sheet_name= 'data', index=False)

    print(f"Output saved to {output_file} successfully.")

#Replicates the input directory to store excel files corresponding to each pdf file
def process_folder(input_folder, output_folder):
    for root, dirs, files in os.walk(input_folder):
        for dir_name in dirs:
            input_subfolder_path = os.path.join(root, dir_name)
            output_subfolder_path = input_subfolder_path.replace(input_folder, output_folder)
            os.makedirs(output_subfolder_path, exist_ok=True)
            process_subfolder(input_subfolder_path, output_subfolder_path)

#Fuction being called 
process_folder(input_folder_path, output_folder_path)

#Removes the Intermediate folder containing all the chunks and the images
shutil.rmtree(input_folder_path)