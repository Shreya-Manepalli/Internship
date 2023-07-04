'''
from PIL import Image
import cv2
import numpy as np
import os

Image.MAX_IMAGE_PIXELS = None

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


input_folder = 'C:\Local\Project\output_images'  
output_folder = 'C:\Local\Project\output1'  


if not os.path.exists(output_folder):
    os.makedirs(output_folder)


image_files = [f for f in os.listdir(input_folder) if f.endswith('.jpg')]

for image_file in image_files:
    image_path = os.path.join(input_folder, image_file)
    image_name = os.path.splitext(image_file)[0]
    output_subfolder = os.path.join(output_folder, f"I_{image_name}")
    #output_subfolder = os.path.join(output_folder, os.path.splitext(image_file)[0])

    
    if not os.path.exists(output_subfolder):
        os.makedirs(output_subfolder)

    cropped_chunks = crop_image_into_chunks(image_path)
    for i, chunk in enumerate(cropped_chunks):
        chunk.save(os.path.join(output_subfolder, f'chunk_{i+1}.jpg'))
'''


### SAME STRUCTURE AS INPUT FOLDER ALONG WITH SUBFOLDERS FOR EACH FILE CONSISTING THE CHUNKS OF THE IMAGE


from PIL import Image
import cv2
import numpy as np
import os
import shutil

Image.MAX_IMAGE_PIXELS = None

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


input_folder = 'C:\Local\Project\File_container_output'

for root, dirs, files in os.walk(input_folder):
    for file in files:
        if file.endswith('.jpg'):
            input_path = os.path.join(root, file)

            cropped_chunks = crop_image_into_chunks(input_path)
            for i, chunk in enumerate(cropped_chunks):
                chunk.save(os.path.join(root, f'chunk_{i+1}.jpg'))

            os.remove(input_path)





