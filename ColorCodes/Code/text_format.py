'''
import easyocr
import os
import re

model_path = 'new_model.pth'
reader = easyocr.Reader(['en'], model_storage_directory=model_path)
folder_path = 'output'
files = os.listdir(folder_path)

color_number_pairs = []
colors_to_extract = ['GREEN', 'PINK','PURPLE', 'BLUE']
for file in files:
    if file.endswith(('.jpg', '.png', '.jpeg', '.tiff')):
        image_path = os.path.join(folder_path, file)
        result = reader.readtext(image_path, detail=0)

        for text in result:
            if re.match(r'^P(?!URPLE)', text):
                color = 'PINK'
            elif text.isalpha() and text in colors_to_extract:
                color = text
            elif re.match(r'^\d{1,3}(?:\s*-?\s*\d{1,3})*$', text):
                numbers = text.split('-')
                if color:
                    color_number_pairs.append((color, numbers))
                    color = None

formatted_output = ', '.join([f"{color}: {'-'.join(numbers)}" for color, numbers in color_number_pairs])
print(formatted_output)

'''
'''
import easyocr
import os
import re
import cv2

def count_triangles(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    edges = cv2.Canny(gray, 50, 150)
    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    triangle_count = 0
    for contour in contours:
        approx = cv2.approxPolyDP(contour, 0.01 * cv2.arcLength(contour, True), True)
        if len(approx) == 3:
            triangle_count += 1

    return triangle_count

model_path = 'new_model.pth'
reader = easyocr.Reader(['en'], model_storage_directory=model_path)
folder_path = 'output'
files = os.listdir(folder_path)

colors_to_extract = ['GREEN', 'PINK', 'PURPLE', 'BLUE']
color_number_chunks = []
current_chunk = []

for file in files:
    if file.endswith(('.jpg', '.png', '.jpeg', '.tiff')):
        image_path = os.path.join(folder_path, file)
        image = cv2.imread(image_path)
        result = reader.readtext(image_path, detail=0)

        for text in result:
            if re.match(r'^P(?!URPLE)', text):
                color = 'PINK'
            elif text.isalpha() and text in colors_to_extract:
                color = text
            elif re.match(r'^\d{1,3}(?:\s*-?\s*\d{1,3})*$', text):
                numbers = text.split('-')
                if color:
                    current_chunk.append((color, numbers))
                    color = None
        triangle_count = count_triangles(image)
        if current_chunk:
            color_number_chunks.append((current_chunk, triangle_count))
            current_chunk = []

for chunk, triangle_count in color_number_chunks:
    print("Chunk:")
    for color, numbers in chunk:
        print(f"Color: {color}, Number: {'-'.join(numbers)}")
    print(f"Number of triangles: {triangle_count}")
    print()

'''
'''
import easyocr
import os
import re

model_path = 'new_model.pth'
reader = easyocr.Reader(['en'], model_storage_directory=model_path)
folder_path = 'output'
files = os.listdir(folder_path)

color_number_pairs = []
colors_to_extract = ['GREEN', 'PINK','PURPLE', 'BLUE']
color = None  # Initialize color variable outside the loop

for file in files:
    if file.endswith(('.jpg', '.png', '.jpeg', '.tiff')):
        image_path = os.path.join(folder_path, file)
        result = reader.readtext(image_path, detail=0)

        for text in result:
            if re.match(r'^P(?!URPLE)', text):
                color = 'PINK'
            elif text.isalpha() and text in colors_to_extract:
                if color:
                    color_number_pairs.append((color, numbers))
                color = text
                numbers = []  # Initialize numbers list
            elif re.match(r'^\d{1,3}(?:\s*-?\s*\d{1,3})*$', text):
                numbers += text.split('-')

        # Append the last color and numbers if any
        if color:
            color_number_pairs.append((color, numbers))

formatted_output = ', '.join([f"{color}: {'-'.join(numbers)}" for color, numbers in color_number_pairs])
print(formatted_output)
'''
'''
import easyocr
import os
import re

model_path = 'new_model.pth'
reader = easyocr.Reader(['en'], model_storage_directory=model_path)
folder_path = 'output'
files = os.listdir(folder_path)

color_number_pairs = []
colors_to_extract = ['GREEN', 'PINK','PURPLE', 'BLUE']
color = None  # Initialize color variable outside the loop

for file in files:
    if file.endswith(('.jpg', '.png', '.jpeg', '.tiff')):
        image_path = os.path.join(folder_path, file)
        result = reader.readtext(image_path, detail=0)

        for text in result:
            if re.match(r'^P(?!URPLE)', text):
                color = 'PINK'
            elif text.isalpha() and text in colors_to_extract:
                if color:
                    color_number_pairs.append((color, numbers))
                color = text
                numbers = []  # Initialize numbers list
            elif re.match(r'^\d{1,3}(?:\s*-?\s*\d{1,3})*$', text):
                if numbers:
                    color_number_pairs.append((color, numbers))
                numbers = text.split('-')

        # Append the last color and numbers if any
        if color and numbers:
            color_number_pairs.append((color, numbers))

formatted_output = ', '.join([f"{color}: {'-'.join(numbers)}" for color, numbers in color_number_pairs])
print(formatted_output)
'''
import easyocr
import os
import re

model_path = 'new_model.pth'
reader = easyocr.Reader(['en'], model_storage_directory=model_path)
folder_path = 'output'
files = os.listdir(folder_path)

color_number_pairs = []
colors_to_extract = ['GREEN', 'PINK', 'PURPLE', 'BLUE']
color = None  # Initialize color variable outside the loop

for file in files:
    if file.endswith(('.jpg', '.png', '.jpeg', '.tiff')):
        image_path = os.path.join(folder_path, file)
        result = reader.readtext(image_path, detail=0)

        for text in result:
            if re.match(r'^P(?!URPLE)', text):
                color = 'PINK'
            elif text.isalpha() and text in colors_to_extract:
                if color and numbers:
                    color_number_pairs.append((color, tuple(numbers)))
                color = text
                numbers = []  # Initialize numbers list
            elif re.match(r'^\d{1,3}(?:\s*-\s*\d{1,3})*$', text):
                if numbers:
                    color_number_pairs.append((color, tuple(numbers)))
                numbers = text.split('-')

        # Append the last color and numbers if any
        if color and numbers:
            color_number_pairs.append((color, tuple(numbers)))

# Remove redundant color-number pairs
unique_pairs = set(color_number_pairs)

formatted_output = ', '.join([f"{color}: {'-'.join(numbers)}" for color, numbers in unique_pairs])
print(formatted_output)

