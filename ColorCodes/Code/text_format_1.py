##### WITHOUT THREADING
'''
import easyocr
import os
import re
import pandas as pd

model_path = 'new_model.pth'
reader = easyocr.Reader(['en'], model_storage_directory=model_path)

folder_path = 'C:\Local\Project\output1\I286A5111-2359-WRR0066 (V5)'
files = os.listdir(folder_path)

color_number_dict = {}
colors_to_extract = ['GREEN', 'PINK', 'PURPLE', 'BLUE']

for file in files:
    if file.endswith(('.jpg', '.png', '.jpeg', '.tiff')):
        image_path = os.path.join(folder_path, file)
        result = reader.readtext(image_path, detail=0)

        color = None 
        numbers = [] 

        for text in result:
            if re.match(r'^P(?!URPLE)', text):
                color = 'PINK'
            elif text.isalpha() and text in colors_to_extract:
                if color and numbers:
                    if color not in color_number_dict:
                        color_number_dict[color] = []
                    color_number_dict[color].append(', '.join(numbers)) 
                color = text
                numbers = [] 
            elif re.match(r'^\d{1,3}(?:\s*-\s*\d{1,3})*$', text):
                numbers.append(text)
        if color and numbers:
            if color not in color_number_dict:
                color_number_dict[color] = []
            color_number_dict[color].append(', '.join(numbers)) 

max_numbers = max(len(numbers_list) for numbers_list in color_number_dict.values())

for color, numbers_list in color_number_dict.items():
    if len(numbers_list) < max_numbers:
        numbers_list.extend([None] * (max_numbers - len(numbers_list)))

df = pd.DataFrame(color_number_dict)

output_file = 'output.xlsx'
df.to_excel(output_file, index=False)

print(f"Output saved to {output_file} successfully.")
'''
##### WITH THREADING PER IMAGE FOR ONE FILE

'''

import easyocr
import os
import re
import pandas as pd
import threading

model_path = 'new_model.pth'
reader = easyocr.Reader(['en'], model_storage_directory=model_path)

folder_path = 'C:\Local\Project\output1\I286A3309-2369-WRR0061 (V5)'
files = os.listdir(folder_path)

color_number_dict = {}
colors_to_extract = ['GREEN', 'PINK', 'PURPLE', 'BLUE']

lock = threading.Lock()  
def process_image(file):
    if file.endswith(('.jpg', '.png', '.jpeg', '.tiff')):
        image_path = os.path.join(folder_path, file)
        result = reader.readtext(image_path, detail=0)

        color = None 
        numbers = [] 

        for text in result:
            if re.match(r'^P(?!URPLE)', text):
                color = 'PINK'
            elif text.isalpha() and text in colors_to_extract:
                if color and numbers:
                    with lock:
                        if color not in color_number_dict:
                            color_number_dict[color] = []
                        color_number_dict[color].append(', '.join(numbers))
                color = text
                numbers = [] 
            elif re.match(r'^\d{1,3}(?:\s*-\s*\d{1,3})*$', text):
                numbers.append(text)
        if color and numbers:
            with lock:
                if color not in color_number_dict:
                    color_number_dict[color] = []
                color_number_dict[color].append(', '.join(numbers))

threads = []

for file in files:
    thread = threading.Thread(target=process_image, args=(file,))
    thread.start()
    threads.append(thread)

for thread in threads:
    thread.join()

max_numbers = max(len(numbers_list) for numbers_list in color_number_dict.values())

for color, numbers_list in color_number_dict.items():
    if len(numbers_list) < max_numbers:
        numbers_list.extend([None] * (max_numbers - len(numbers_list)))

df = pd.DataFrame(color_number_dict)

output_file = 'output.xlsx'
df.to_excel(output_file, index=False)

print(f"Output saved to {output_file} successfully.")

'''

### WITH THREADING FOR FOLDER INPUT

import easyocr
import os
import re
import pandas as pd
import threading

model_path = 'new_model.pth'
reader = easyocr.Reader(['en'], model_storage_directory=model_path)

parent_folder_path = 'C:\Local\Project\File_container_output\I_286A5111-2359-WRR0066'

subfolders = [subfolder for subfolder in os.listdir(parent_folder_path) if os.path.isdir(os.path.join(parent_folder_path, subfolder))]

colors_to_extract = ['GREEN', 'PINK', 'PURPLE', 'BLUE']

lock = threading.Lock() 
def process_subfolder(subfolder):
    subfolder_path = os.path.join(parent_folder_path, subfolder)
    files = os.listdir(subfolder_path)

    color_number_dict = {}

    for file in files:
        if file.endswith(('.jpg', '.png', '.jpeg', '.tiff')):
            image_path = os.path.join(subfolder_path, file)
            result = reader.readtext(image_path, detail=0)

            color = None
            numbers = []

            for text in result:
                if re.match(r'^P(?!URPLE)', text):
                    color = 'PINK'
                elif text.isalpha() and text in colors_to_extract:
                    if color and numbers:
                        with lock:
                            if color not in color_number_dict:
                                color_number_dict[color] = []
                            color_number_dict[color].append(', '.join(numbers))
                    color = text
                    numbers = []
                elif re.match(r'^\d{1,3}(?:\s*-\s*\d{1,3})*$', text):
                    numbers.append(text)
            if color and numbers:
                with lock:
                    if color not in color_number_dict:
                        color_number_dict[color] = []
                    color_number_dict[color].append(', '.join(numbers))

    max_numbers = max(len(numbers_list) for numbers_list in color_number_dict.values())

    for color, numbers_list in color_number_dict.items():
        if len(numbers_list) < max_numbers:
            numbers_list.extend([None] * (max_numbers - len(numbers_list)))

    df = pd.DataFrame(color_number_dict)

    output_file = os.path.join(subfolder_path, 'output.xlsx')
    df.to_excel(output_file, index=False)

    print(f"Output saved to {output_file} successfully.")

threads = []

for subfolder in subfolders:
    thread = threading.Thread(target=process_subfolder, args=(subfolder,))
    thread.start()
    threads.append(thread)

for thread in threads:
    thread.join()





