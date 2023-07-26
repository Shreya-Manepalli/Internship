'''
import easyocr
import os
import re
import pandas as pd
import shutil

model_path = 'new_model.pth'
reader = easyocr.Reader(['en'], model_storage_directory=model_path)

input_folder_path = input("Enter the path of the input folder: ")
output_folder_path = input("Enter the path of the output folder: ")
colors_to_extract = ['GREEN', 'PINK', 'PURPLE', 'BLUE']

def process_subfolder(input_subfolder_path, output_subfolder_path):
    files = os.listdir(input_subfolder_path)

    color_number_dict = {}

    for file in files:
        if file.endswith(('.jpg', '.png', '.jpeg', '.tiff')):
            image_path = os.path.join(input_subfolder_path, file)
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
                        color_number_dict[color].extend(numbers)
                    color = text
                    numbers = []
                elif re.match(r'^\d{1,3}(?:\s*-\s*\d{1,3})*$', text):
                    numbers.append(text)
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
            
            for sub_root, sub_dirs, sub_files in os.walk(input_subfolder_path):
                for sub_dir_name in sub_dirs:
                    sub_input_subfolder_path = os.path.join(sub_root, sub_dir_name)
                    sub_output_subfolder_path = sub_input_subfolder_path.replace(input_folder, output_folder)
                    os.makedirs(sub_output_subfolder_path, exist_ok=True)
                    process_subfolder(sub_input_subfolder_path, sub_output_subfolder_path)


process_folder(input_folder_path, output_folder_path)

'''
'''
import easyocr
import os
import re
import pandas as pd
import shutil
import concurrent.futures

model_path = 'new_model.pth'
reader = easyocr.Reader(['en'], model_storage_directory=model_path)

input_folder_path = input("Enter the path of the input folder: ")
output_folder_path = input("Enter the path of the output folder: ")
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

'''


import easyocr
import os
import re
import pandas as pd
import shutil
import concurrent.futures

model_path = 'new_model.pth'
reader = easyocr.Reader(['en'], model_storage_directory=model_path)

input_folder_path = input("Enter the path of the input folder: ")
output_folder_path = input("Enter the path of the output folder: ")
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

    modified_color_number_dict = {}

    for color, numbers_list in color_number_dict.items():
        if color == 'PINK':
            modified_numbers_list = []
            index_to_skip = set()
            for i in range(len(numbers_list)):
                if i in index_to_skip:
                    continue
                number1 = numbers_list[i]
                for j in range(i + 1, len(numbers_list)):
                    number2 = numbers_list[j]
                    if number1[:3] == number2[:3] == color[:3]:
                        index_to_skip.add(j)
                        break
                else:
                    modified_numbers_list.append(number1)
            modified_color_number_dict[color] = modified_numbers_list
        else:
            modified_color_number_dict[color] = numbers_list

    max_numbers = max(len(numbers_list) for numbers_list in modified_color_number_dict.values())

    for color, numbers_list in modified_color_number_dict.items():
        if len(numbers_list) < max_numbers:
            numbers_list.extend([None] * (max_numbers - len(numbers_list)))

    df = pd.DataFrame(modified_color_number_dict)

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

