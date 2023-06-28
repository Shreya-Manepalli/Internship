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

