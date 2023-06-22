'''
import cv2
import numpy as np
import easyocr

im_1_path = 'C:\Local\Project\convert_pdf_to_png_page_1_out.png'
def recognize_text(img_path):
    
    reader = easyocr.Reader(['en'])
    return reader.readtext(img_path, detail = 0)
result = recognize_text(im_1_path)
print(result)
'''
'''
import cv2
import easyocr


image_path = "C:\Local\Project\convert_pdf_to_png_page_1_out.png"
image = cv2.imread(image_path)

reader = easyocr.Reader(['en'])  
result = reader.readtext(image)

desired_colors = ["GREEN", "PINK", "PURPLE"]

data = {}
for detection in result:
    text = detection[1].strip()
    if text in desired_colors:
        number_detection = result[result.index(detection) + 1]
        if number_detection[1].isdigit():
            data[text] = int(number_detection[1])

for color, number in data.items():
    print(f"Color: {color}")
    print(f"Color: {color}, Number: {number}")

'''
import cv2
import easyocr


image_path = "C:\Local\Project\convert_pdf_to_png_page_1_out.png"
image = cv2.imread(image_path)

reader = easyocr.Reader(['en'])  
result = reader.readtext(image)

data = {}
for i in range(len(result) - 1):
    current_text = result[i][1].strip().lower()
    next_text = result[i + 1][1].strip()

    if current_text in ["green", "pink", "purple"]:
        if next_text.isdigit():
            data[current_text] = int(next_text)

# Print the extracted color-value pairs
for color, value in data.items():
    print(f"Color: {color}, Value: {value}")
