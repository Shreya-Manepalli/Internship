# import easyocr

# # Provide the path to the model file
# model_path = 'new_model.pth'

# # Initialize the reader with the 'en' language for English and the custom model
# reader = easyocr.Reader(['en'], model_storage_directory=model_path)

# # Provide the path to the image you want to perform OCR on
# image_path = 'output_page_0.png'

# # Read the text from the image
# result = reader.readtext(image_path, detail=0)

# # Process the result+
# for text in result:
#     print(text)

import easyocr
import os

# Provide the path to the model file
model_path = 'new_model.pth'

# Initialize the reader with the 'en' language for English and the custom model
reader = easyocr.Reader(['en'], model_storage_directory=model_path)

# Provide the path to the folder containing the images
folder_path = 'output'  # Replace with the actual folder path

# Get the list of files in the folder
files = os.listdir(folder_path)

# Filter and process each image file
for file in files:
    # Check if the file is an image (you can customize this condition based on your specific requirements)
    if file.endswith(('.jpg', '.png', '.jpeg')):
        # Construct the full image path
        image_path = os.path.join(folder_path, file)
        
        # Read the text from the image
        result = reader.readtext(image_path, detail=0)
        
        # Process the result
        for text in result:
            print(text)
        
        print('---')  # Separator between images

