'''
from PIL import Image
Image.MAX_IMAGE_PIXELS = None 
def crop_image_into_chunks(image_path):
    image = Image.open(image_path)
    width, height = image.size
    chunk_x = width // 12
    chunk_y = height // 12
    padding = 380
    chunks = []
    for i in range(12):  
        for j in range(12): 
            left = j * chunk_x
            upper = i * chunk_y
            right = left + chunk_x
            lower = upper + chunk_y
            left_padding = padding if j > 0 else 0
            right_padding = padding if j < 2 else 0
            left -= left_padding
            right += right_padding
            chunk = image.crop((left, upper, right, lower))
            chunks.append(chunk)
    return chunks

image_path = 'output_page_0.jpg'
cropped_chunks = crop_image_into_chunks(image_path)
for i, chunk in enumerate(cropped_chunks):
    chunk.save(f'output/chunk_{i+1}.jpg')
'''
'''
from PIL import Image
import cv2
import numpy as np

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
    padding = 480
    padding_top_bottom = 380
    chunks = []
    for i in range(14):
        for j in range(14):
            left = j * chunk_x
            upper = i * chunk_y
            right = left + chunk_x
            lower = upper + chunk_y
            left_padding = padding if j > 0 else 0
            right_padding = padding if j < 2 else 0
            top_padding = padding_top_bottom if i > 0 else 0
            bottom_padding = padding_top_bottom if i < 2 else 0
            left -= left_padding
            right += right_padding
            upper -= top_padding
            lower += bottom_padding
            chunk = image.crop((left, upper, right, lower))
            if has_data(chunk):
                chunks.append(chunk)
    return chunks


image_path = 'output_page_0.jpg'
cropped_chunks = crop_image_into_chunks(image_path)
for i, chunk in enumerate(cropped_chunks):
    chunk.save(f'output/chunk_{i+1}.jpg')
    '''

from PIL import Image
import cv2
import numpy as np

Image.MAX_IMAGE_PIXELS = None

def detect_triangles(chunk):
    # Convert the chunk to grayscale and apply Canny edge detection
    chunk_gray = chunk.convert('L')
    edges = cv2.Canny(np.array(chunk_gray), 50, 150)

    # Find contours in the edge map
    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    # Count the number of triangles
    triangle_count = 0
    for contour in contours:
        approx = cv2.approxPolyDP(contour, 0.03 * cv2.arcLength(contour, True), True)
        if len(approx) == 3:
            triangle_count += 1

    return triangle_count

def has_data(chunk):
    # Convert the chunk to grayscale and apply Canny edge detection
    chunk_gray = chunk.convert('L')
    edges = cv2.Canny(np.array(chunk_gray), 50, 150)

    # Find contours in the edge map
    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    # Filter contours based on area and aspect ratio
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
    padding = 480
    padding_top_bottom = 380
    chunks = []
    triangle_counts = []
    for i in range(14):
        for j in range(14):
            left = j * chunk_x
            upper = i * chunk_y
            right = left + chunk_x
            lower = upper + chunk_y
            left_padding = padding if j > 0 else 0
            right_padding = padding if j < 2 else 0
            top_padding = padding_top_bottom if i > 0 else 0
            bottom_padding = padding_top_bottom if i < 2 else 0
            left -= left_padding
            right += right_padding
            upper -= top_padding
            lower += bottom_padding
            chunk = image.crop((left, upper, right, lower))
            if has_data(chunk):
                chunks.append(chunk)
                triangle_count = detect_triangles(chunk)
                triangle_counts.append(triangle_count)
    return chunks, triangle_counts
image_path = 'output_page_0.jpg'
cropped_chunks = crop_image_into_chunks(image_path)
for i, chunk in enumerate(cropped_chunks):
    chunk.save(f'output/chunk_{i+1}.jpg')