'''
import cv2

img = cv2.imread("images-pdf\scatterplot-2.png")
gray_image = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
thresh_image = cv2.threshold(gray_image, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

# Find contours
contours, _ = cv2.findContours(thresh_image, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

# Create a list to store individual character bounding boxes
bounding_boxes = []

# Iterate over contours
for contour in contours:
    # Get bounding rectangle coordinates
    x, y, w, h = cv2.boundingRect(contour)
    
    # Add some padding to the bounding box
    padding = 5
    x -= padding
    y -= padding
    w += 2 * padding
    h += 2 * padding
    
    # Append the bounding box coordinates to the list
    bounding_boxes.append((x, y, w, h))

# Draw bounding boxes around the characters
for box in bounding_boxes:
    x, y, w, h = box
    cv2.rectangle(img, (x, y), (x + w, y + h), (255, 0, 0), 4)

cv2.imshow('Bounding Boxes', img)
cv2.waitKey(0)
'''

import cv2

img = cv2.imread("images-pdf\scatterplot-2.png")
gray_image = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
thresh_image = cv2.threshold(gray_image, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

# Find contours
contours, _ = cv2.findContours(thresh_image, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

# Create a list to store bounding boxes of filled circles
bounding_boxes = []

# Iterate over contours
for contour in contours:
    # Fit a circle around the contour
    (x, y), radius = cv2.minEnclosingCircle(contour)
    
    # Calculate the circularity of the contour
    circularity = 4 * 3.1415 * cv2.contourArea(contour) / (cv2.arcLength(contour, True) ** 2)
    
    # Filter only the filled circles based on circularity
    if circularity > 0.8:  # You can adjust this threshold for circularity based on your images
        # Calculate bounding box coordinates based on circle's center and radius
        x, y, radius = int(x), int(y), int(radius)
        padding = int(radius * 0.5)  # Add padding to the bounding box
        x -= padding
        y -= padding
        w = h = int(2 * radius) + 2 * padding
        
        # Append the bounding box coordinates to the list
        bounding_boxes.append((x, y, w, h))

# Draw bounding boxes around each filled circle
for box in bounding_boxes:
    x, y, w, h = box
    cv2.rectangle(img, (x, y), (x + w, y + h), (255, 0, 0), 4)

cv2.imshow('Bounding Boxes', img)
cv2.waitKey(0)

