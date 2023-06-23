import cv2

img = cv2.imread("C:\Local\Project\Images\convert_pdf_to_png_page_3_out.png")
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
import numpy as np

def merge_small_boxes(boxes, size_thresh):
    # If there are no boxes, return an empty list
    if len(boxes) == 0:
        return []

    # Convert the bounding boxes to numpy array
    boxes = np.array(boxes)

    # Calculate the areas of the boxes
    areas = boxes[:, 2] * boxes[:, 3]

    # Find the indices of small boxes
    small_indices = np.where(areas <= size_thresh)[0]

    # If there are no small boxes, return the original boxes
    if len(small_indices) == 0:
        return boxes.tolist()

    # Calculate the coordinates of the merged box
    x = np.min(boxes[small_indices, 0])
    y = np.min(boxes[small_indices, 1])
    w = np.max(boxes[small_indices, 0] + boxes[small_indices, 2]) - x
    h = np.max(boxes[small_indices, 1] + boxes[small_indices, 3]) - y

    # Create the merged box
    merged_box = np.array([x, y, w, h])

    # Remove the small boxes from the original boxes
    boxes = np.delete(boxes, small_indices, axis=0)

    # Append the merged box to the boxes
    boxes = np.vstack([boxes, merged_box])

    # Return the updated boxes
    return boxes.tolist()

# Load the image
img = cv2.imread("C:\Local\Project\Images\convert_pdf_to_png_page_3_out.png")

# Convert to grayscale
gray_image = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

# Threshold the image
_, thresh_image = cv2.threshold(gray_image, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

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

# Merge small bounding boxes with a size threshold of 100 pixels
merged_boxes = merge_small_boxes(bounding_boxes, size_thresh=200)

# Draw merged bounding boxes on the image
for box in merged_boxes:
    x, y, w, h = box
    cv2.rectangle(img, (x, y), (x + w, y + h), (255, 0, 0), 4)

# Display the image with merged bounding boxes
cv2.imshow('Merged Bounding Boxes', img)
cv2.waitKey(0)
cv2.destroyAllWindows()

'''