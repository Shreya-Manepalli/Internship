from PIL import Image

def crop_image_into_chunks(image_path):
    # Open the image
    image = Image.open(image_path)
    
    # Get the dimensions of the image
    width, height = image.size
    
    # Calculate the size of each chunk
    chunk_x = width // 3
    chunk_y = height // 3
    
    # Create a list to store the cropped chunks
    chunks = []
    
    # Iterate over each row and column
    for i in range(3): # height
        for j in range(3): #width
            # Calculate the coordinates for cropping
            left = j * chunk_x
            upper = i * chunk_y
            right = left + chunk_x
            lower = upper + chunk_y
            
            # Crop the image using the coordinates
            chunk = image.crop((left, upper, right, lower))
            
            # Add the cropped chunk to the list
            chunks.append(chunk)
    
    # Return the list of cropped chunks
    return chunks

# Path to the input image
image_path = 'output_page_0.png'

# Crop the image into chunks
cropped_chunks = crop_image_into_chunks(image_path)

# Save each cropped chunk
for i, chunk in enumerate(cropped_chunks):
    chunk.save(f'output\chunk_{i+1}.png')
