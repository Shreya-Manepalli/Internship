from PIL import Image
Image.MAX_IMAGE_PIXELS = None 
def crop_image_into_chunks(image_path):
    image = Image.open(image_path)
    width, height = image.size
    chunk_x = width // 3
    chunk_y = height // 3
    padding = 380
    chunks = []
    for i in range(3):  
        for j in range(3): 
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
