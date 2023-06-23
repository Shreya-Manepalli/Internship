import fitz
from PIL import Image

# Path of the PDF file
pdf_path = 'Sample Dwg 1.pdf'

# Define DPI (adjust as needed)
dpi = 900

# Convert PDF to images
doc = fitz.open(pdf_path)
for page_index in range(len(doc)):
    page = doc[page_index]
    zoom = dpi / 72.0
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    output_path = f'output_page_{page_index}.png'
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    img.save(output_path, dpi=(dpi, dpi))

# Close the PDF file
doc.close()
