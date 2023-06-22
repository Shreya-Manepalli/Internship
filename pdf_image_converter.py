import io
import aspose.pdf as ap

input_pdf = "C:\Local\Project\PDFImages\I286A3309-2369-WRR0061.pdf"
output_pdf = "convert_pdf_to_png"

# Open PDF document
document = ap.Document(input_pdf)

# Create Resolution object
resolution = ap.devices.Resolution(300)
device = ap.devices.PngDevice(resolution)

for i in range(0, len(document.pages)):
    # Create file for save
    imageStream = io.FileIO(
        output_pdf + "_page_" + str(i + 1) + "_out.png", 'x'
    )
    
    # Convert a particular page and save the image to stream
    device.process(document.pages[i + 1], imageStream)
    imageStream.close()