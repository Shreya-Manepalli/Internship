from PIL import Image

def convert_to_black_and_white(input_image_path, output_image_path):
    try:
        # Open the image
        image = Image.open(input_image_path)

        # Convert the image to grayscale (black and white)
        black_and_white_image = image.convert("L")

        # Save the black and white image
        black_and_white_image.save(output_image_path)

        print("Image converted to black and white successfully!")
    except Exception as e:
        print("An error occurred:", e)

if __name__ == "__main__":
    input_image_path = "images-pdf\scatterplot-2.png"   # Replace with the path to your input image
    output_image_path = "output_image.jpg" # Replace with the desired output path
    convert_to_black_and_white(input_image_path, output_image_path)