'''
import cv2
import matplotlib.pyplot as plt
import pandas as pd
import os


def on_click(event, points):
    if event.button == 1:  # Left mouse button click
        x, y = event.xdata, event.ydata
        points.append((x, y))
        ax.plot(x, y, 'ro')  # Plot a red dot at the clicked point
        plt.draw()

def digitize_scatterplot(image_path):
    # Load the image using OpenCV
    img = cv2.imread(image_path)

    # Convert the image from BGR to RGB (Matplotlib expects RGB format)
    img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

    # Plot the image for interactive digitization
    points = []  # Local variable to store data points
    global ax
    fig, ax = plt.subplots()
    ax.imshow(img_rgb)
    ax.set_title("Digitize the scatterplot by clicking on data points.")
    ax.set_xlabel("X-axis")
    ax.set_ylabel("Y-axis")
    plt.connect('button_press_event', lambda event: on_click(event, points))
    plt.show()

    # Create a Pandas DataFrame to store the coordinates
    df = pd.DataFrame(points, columns=["X", "Y"])

    return df

if __name__ == "__main__":
    folder_path = input("Enter the folder path containing the graphs: ")  # Replace with the path to your folder containing graphs

    # Extract the folder name
    folder_name = os.path.basename(folder_path)

    # Create an Excel writer with the folder name
    excel_filename = f"{folder_name}_data.xlsx"
    writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')

    # Process each image in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".png"):  # Adjust the extension based on your image files
            image_path = os.path.join(folder_path, filename)
            df = digitize_scatterplot(image_path)

            # Extract the image name without the file extension
            image_name = os.path.splitext(filename)[0]

            # Save the DataFrame to a separate sheet in the Excel file
            df.to_excel(writer, sheet_name=image_name, index=False)

    # Save and close the Excel writer
    writer.save()

    print("Data points extracted and saved to", excel_filename)
'''
import cv2
import numpy as np
#open the main image and convert it to gray scale image
main_image = cv2.imread('images-pdf\scatterplot-2.png')
gray_image = cv2.cvtColor(main_image, cv2.COLOR_BGR2GRAY)
#open the template as gray scale image
template = cv2.imread('images-pdf\get2.png', 0)
width, height = template.shape[::-1] #get the width and height
#match the template using cv2.matchTemplate
match = cv2.matchTemplate(gray_image, template, cv2.TM_CCOEFF_NORMED)
threshold = 0.8
position = np.where(match >= threshold) #get the location of template in the image
for point in zip(*position[::-1]): #draw the rectangle around the matched template
   cv2.rectangle(main_image, point, (point[0] + width, point[1] + height), (0, 204, 153), 0)
cv2.imshow('Template Found', main_image)
cv2.waitKey(0)