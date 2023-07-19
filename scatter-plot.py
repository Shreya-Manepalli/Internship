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
