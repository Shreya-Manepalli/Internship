import cv2
import numpy as np
#open the main image and convert it to gray scale image
main_image = cv2.imread('images-pdf\scatterplot-2.png')
gray_image = cv2.cvtColor(main_image, cv2.COLOR_BGR2GRAY)
#open the template as gray scale image
template = cv2.imread('images-pdf\get.png', 0)
width, height = template.shape[::-1] #get the width and height
#match the template using cv2.matchTemplate
match = cv2.matchTemplate(gray_image, template, cv2.TM_CCOEFF_NORMED)
threshold = 0.8
position = np.where(match >= threshold) #get the location of template in the image
for point in zip(*position[::-1]): #draw the rectangle around the matched template
   cv2.rectangle(main_image, point, (point[0] + width, point[1] + height), (0, 204, 153), 0)
cv2.imshow('Template Found', main_image)
cv2.waitKey(0)