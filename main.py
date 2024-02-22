#import the necessary libraries</pre> 
from PIL import ImageFont, ImageDraw, Image
import numpy as np

#PyQT libraries
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import QApplication

app = QApplication([])  # Create a QApplication object

import cv2 as cv 
import openpyxl 
import os

##Second imports
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QDesktopWidget
from PyQt5.QtGui import QImage, QPixmap
from PyQt5.uic import loadUi
import sys
import cv2
import numpy as np

from fontTools.ttLib import TTFont
from fontTools.varLib import models

from fontTools.varLib import mutator
# 

# # Load the variable font
# font = TTFont("FONTS/Montserrat-VariableFont_wght.ttf")

# # Get the 'wght' (weight) axis
# for axis in font["fvar"].axes:
#     if axis.axisTag == 'wght':
#         wght_axis = axis
#         break

# # Define the desired variation
# variation = {"wght": 700}  # 700 is commonly used for bold

# # Create the specific instance
# bold_font = mutator.instantiateVariableFont(font, variation)

# # Save the bold font to a new TTF file
# bold_font.save("FONTS/Montserrat-Bold.ttf")

# print("Font saved successfully")

def create_bold_font():
    bold_font_file = "FONTS/Montserrat-Bold.ttf"

    # Check if the bold font file already exists
    if os.path.exists(bold_font_file):
        print("Bold font file already exists.")
        return

    # Load the variable font
    font = TTFont("FONTS/Montserrat-VariableFont_wght.ttf")

    # Get the 'wght' (weight) axis
    for axis in font["fvar"].axes:
        if axis.axisTag == 'wght':
            wght_axis = axis
            break

    # Define the desired variation
    variation = {"wght": 500}  # 700 is commonly used for bold

    # Create the specific instance
    bold_font = mutator.instantiateVariableFont(font, variation)

    # Save the bold font to a new TTF file
    bold_font.save(bold_font_file)

    print("Bold font file created successfully.")

# Call the function
create_bold_font()

image_Path = 'Reconocimiento-BREAS copy.png'

# Load image in OpenCV
image = cv.imread('Reconocimiento-BREAS copy.png')

# Convert the image to RGB (OpenCV uses BGR)
cv_img = cv.cvtColor(image, cv.COLOR_BGR2RGB)

# Convert the OpenCV image to PIL image
pil_img = Image.fromarray(cv_img)
draw = ImageDraw.Draw(pil_img)

# Use a truetype font
try:
    font = ImageFont.truetype("FONTS/Montserrat-Bold.ttf", 220)
    
    print("Font loaded successfully")
except IOError:
    print("Error loading font")

print(font, " <<<<< FONTLoaded")

# Draw text
# First Get width and height of the image
pixmap = QPixmap(image_Path)  # Load the image
if pixmap.isNull():
    print("Failed to load image.")
else:
    print("Image loaded successfully.")


width = pixmap.width()  # Get the width of the image
height = pixmap.height()  # Get the height of the image

print("Width:", width, "Height:", height)
draw.text((width/2, height*.66), "Hello World 2", anchor="ms",  font=font, fill=(255,255,255), align = 'center')
print(draw.text, " <<<<< This is just an example")

# Convert back to OpenCV image and save
cv_img = cv.cvtColor(np.array(pil_img), cv.COLOR_RGB2BGR)
cv.imwrite('output.png', cv_img)

# Get the aspect ratio of the image
aspect_ratio = cv_img.shape[1] / cv_img.shape[0]  # width / height
print(cv_img.shape[1], " > 1 /n", cv_img.shape[0], " > 0 /n", aspect_ratio, " > aspect ratio of the image")

# preview image
cv.namedWindow('window', cv.WINDOW_NORMAL)
cv.imshow('window', cv_img)
# Resize the window
cv.resizeWindow('window', int(width / 2), int(height / 2))  # Width and height

cv.waitKey(0)  # Keep the window open


print(os.getcwd(), "shown preview")
	
# template1.png is the template 
# certificate 
template_path = 'template01.png'

# Excel file containing names of 
# the participants 
details_path = 'Book1.xlsx'

# Output Paths 
output_path = 'output/'



# Setting the font size and font 
# colour 
font_size = 32
font_color = (0,255,0) 

# Coordinates on the certificate where 
# will be printing the name (set 
# according to your own template) 
coordinate_y_adjustment = 15
coordinate_x_adjustment = 7

# loading the details.xlsx workbook 
# and grabbing the active sheet 
obj = openpyxl.load_workbook(details_path) 
sheet = obj.active 

# printing for the first 10 names in the 
# excel sheet 
for i in range(1,11): 
	
	# grabs the row=i and column=1 cell 
	# that contains the name value of that 
	# cell is stored in the variable certi_name 
	get_name = sheet.cell(row = i ,column = 1) 
	certi_name = get_name.value 
							
	# read the certificate template 
	img = cv.imread(template_path) 
								
	# choose the font from opencv 
	font = cv.FONT_HERSHEY_PLAIN			 

	# get the size of the name to be 
	# printed 
	text_size = cv.getTextSize(certi_name, font, font_size, 10)[0]	 

	# get the (x,y) coordinates where the 
	# name is to written on the template 
	# The function cv.putText accepts only 
	# integer arguments so convert it into 'int'. 
	text_x = (img.shape[1] - text_size[0]) / 2 + coordinate_x_adjustment 
	text_y = (img.shape[0] + text_size[1]) / 2 - coordinate_y_adjustment 
	text_x = int(text_x) 
	text_y = int(text_y) 
	cv.putText(img, certi_name, 
			(text_x ,text_y ), 
			font, 
			font_size, 
			font_color, 10) 

	# Output path along with the name of the 
	# certificate generated 
	certi_path = output_path + '/certi' + '.png'
	
	 # Save the certificate	

# print('Certificate of ', certi_name, ' created',img.shape)
# cv.imshow('Image', img)
# cv.waitKey(0)
# cv.destroyAllWindows()		
# cv.imwrite(certi_path,img)



from UI import MainWindow  # Import the UI file

# Create the PyQt application
app = QApplication(sys.argv)

# Create and show the PyQt window
window = MyApp()
window.show()

sys.exit(app.exec_())