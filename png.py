from docx import Document
from docx.shared import Inches
from PIL import ImageGrab
import io

# Open the document
doc = Document("Invoice_Template_No_Borders.docx")

# Capture a screenshot of the screen (adjust coordinates)
screenshot = ImageGrab.grab(bbox=(0, 0, 50, 50))  # Top-left (x1,y1), Bottom-right (x2,y2)

# Save screenshot to a temporary file
temp_img = "temp_screenshot.png"
screenshot.save(temp_img)

# Add the image to the document
doc.add_picture(temp_img, width=Inches(3.0))  # Adjust width as needed

# Save the modified document
doc.save("modified_document.docx")