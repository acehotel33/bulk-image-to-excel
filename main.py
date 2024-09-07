from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
import os

# Create a new workbook and select the active sheet
wb = Workbook()
ws = wb.active

# Set square cell dimensions to 200x200 pixels
cell_size = 200  # 200 pixels width and height for a square

# Function to set column width (in Excel's measurement units)
def set_column_width(sheet, col, width_px):
    col_letter = get_column_letter(col)
    sheet.column_dimensions[col_letter].width = width_px / 7  # Roughly 1 Excel unit = 7 pixels

# Function to set row height (in points)
def set_row_height(sheet, row, height_px):
    sheet.row_dimensions[row].height = height_px * 0.75  # 1 point is approximately 0.75 pixels

# Folder containing images
image_folder = "./images"

# Iterate through image files in the folder, starting from B2
start_row = 2
start_col = 2
row = start_row
col = start_col

for filename in os.listdir(image_folder):
    if filename.endswith(".png") or filename.endswith(".jpg") or filename.endswith(".jpeg"):
        img_path = os.path.join(image_folder, filename)
        
        # Open the image using the Image class
        img = Image(img_path)

        # Set the column width and row height to ensure square cells
        set_column_width(ws, col, cell_size)
        set_row_height(ws, row, cell_size)

        # Resize the image to fit within the cell
        img.width = cell_size
        img.height = cell_size
        
        # Insert the image into the worksheet at the top-left of the cell
        cell_location = f"{get_column_letter(col)}{row}"
        ws.add_image(img, cell_location)

        # Move to the next row for the next image
        row += 1

# Save the workbook
wb.save("images_in_square_cells_B2.xlsx")

