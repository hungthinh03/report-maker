from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker, XDRPositiveSize2D
from datetime import datetime
import os, zipfile


# == USER SETTINGS ==
START_ROW = 11               # default = 11
START_COL = 9                # column index -> J=9, A=0, B=1, C=2,..
NEW_WIDTH = 200              # default = 200 
MAX_HEIGHT = 350             # default = 350 
SPACING = 10                 # default = 10
IMAGES_PER_ROW = 3           # default = 3


# == APPLICATION ==
# 1. Find template Excel file
BASE_FOLDER = "default"   # folder holding your template .xlsx file
IMG_FOLDER = "images"        # folder for images or a .zip file

excel_file = None
for file in os.listdir(BASE_FOLDER):
    if file.lower().endswith(".xlsx"):
        excel_file = os.path.join(BASE_FOLDER, file)
        break

if not excel_file:
    raise FileNotFoundError("\n\nNo .xlsx template file found in base_excel/!\n\n")

# 2. Check / unzip images
img_files = [f for f in os.listdir(IMG_FOLDER) if f.lower().endswith((".png", ".jpg", ".jpeg"))]

if not img_files:
# no images -> look for zip
    for f in os.listdir(IMG_FOLDER):
        if f.lower().endswith(".zip"):
            zip_path = os.path.join(IMG_FOLDER, f)
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(IMG_FOLDER)
            os.remove(zip_path)  # delete the zip after extraction
            print(f"\n\nUnzipped and deleted {f} into {IMG_FOLDER}\n\n")
            break

# Refresh after unzip
img_files = [f for f in os.listdir(IMG_FOLDER) if f.lower().endswith((".png", ".jpg", ".jpeg"))]

if not img_files:
    raise FileNotFoundError("\nNo images found in /images (even after unzipping)!\n")

# 3. Prepare output name = template-name + timestamp
template_name = os.path.splitext(os.path.basename(excel_file))[0]
timestamp = datetime.now().strftime("%d%m%H%M%S")
output_file = f"{template_name}-{timestamp}.xlsx"

# 4. Load workbook and active sheet
wb = load_workbook(excel_file)
ws = wb.active

# Clear existing images (if any)
ws._images = []

# 5. Insert images
row = START_ROW
col = START_COL
images = sorted(img_files)

for i in range(0, len(images), IMAGES_PER_ROW):  # group images in threes
    group = images[i:i+IMAGES_PER_ROW]
    offset_x = 0

    for file in group:
        img_path = os.path.join(IMG_FOLDER, file)
        img = Image(img_path)

        # Resize by width and keep aspect ratio
        aspect_ratio = img.height / img.width
        img.width = NEW_WIDTH
        img.height = int(NEW_WIDTH * aspect_ratio)

        # Limit max height
        if img.height > MAX_HEIGHT:
            scale = MAX_HEIGHT / img.height
            img.width = int(img.width * scale)
            img.height = MAX_HEIGHT

        # Anchor with proper size
        marker = AnchorMarker(col=col,colOff=(offset_x + SPACING) * 9525, row=row-1, rowOff=SPACING * 9525)
        size = XDRPositiveSize2D(cx=img.width*9525, cy=img.height*9525)
        img.anchor = OneCellAnchor(_from=marker, ext=size)

        ws.add_image(img)

        # Increment offset for next image
        offset_x += img.width + SPACING

    row += 1  # move to next row

# 6. Save as new file
wb.save(output_file)
print(f"\n\nDone! Images inserted into {output_file}\n\n")
