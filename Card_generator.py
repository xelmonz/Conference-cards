import os
import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
from PyPDF2 import PdfMerger

# 1. Generate QR Codes from vCard Data
def generate_qr_code(vcard_data, file_name):
    # Create a QR code from the vCard data
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(vcard_data)
    qr.make(fit=True)

    # Create an image from the QR code
    img = qr.make_image(fill='black', back_color='white')

    # Create a directory to save the image
    output_dir = "qr_codes"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Full path of the saved image
    file_path = os.path.join(output_dir, f"{file_name}.png")

    # Save the image with the filename based on the `FN` field
    img.save(file_path)
    print(f"QR code saved at: {file_path}")


# Load the Excel file
# Replace 'your_excel_file.xlsx' with the path to your Excel file
excel_file = 'final.xlsx'
df = pd.read_excel(excel_file)

# Iterate over the rows of the Excel file
for index, row in df.iterrows():
    # Extract data from the Excel columns
    full_name = row['FN']  # Assuming the name is in a column named 'FN'
    email = row['EMAIL']  # Assuming email is in a column named 'EMAIL'
    org = row['ORG']  # Assuming organization is in a column named 'ORG'

    # Construct the vCard data for each row
    vcard_data = f"""BEGIN:VCARD
VERSION:3.0
FN:{full_name}
EMAIL:{email}
ORG:{org}
END:VCARD"""

    # Generate the QR code and save it as a PNG file
    generate_qr_code(vcard_data, full_name)

#_________________________________________________________________________________________________________

# 2. Add QR Codes to Back Pages

# Folder paths
qr_code_folder = 'qr_codes'  # Folder containing your 300 QR code images
main_card_image_path = 'Identification card_page-0002.png'  # Path to your main card image
output_folder = 'Final pics'  # Folder to save the final images

# Create output folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Define the position where the QR code should be placed on the card
qr_position = (610, 128)

# Define the size to which QR codes should be resized
qr_size = (210, 210)

# Open the main card image
main_card_image = Image.open(main_card_image_path).convert('RGB')

# Iterate over the QR code images in the folder
for qr_filename in os.listdir(qr_code_folder):
    if qr_filename.endswith('.jpg') or qr_filename.endswith('.png'):  # Make sure only images are processed
        # Open the QR code image
        qr_image_path = os.path.join(qr_code_folder, qr_filename)
        qr_image = Image.open(qr_image_path)

        # Resize the QR code image to the specified size
        qr_image_resized = qr_image.resize(qr_size)

        # Create a copy of the main card image to modify
        card_with_qr = main_card_image.copy()

        # If QR image has an alpha channel (transparency), use it as a mask
        if qr_image_resized.mode == 'RGBA':
            # Split the QR image into RGB and alpha (transparency) channels
            qr_rgb, qr_alpha = qr_image_resized.split()[:3], qr_image_resized.split()[3]

            # Paste the RGB QR code using its alpha channel as a mask
            card_with_qr.paste(qr_image_resized.convert('RGB'), qr_position, mask=qr_alpha)
        else:
            # No alpha channel, paste directly
            card_with_qr.paste(qr_image_resized, qr_position)

        # Define the output path for the final image
        output_path = os.path.join(output_folder, qr_filename)  # Save the image with the same name as the QR code

        # Save the final image
        card_with_qr.save(output_path)

        print(f"Saved: {output_path}")

print("All images processed.")

#_________________________________________________________________________________________________________

# 3. Create Front Pages with Name and Country
# File paths
excel_file = 'final.xlsx'  # Excel file containing FN and Country columns
main_image_path = 'Identification card_page-0001.png'  # Path to the main image for the front page
output_folder = 'Final pics'  # Folder to save the final images

# Create output folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Load the Excel data
df = pd.read_excel(excel_file)

# Define the Y positions for the name and country
name_y_position = 570  # Y position for the name (vertical position)
country_y_position = 640  # Y position for the country (vertical position)

# Font settings (load a custom TTF font and specify sizes)
# Adjust the path and size as needed
name_font = ImageFont.truetype("arial.ttf", 56)  # Font for Full Name with size 36
country_font = ImageFont.truetype("arial.ttf", 40)  # Font for Country with size 24

# Load one of the card images to get its dimensions (adjust based on your card image size)
main_image = Image.open(main_image_path)

# Ensure the main image is in RGB mode (supports full color)
main_image = main_image.convert('RGB')

# Get the dimensions of the image
card_width, card_height = main_image.size

# Iterate over the rows of the Excel file
for index, row in df.iterrows():
    full_name = row['FN']
    country = row['Country']

    # Create a copy of the main image for each row
    card_image = main_image.copy()

    # Initialize ImageDraw to write text
    draw = ImageDraw.Draw(card_image)

    # Measure the width of the name text using textbbox (Pillow 8.0.0+)
    name_bbox = draw.textbbox((0, 0), full_name, font=name_font)
    name_text_width = name_bbox[2] - name_bbox[0]

    # Calculate the X position to center the name
    name_x_position = (card_width - name_text_width) / 2

    # Add Full Name at the calculated X and specified Y position
    draw.text((name_x_position, name_y_position), full_name, font=name_font, fill='black')

    # Measure the width of the country text using textbbox (Pillow 8.0.0+)
    country_bbox = draw.textbbox((0, 0), country, font=country_font)
    country_text_width = country_bbox[2] - country_bbox[0]

    # Calculate the X position to center the country
    country_x_position = (card_width - country_text_width) / 2

    # Add Country at the calculated X and specified Y position
    draw.text((country_x_position, country_y_position), country, font=country_font, fill='black')

    # Output path for the final card front page
    output_path = os.path.join(output_folder, f"{full_name}_front_page.png")

    # Save the final image
    card_image.save(output_path)

    print(f"Saved: {output_path}")

print("All images processed.")

#_________________________________________________________________________________________________________

# 4. Merge png files into one pdf

# Set the folder where your PNG files are located
image_folder = 'Final pics'
output_folder = 'Final'

# Create output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Get a list of all files in the folder
image_files = os.listdir(image_folder)

# Filter out the front and back PNGs
front_files = [f for f in image_files if '_front_page' in f]
back_files = [f for f in image_files if '_front_page' not in f and f.endswith('.png')]

for back_file in back_files:
    # Derive the base name without the extension and without '_front_page'
    base_name = os.path.splitext(back_file)[0]

    # Derive the front page filename
    front_file = f"{base_name}_front_page.png"

    # Check if corresponding front file exists
    if front_file in front_files:
        # Open the front and back images
        front_image = Image.open(os.path.join(image_folder, front_file))
        back_image = Image.open(os.path.join(image_folder, back_file))

        # Convert both images to PDF
        front_pdf_path = os.path.join(output_folder, f"{base_name}_front.pdf")
        back_pdf_path = os.path.join(output_folder, f"{base_name}_back.pdf")

        front_image.save(front_pdf_path, "PDF", resolution=100.0)
        back_image.save(back_pdf_path, "PDF", resolution=100.0)

        # Merge the front and back PDF into a single PDF
        merger = PdfMerger()
        merger.append(front_pdf_path)
        merger.append(back_pdf_path)

        # Output the final PDF
        output_pdf_path = os.path.join(output_folder, f"{base_name}.pdf")
        with open(output_pdf_path, "wb") as f_out:
            merger.write(f_out)

        # Clean up intermediate PDFs
        merger.close()
        os.remove(front_pdf_path)
        os.remove(back_pdf_path)

        print(f"Merged: {base_name}.pdf")
    else:
        print(f"Front page not found for: {base_name}")

print("All PDFs created.")

