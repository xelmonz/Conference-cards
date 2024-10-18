Participant Card Generator

Overview:
This project automates the creation of personalized identification cards for participants using data from an Excel file. 
Each card features a unique QR code, along with the participant's name and organization, making it a versatile tool for conferences, events, or any gathering where participant identification is essential.

Features:
- Automatic QR Code Generation: Utilizes the qrcode library to create scannable QR codes based on participant information.
- Customizable Card Design: Supports customizable front and back designs using PIL (Python Imaging Library) for a personalized touch.
- Batch Processing: Efficiently processes multiple participants in one go, reading data directly from an Excel file for seamless integration.
- PDF Generation: Merges front and back pages of cards into a single PDF file for easy printing and distribution.

Instructions:
Replace the placeholder file paths (like 'final.xlsx', 'Identification card_page-0001.png', etc.) with the actual paths to your files.
Ensure that you have the required dependencies installed (qrcode, pandas, Pillow, and PyPDF2).

