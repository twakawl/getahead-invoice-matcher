from pdf2image import convert_from_path
import pytesseract
import re
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

pdf_path = "input/invoices/FXX-BakkerijRenzema.jpeg"
# images = convert_from_path(str(pdf_path))
images =[pdf_path]
logging.info(f"Processing PDF: {pdf_path}")
text = "".join(pytesseract.image_to_string(image) for image in images)
logging.info(text[:500])  # Log first 500 characters of text for debugging
text = re.sub(r" ", "", text).lower()
logging.info(text[:500])  # Log first 500 characters of cleaned text for debugging


