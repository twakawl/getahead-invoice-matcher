from PIL import Image
import pytesseract
import re
import logging
from pathlib import Path

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

image_path = Path("input/invoices/FXX-BakkerijRenzema.jpeg")

if not image_path.exists():
    logging.error(f"Test image not found at {image_path}")
else:
    logging.info(f"Processing image: {image_path}")
    try:
        image = Image.open(image_path)
        text = pytesseract.image_to_string(image)
        logging.info(text[:500])  # Log first 500 characters of text for debugging
        text = re.sub(r" ", "", text).lower()
        logging.info(text[:500])  # Log first 500 characters of cleaned text for debugging
    except Exception as e:
        logging.error(f"An error occurred during OCR: {e}")
