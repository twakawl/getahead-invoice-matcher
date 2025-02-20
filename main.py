import os
import re
import json
import logging
from typing import Dict, Optional, List
import pandas as pd
import pytesseract
from PIL import Image
from pdf2image import convert_from_path
from PyPDF2 import PdfReader
import openai

# Setup logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

# Load settings
with open("settings.json", "r") as f:
    settings = json.load(f)

INPUT_EXCEL = settings["input_excel"]
OUTPUT_EXCEL = settings["output_excel"]
PDF_DIR = settings["pdf_directory"]
OUTPUT_PDF_DIR = settings["output_pdf_directory"]
EXTRACTION_METHOD = settings["extraction_method"]  # Options: "standard", "ocr", "gpt"
TESSERACT_PATH = settings.get("tesseract_path")  # Optional
OPENAI_API_KEY = settings.get("openai_api_key")

# Configure Tesseract executable path if required
if TESSERACT_PATH:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

# Configure OpenAI API key if required
if OPENAI_API_KEY:
    openai.api_key = OPENAI_API_KEY


class InvoiceMatcher:
    def __init__(self, excel_path: str, pdf_dir: str, output_pdf_dir: str, method: str):
        """
        Initialize the InvoiceMatcher class.

        Args:
            excel_path (str): Path to the input Excel file.
            pdf_dir (str): Directory containing the PDF files.
            output_pdf_dir (str): Directory to store renamed PDF files.
            method (str): Extraction method ("standard", "ocr", "gpt").
        """
        self.excel_path = excel_path
        self.pdf_dir = pdf_dir
        self.output_pdf_dir = output_pdf_dir
        self.method = method
        self.data = self.load_excel()
        self.extracted_data = []
        self.all_combined_hits = pd.DataFrame()

    def load_excel(self) -> pd.DataFrame:
        """Load the Excel file into a Pandas DataFrame."""
        logging.info("Loading Excel file...")
        df = pd.read_excel(self.excel_path)
        df.columns = df.columns.str.strip()  # Clean column names
        return df

    def extract_invoice_data(self, pdf_path: str) -> Optional[Dict[str, str]]:
        """
        Extract invoice data from a PDF using the specified method.

        Args:
            pdf_path (str): Path to the PDF file.

        Returns:
            Optional[Dict[str, str]]: Extracted data containing "invoice_number", "total_amount", and "receiver".
        """
        if self.method == "standard":
            return self.extract_standard(pdf_path)
        elif self.method == "ocr":
            return self.extract_ocr(pdf_path)
        elif self.method == "gpt":
            return self.extract_gpt(pdf_path)
        else:
            logging.error(f"Unknown extraction method: {self.method}")
            return None

    def extract_standard(self, pdf_path: str) -> Optional[Dict[str, str]]:
        """
        Extract invoice data using standard Python libraries.

        Args:
            pdf_path (str): Path to the PDF file.

        Returns:
            Optional[Dict[str, str]]: Extracted data containing "invoice_number", "total_amount", and "receiver".
        """
        logging.debug(f"Extracting data using standard method: {pdf_path}")
        try:
            reader = PdfReader(pdf_path)

            # Extract text from all pages
            text = "\n".join(
                page.extract_text() for page in reader.pages if page.extract_text()
            )
            # extract spaces, keep newlines, all to lowercase
            text = re.sub(r" ", "", text).lower()
            logging.debug(f"Extracted text: {text}")

            # extract invoice number, total amount, and receiver
            invoice_number = re.search(r"Invoice Number: (\d+)", text)

            # Adjusted regex pattern (optional words, optional colon, optional spaces, optional currency symbol, thousand numbers and comma seperated)
            total_amount_pattern = r"(?:totaal|totaalbedrag|factuurbedrag)(?:\:?)\s*€?\s*([\d,]+\.\d{2}|[\d,]+,\d{2})"
            total_amount = re.search(total_amount_pattern, text)
            logging.debug(f"Extracted total amount: {total_amount}")

            # Extract IBAN
            iban_pattern = r"[a-z]{2}\d{2}[a-z0-9]{4}\d{10}"
            receiver = re.search(iban_pattern, text)

            return {
                "invoice_number": invoice_number.group(1) if invoice_number else None,
                "total_amount": (
                    float(
                        total_amount.group(1)
                        .replace(",", ".")
                        .replace(" ", "")
                        .replace("€","")
                    )
                    if total_amount
                    else None
                ),
                "receiver": receiver.group() if receiver else None,
            }
        except Exception as e:
            logging.error(f"Error reading PDF {pdf_path}: {e}")
            return None

    def extract_ocr(self, pdf_path: str) -> Optional[Dict[str, str]]:
        """
        Extract invoice data using OCR.

        Args:
            pdf_path (str): Path to the PDF file.

        Returns:
            Optional[Dict[str, str]]: Extracted data containing "invoice_number", "total_amount", and "receiver".
        """
        logging.info(f"Extracting data using OCR: {pdf_path}")
        try:
            images = convert_from_path(pdf_path)
            text = " ".join(pytesseract.image_to_string(image) for image in images)
            invoice_number = re.search(r"Invoice Number[:\s]+(\d+)", text)
            total_amount = re.search(
                r"(?:Total Amount|Totaal|Totaalbedrag)[:\s]+([\d,]+\.\d{2})", text
            )
            receiver = re.search(r"(?:Receiver|Ontvanger)[:\s]+([\w\s]+)", text)
            logging.debug(f"Extracted text: {text}")
            return {
                "invoice_number": invoice_number.group(1) if invoice_number else None,
                "total_amount": total_amount.group(1) if total_amount else None,
                "receiver": receiver.group(1) if receiver else None,
                "raw_text": text
            }
        except Exception as e:
            logging.error(f"Error processing PDF {pdf_path}: {e}")
            return None

    def extract_gpt(self, pdf_path: str) -> Optional[Dict[str, str]]:
        """
        Extract invoice data using GPT-4.

        Args:
            pdf_path (str): Path to the PDF file.

        Returns:
            Optional[Dict[str, str]]: Extracted data containing "invoice_number", "total_amount", and "receiver".
        """
        logging.info(f"Extracting data using GPT-4: {pdf_path}")
        try:
            with open(pdf_path, "rb") as pdf_file:
                pdf_content = pdf_file.read()

            response = openai.ChatCompletion.create(
                model="gpt-4-0613",
                messages=[
                    {
                        "role": "system",
                        "content": "You are an assistant that extracts structured data from invoice PDFs.",
                    },
                    {
                        "role": "user",
                        "content": f"Extract the invoice number, total amount, and receiver from this PDF: {pdf_content}. Return all extracted text, end with the asked metrics.",
                    },
                ],
            )

            text = response.choices[0].message.content
            invoice_number = re.search(r"Invoice Number[:\s]+(\d+)", text)
            total_amount = re.search(
                r"(?:Total Amount|Totaal|Totaalbedrag)[:\s]+([\d,]+\.\d{2})",
                text,
            )
            receiver = re.search(
                r"(?:Receiver|Ontvanger)[:\s]+([\w\s]+)", text
            )
            logging.debug(f"Extracted text: {text}")
            return {
                "invoice_number": invoice_number.group(1) if invoice_number else None,
                "total_amount": total_amount.group(1) if total_amount else None,
                "receiver": receiver.group(1) if receiver else None,
                "text": text
            }
        except Exception as e:
            logging.error(f"Error processing PDF {pdf_path}: {e}")
            return None

    def extract_all_data(self):
        """
        Extract data from all PDFs in the directory and save them in a list.
        """
        logging.info("Extracting data from all PDFs...")
        for pdf_file in os.listdir(self.pdf_dir):
            pdf_path = os.path.join(self.pdf_dir, pdf_file)
            invoice_data = self.extract_invoice_data(pdf_path)
            if invoice_data:
                invoice_data["file_path"] = pdf_path
                self.extracted_data.append(invoice_data)

    def compare_and_update(self):
        """
        Compare extracted data with Excel file and update rows accordingly.
        Filters the Excel file for matches on receiver, IBAN, and total amount,
        logs the number of records per step, and processes hits.
        """
        logging.info("Comparing extracted data with Excel file...")
        logging.debug(f"Extracted data: {self.extracted_data}")

        # Filter rows where "Factuur op SP" is TBD
        tbd_rows = self.data[self.data["Factuur op SP"].str.lower() == "tbd"]
        logging.info(f"Initial rows to process (Factuur op SP = TBD): {len(tbd_rows)}")

        # Iterate over extracted data from every invoice to filter Excel on potential matches
        for invoice_data in self.extracted_data:
            # Initialize lists to store hits for each filter
            receiver_hits, amount_hits, iban_hits = [], [], []
            
            # Filter on receiver
            receiver_match = tbd_rows[tbd_rows["Ontvanger"] == invoice_data["receiver"]]
            logging.info(f"Receiver hits for {invoice_data['receiver']}: {len(receiver_match)}")
            receiver_hits.append(receiver_match)

            # Filter on total amount
            amount_match = tbd_rows[tbd_rows["Totaal Bedrag"].astype(str) == str(invoice_data["total_amount"])]
            logging.info(f"Total amount hits for {invoice_data['total_amount']}: {len(amount_match)}")
            amount_hits.append(amount_match)

            # Filter on IBAN
            if "iban" in invoice_data and invoice_data["iban"]:
                iban_match = tbd_rows[tbd_rows["Betaling_Rekening"] == invoice_data["iban"]]
                logging.info(f"IBAN hits for {invoice_data['iban']}: {len(iban_match)}")
                iban_hits.append(iban_match)

            # Combine all hits into one DataFrame
            combined_hits = pd.concat(
                receiver_hits + amount_hits + iban_hits
            ).drop_duplicates()
            logging.info(f"Total unique hits after combining filters: {len(combined_hits)}")

            # Process hits, if single match and update, if multiple write to a separate file, else warning
            if len(combined_hits) == 1:
                logging.info("!!!!!!!!! Single match found for the extracted data.")
                # row = combined_hits.iloc[0]

                # new_invoice_number = (
                #     max(self.data["Factuur_nummer"].dropna().astype(int), default=0) + 1
                # )
                # self.data.loc[row.name, "Factuur_nummer"] = new_invoice_number
                # logging.info(
                #     f"Matched and updated: {row['Factuur_nummer']} -> {new_invoice_number}"
                # )

                # new_pdf_name = f"Invoice_{new_invoice_number}.pdf"
                # os.rename(
                #     invoice_data["file_path"],
                #     os.path.join(self.output_pdf_dir, new_pdf_name),
                # )
                # logging.info(
                #     f"Renamed and saved: {invoice_data['file_path']} -> {new_pdf_name}"
                # )
            # Multiple matches found: log a warning
            elif len(combined_hits) > 1:
                # Add file name to the combined hits & store in all_combined_hits
                combined_hits["Bestandsnaam"] = invoice_data["file_path"]
                self.all_combined_hits = pd.concat(
                    [combined_hits, self.all_combined_hits]
                ).drop_duplicates()
                logging.warning(
                    f"Multiple matches found for the extracted data: {len(combined_hits)}"
                )
                logging.debug(combined_hits)
            else:
                logging.warning("No matches found for the extracted data.")

    def save_results(self):
        """
        Save the updated Excel file.
        """
        logging.info("Saving updated Excel file...")
        self.data.to_excel(OUTPUT_EXCEL, index=False)
        self.all_combined_hits.to_excel("combined_hits.xlsx", index=False)


if __name__ == "__main__":
    matcher = InvoiceMatcher(INPUT_EXCEL, PDF_DIR, OUTPUT_PDF_DIR, EXTRACTION_METHOD)
    matcher.extract_all_data()
    matcher.compare_and_update()
    matcher.save_results()
    logging.info("Invoice matching completed.")
