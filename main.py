import os
import re
import logging
from pathlib import Path
from typing import Dict, Optional, List
import pandas as pd
import base64

import yaml
import json
from dotenv import load_dotenv
import pytesseract
from PIL import Image
from pdf2image import convert_from_path
from PyPDF2 import PdfReader
from openai import OpenAI


# Setup logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

# load .env file if present
try:
    load_dotenv(".env")
    logging.info("Loaded environment variables from .env file.")
except ImportError:
    logging.warning("dotenv module not found, skipping environment variable loading.")

# Load settings from YAML
with open("settings.yaml", "r") as f:
    settings = yaml.safe_load(f)

INPUT_PATH = Path(settings.get("input_path", "input"))
OUTPUT_PATH = Path(settings.get("output_path", "output"))
INPUT_EXCEL = INPUT_PATH / settings["input_excel"]
OUTPUT_EXCEL = OUTPUT_PATH / settings["output_excel"]
PDF_DIR = INPUT_PATH / settings["pdf_directory"]
OUTPUT_PDF_DIR = OUTPUT_PATH / settings["output_pdf_directory"]
EXTRACTION_METHOD = settings["extraction_method"]  # Options: "standard", "ocr", "gpt"
TESSERACT_PATH = settings.get("tesseract_path")  # Optional
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")  # Optional, set in .env file


# Configure Tesseract executable path if required
if TESSERACT_PATH:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

# Configure OpenAI API key if required
if OPENAI_API_KEY:
    openai_client = OpenAI(api_key=OPENAI_API_KEY)


class InvoiceMatcher:

    def __init__(
        self, input_file: Path, pdf_dir: Path, output_pdf_dir: Path, method: str
    ):
        self.input_file = input_file
        self.pdf_dir = pdf_dir
        self.output_pdf_dir = output_pdf_dir
        self.method = method
        self.data = self.load_input()
        self.extracted_data = []
        self.matched_rows = []
        self.uncertain_matches = []
        self.unmatched_invoices = []

    def load_input(self) -> pd.DataFrame:
        logging.info(f"Loading Excel file from {self.input_file}")
        df = pd.read_csv(self.input_file, sep=";", encoding="utf-8", decimal=",")
        df.columns = df.columns.str.strip().str.lower()
        logging.debug(f"Loaded columns: {df.columns.tolist()}")
        return df

    def extract_invoice_data(self, pdf_path: Path) -> Optional[Dict[str, str]]:
        if self.method == "standard":
            return self.extract_standard(pdf_path)
        elif self.method == "ocr":
            return self.extract_ocr(pdf_path)
        elif self.method == "gpt":
            return self.extract_gpt(pdf_path)
        elif self.method == "ocr_gpt":
            return self.extract_ocr_gpt(pdf_path)
        else:
            logging.error(f"Unknown extraction method: {self.method}")
            return None

    def extract_standard(self, pdf_path: Path) -> Optional[Dict[str, str]]:
        logging.info(f"Extracting (standard) from: {pdf_path}")
        try:
            reader = PdfReader(str(pdf_path))
            text = "\n".join(
                page.extract_text() for page in reader.pages if page.extract_text()
            )
            text = re.sub(r" ", "", text).lower()
            invoice_number = re.search(r"invoice\s*number[:\s]*(\d+)", text)
            total_amount = re.search(
                r"(?:totaal|totaalbedrag|factuurbedrag)[:\s]*€?\s*([\d,.]+)", text
            )
            receiver = re.search(r"[a-z]{2}\d{2}[a-z0-9]{4}\d{10}", text)
            return {
                "invoice_number": invoice_number.group(1) if invoice_number else None,
                "total_amount": (
                    float(total_amount.group(1).replace(",", "."))
                    if total_amount
                    else None
                ),
                "receiver": receiver.group() if receiver else None,
                "file_path": str(pdf_path),
            }
        except Exception as e:
            logging.error(f"Error during standard extraction: {e}")
            return None

    def extract_ocr(self, pdf_path: Path) -> Optional[Dict[str, str]]:
        logging.debug(f"Extracting (OCR) from: {pdf_path}")
        pytesseract.pytesseract.tesseract_cmd = "/opt/homebrew/bin/tesseract"
        logging.info(
            f"Using Tesseract from: {pytesseract.pytesseract.tesseract_cmd}, version: {pytesseract.get_tesseract_version()}"
        )

        try:
            images = convert_from_path(str(pdf_path))
            text = "".join(pytesseract.image_to_string(image) for image in images)
            text = re.sub(r" ", "", text).lower()
            invoice_number = re.search(r"invoice\s*number[:\s]*(\d+)", text)
            total_amount = re.search(
                r"(?:totaal|totaalbedrag|factuurbedrag)[:\s]*€?\s*([\d,.]+)", text
            )
            receiver = re.search(r"[a-z]{2}\d{2}[a-z0-9]{4}\d{10}", text)
            return {
                "invoice_number": invoice_number.group(1) if invoice_number else None,
                "total_amount": (
                    float(total_amount.group(1).replace(",", "."))
                    if total_amount
                    else None
                ),
                "receiver": receiver.group() if receiver else None,
                "file_path": str(pdf_path),
            }
        except Exception as e:
            logging.error(f"Error during OCR extraction: {e}")
            return None

    def extract_gpt(self, pdf_path: Path) -> Optional[Dict[str, str]]:
        logging.info(f"Extracting (GPT) from: {pdf_path}")

        try:
            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()

            base64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")

            prompt = """Lees de inhoud van de PDF en extraheer de volgende gegevens:
                        - Factuurnummer (bijv. factuurnummer, invoice number)
                        - Totaal bedrag (vaak onderaan, bijv. totaal, totaal bedrag, total amount, factuurbedrag)
                        - Ontvanger (IBAN), die volgt patroon [a-z]{2}\d{2}[a-z0-9]{4}\d{10}

                        Geef de resultaten terug in JSON-formaat met de volgende structuur:
                        1. invoice_number
                        2. total_amount
                        3. receiver
                        Bij missende waarden, gebruik null

                        Getallen kunnen komma's bevatten, dus converteer ze naar een float.
                    """

            response = openai_client.chat.completions.create(
                model="gpt-4.1-mini",
                messages=[
                    {
                        "role": "system",
                        "content": "Jij bent een expert in het extraheren van factuurgegevens uit PDF's.",
                    },
                    {
                        "role": "user",
                        "content": f"""{prompt} \n
                        (PDF contents encoded in base64 below) \n
                        {base64_pdf}""",
                    },
                ],
            )

            # Print usage
            usage = response.usage
            logging.info(
                f"Prompt tokens: {usage.prompt_tokens}, Completion tokens: {usage.completion_tokens}, Total: {usage.total_tokens}"
            )

            result = response.choices[0].message.content
            data = json.loads(result)
            logging.info(f"Extracted data: {data}")

            return {
                "invoice_number": data.get("invoice_number"),
                "total_amount": (
                    float(data.get("total_amount"))
                    if data.get("total_amount")
                    else None
                ),
                "receiver": data.get("receiver"),
                "file_path": str(pdf_path),
            }
        except Exception as e:
            logging.error(f"Error during GPT extraction: {e}")
            return None

    def extract_ocr_gpt(self, pdf_path: Path) -> Optional[Dict[str, str]]:
        logging.info(f"Extracting (OCR + GPT) from: {pdf_path}")
        pytesseract.pytesseract.tesseract_cmd = "/opt/homebrew/bin/tesseract"
        logging.debug(
            f"Using Tesseract from: {pytesseract.pytesseract.tesseract_cmd}, version: {pytesseract.get_tesseract_version()}"
        )

        try:
            # First, use OCR to extract text
            images = convert_from_path(str(pdf_path))
            text = "".join(pytesseract.image_to_string(image) for image in images)
            text = re.sub(r" ", "", text).lower()

            # Then, use GPT to extract structured data
            prompt = """Lees de inhoud van de PDF en extraheer de volgende gegevens:
                        - Factuurnummer (bijv. factuurnummer, invoice number)
                        - Totaal bedrag (vaak onderaan, bijv. totaal, totaal bedrag, total amount, factuurbedrag)
                        - Ontvanger (IBAN), die volgt patroon [a-z]{2}\d{2}[a-z0-9]{4}\d{10}

                        Geef de resultaten terug in JSON-formaat met de volgende structuur:
                        1. invoice_number (Factuurnummer)
                        2. total_amount (Totaal bedrag)
                        3. receiver (IBAN)
                        Bij missende waarden, gebruik null

                        Getallen kunnen komma's bevatten, dus converteer ze naar een float.
                    """

            response = openai_client.chat.completions.create(
                model="gpt-4.1-mini",
                messages=[
                    {
                        "role": "system",
                        "content": """
                            Jij bent een expert in het extraheren van factuurgegevens uit PDF's. 
                            Je bent erg secuur en neemt de tijd om een goede en gevalideerder JSON als output te geven.
                            Je reageert in plain text, zonder markdown of andere opmaak. Dus geen ```json of dergelijke.
                            Neem de tijd om de PDF goed te lezen en de juiste gegevens te extraheren.
                            Kijk goed of je een IBAN kunt vinden, en of het totaal bedrag een getal is.
                            """,
                    },
                    {
                        "role": "user",
                        "content": f"""{prompt} \n
                        (PDF contents extracted via OCR) \n
                        {text}""",
                    },
                ],
            )
        except Exception as e:
            logging.error(f"Error during OCR + GPT extraction: {e}")
            return None

        # Print usage
        usage = response.usage
        logging.info(
            f"Prompt tokens: {usage.prompt_tokens}, Completion tokens: {usage.completion_tokens}, Total: {usage.total_tokens}"
        )

        result = response.choices[0].message.content
        logging.info(result)
        data = json.loads(result)
        logging.info(f"Extracted data: {data}")

        return {
            "invoice_number": data.get("invoice_number"),
            "total_amount": (
                float(data.get("total_amount")) if data.get("total_amount") else None
            ),
            "receiver": data.get("receiver"),
            "file_path": str(pdf_path),
        }

    def extract_all_data(self):
        logging.info("Extracting data from all PDFs...")
        for pdf_file in self.pdf_dir.iterdir():
            if pdf_file.suffix.lower() == ".pdf":
                invoice_data = self.extract_invoice_data(pdf_file)
                if invoice_data:
                    self.extracted_data.append(invoice_data)

    def compare_and_update(self):
        logging.info("Comparing extracted invoices with Excel data...")
        # receiver_col = "ontvanger"
        iban_col = "iban betaald"
        amount_col = "totaal bedrag"

        if iban_col not in self.data.columns or amount_col not in self.data.columns:
            logging.error(
                "Required columns 'ontvanger' or 'totaal bedrag' not found in Excel."
            )
            return

        for invoice in self.extracted_data:
            iban = invoice.get("receiver")
            amount = invoice.get("total_amount")
            matches = self.data[
                (self.data[iban_col] == iban)
                & (self.data[amount_col].astype(float) == amount)
            ]
            if len(matches) == 1:
                logging.info(f"Single match found for invoice {invoice['file_path']}")
                matched_row = matches.copy()
                matched_row["source_file"] = invoice["file_path"]
                self.matched_rows.append(matched_row)
            elif len(matches) > 1:
                logging.warning(f"Multiple matches for invoice {invoice['file_path']}")
                uncertain = matches.copy()
                uncertain["source_file"] = invoice["file_path"]
                self.uncertain_matches.append(uncertain)
            else:
                logging.warning(f"No matches for invoice {invoice['file_path']}")
                logging.warning(
                    f"Invoice {invoice['file_path']} with receiver {iban} and amount {amount} not found in Excel."
                )
                self.unmatched_invoices.append(invoice)

    def save_results(self):
        logging.info("Saving results to Excel...")
        with pd.ExcelWriter(OUTPUT_EXCEL, engine="xlsxwriter") as writer:
            if self.matched_rows:
                pd.concat(self.matched_rows).to_excel(
                    writer, sheet_name="Matches", index=False
                )
            if self.uncertain_matches:
                pd.concat(self.uncertain_matches).to_excel(
                    writer, sheet_name="Uncertain", index=False
                )
            if self.unmatched_invoices:
                pd.DataFrame(self.unmatched_invoices).to_excel(
                    writer, sheet_name="No Match", index=False
                )


if __name__ == "__main__":
    matcher = InvoiceMatcher(INPUT_EXCEL, PDF_DIR, OUTPUT_PDF_DIR, EXTRACTION_METHOD)
    matcher.extract_all_data()
    matcher.compare_and_update()
    matcher.save_results()
    logging.info("Invoice matching completed.")
