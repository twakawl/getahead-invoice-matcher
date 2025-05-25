# Invoice Matcher Setup and Execution

This guide explains how to set up and execute the invoice matcher script. Follow these steps to ensure all dependencies are installed, and the environment is ready.

## Prerequisites

1. **Python**: Ensure Python 3.9 or higher is installed.
2. **UV (Unvirtual)**: Ensure `uv` is installed. Install it via pip if necessary:

```bash
pip install uv
```

3. **Tesseract OCR**: Required for OCR-based extraction.
    - **macOS**: `brew install tesseract`
    - **Linux**: `sudo apt install tesseract-ocr`
    - **Windows**: Install from [UB Mannheim builds](https://github.com/UB-Mannheim/tesseract/wiki)

4. **Poppler**: Required by `pdf2image`.
    - **macOS**: `brew install poppler`
    - **Linux**: `sudo apt install poppler-utils`
    - **Windows**: [Poppler for Windows](http://blog.alivate.com.au/poppler-windows/)

5. **OpenAI API Key**: Required if using GPT-based extraction.

---

## Step 1: Activating the UV Environment

When a new developer downloads the repository, they simply need to activate the `uv` environment:

```bash
uv activate
```

This will automatically set up and activate the environment with all required dependencies.

---

## Step 2: Configure Environment

Create a `.env` file in the root directory with the following:

```env
OPENAI_API_KEY=your_openai_api_key
```

Create a `settings.yaml` file in the root directory:

```yaml
input_path: "input"
output_path: "output"
input_excel: "betalingen_processed.csv"
output_excel: "output.xlsx"
pdf_directory: "invoices"
output_pdf_directory: "invoices"
extraction_method: "standard"  # Options: standard, ocr, gpt
tesseract_path: "/opt/homebrew/bin/tesseract"  # Optional on macOS if not auto-detected
```

---

## Step 3: Run the Script

Activate the environment using `uv activate` if not already activated, then run the script:

```bash
python main.py
```

---

## Additional Notes

1. **Token Estimation**: GPT-based extraction prints estimated token usage. High token counts may result in API errors.
2. **GPT Cost Logging**: Token usage is logged and approximate USD cost is calculated per request.
3. **OCR Path Override**: If your environment doesn't inherit `PATH`, the script explicitly sets Tesseract to `/opt/homebrew/bin/tesseract`.
4. **Fallback Handling**: You can configure the script to try GPT → Standard → OCR if needed.

---

## Troubleshooting

- **Missing Packages**: Run `uv install` to ensure all dependencies are installed.
- **Tesseract Not Found**: Ensure the correct path is set in the `.env` or within the script.
- **Poppler Missing**: If you see `Unable to get page count`, install `poppler`.
- **GPT Token Limits**: If you see HTTP 429 or "Request too large" errors, reduce PDF size or switch to text-based extraction.
- **Rate Limits**: Monitor your [OpenAI usage dashboard](https://platform.openai.com/account/usage).

---

## Future Improvements

- Upload files to OpenAI using Assistant API
- Add retry logic and exponential backoff for GPT rate limits
- Use PDF text extraction for GPT input to reduce token usage
- Extend matching with fuzzy logic or NLP

---

Maintained by: Twan Houwers
