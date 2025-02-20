# Invoice Matcher Setup and Execution

This guide explains how to set up and execute the invoice matcher script. Follow these steps to ensure all dependencies are installed, and the environment is ready.

## Prerequisites

1. **Python**: Ensure Python 3.9 or higher is installed.
2. **UV (Unvirtual)**: Ensure `uv` is installed. Install it via pip if necessary:

```bash
pip install uv
```

---

## Step 1: Activating the UV Environment

When a new developer downloads the repository, they simply need to activate the `uv` environment:

```bash
uv activate
```

This will automatically set up and activate the environment with all required dependencies.

---

## Step 2: Configure Settings

Create a `settings.json` file in the root directory with the following structure:

```json
{
    "input_excel": "path_to_input_excel.xlsx",
    "output_excel": "path_to_output_excel.xlsx",
    "pdf_directory": "path_to_pdf_directory",
    "output_pdf_directory": "path_to_output_pdf_directory",
    "extraction_method": "standard",  
    "tesseract_path": "path_to_tesseract_executable",  
    "openai_api_key": "your_openai_api_key"
}
```

- Replace `extraction_method` with one of the following: `standard`, `ocr`, or `gpt`.
- Ensure paths are valid and accessible.
- Add the OpenAI API key if using GPT-based extraction.

---

## Step 3: Run the Script

Activate the environment using `uv activate` if not already activated, then run the script:

```bash
python combined_invoice_matcher.py
```

---

## Additional Notes

1. **PDF Dependencies**: The `pdf2image` package requires the `poppler-utils` library.
    - **Linux**: `sudo apt install poppler-utils`
    - **Mac**: `brew install poppler`
    - **Windows**: Download [Poppler for Windows](http://blog.alivate.com.au/poppler-windows/).

2. **Tesseract Installation**: Ensure `tesseract_path` is correctly set in `settings.json` if Tesseract is installed outside the default locations.

3. **Permissions**: Ensure the script has access to read/write in the specified directories.

---

## Troubleshooting

- **Missing Packages**: Run `uv install` to ensure all dependencies are installed.
- **Tesseract Not Found**: Verify the `tesseract_path` is correctly set in `settings.json`.
- **OpenAI Errors**: Ensure your API key is valid and your OpenAI usage quota is sufficient.
