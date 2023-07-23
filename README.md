# PDF to Excel Converter

This Python script extracts information from a PDF file and saves it into an Excel file. It can be used for processing conference papers or similar documents that have specific formatting for authors, affiliations, session names, titles, and abstracts.

## Requirements

- Python 3.x
- `fitz` library (PyMuPDF)
- `openpyxl` library

You can install the required libraries using `pip`:
 
`pip install PyMuPDF openpyxl`


## Usage

1. Ensure you have the required Python libraries installed as mentioned in the Requirements section.

2. Place the PDF file (`book.pdf` in this example) you want to extract information from in the same directory as this script.

3. Open a terminal or command prompt and navigate to the directory containing the script and the PDF file.

4. Run the script with the following command:

`python main.py`


5. The script will process the PDF starting from a specified page (provided in the script as `start_page = 44`). It will extract information from the PDF and save it into an Excel file (`result.xlsx` in this example) in the same directory.

6. The Excel file will have the following columns:

   - Name (incl. titles if any mentioned)
   - Affiliations
   - Session name
   - Persons Location (not implemented)
   - Topic Title
   - Presentation Abstract

7. After running the script, you can find the extracted information in the `result.xlsx` file.

## Important Notes

- The script assumes specific formatting in the PDF file, so ensure the PDF follows the expected structure for accurate extraction.

- If you encounter any issues or errors, make sure the PDF file and its structure are correct, and consider adjusting the `start_page` value if needed.

