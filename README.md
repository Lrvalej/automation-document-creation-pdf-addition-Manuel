# automation-document-creation-pdf-addition-Manuel

# Document Automation and Printing

This project automates the creation, merging, and printing of Word and PDF documents using data retrieved from a SQL Server database. It uses the following libraries: `pandas`, `docxtpl`, `pyodbc`, `fitz` (PyMuPDF), and `ironpdf`.

## Features

- Connects to a SQL Server database using Windows Authentication.
- Queries data from the a database table.
- Autofills Word document templates based on the query results.
- Converts Word documents to PDFs.
- Merges generated PDFs with existing PDFs.
- Saves the combined PDFs.
- Prints the combined PDFs to the default printer.

## Requirements

- Python 3.x
- `pandas`
- `docxtpl`
- `pyodbc`
- `fitz` (PyMuPDF)
- `ironpdf`

## Installation

1. Install the required Python libraries:
    ```sh
    pip install pandas docxtpl pyodbc pymupdf ironpdf-python
    ```

## Configuration

1. Update the database connection parameters in the script:
    ```python
    server = 'your_server'  # Replace with your server name
    database = 'your_database'  # Replace with your database name
    ```

2. Ensure you have the necessary templates and data:
    - `template1.docx`: Template for rows with `Condition` equal to 0.
    - `template2.docx`: Template for rows with `Condition` equal to 1.
    - Output directory: Ensure the `output` directory exists or the script can create it.

## Usage

1. Run the script:
    ```sh
    python document_automation.py
    ```

2. The script will:
    - Connect to the SQL Server database.
    - Run the query.
    - Generate Word documents based on the query results.
    - Convert Word documents to PDFs.
    - Merge the generated PDFs with existing PDFs.
    - Save the combined PDFs.
    - Print the combined PDFs to the default printer.
