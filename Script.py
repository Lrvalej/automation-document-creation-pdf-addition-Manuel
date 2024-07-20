import pandas as pd
from docxtpl import DocxTemplate
import os
from subprocess import run
import fitz  # PyMuPDF
import pyodbc

# Database connection parameters
server = 'your_server'  # Replace with your server name
database = 'your_database'  # Replace with your database name

# Create a connection to the SQL Server using Windows Authentication
conn = pyodbc.connect(f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes;')

# Run the query and load data into a pandas dataframe
query = 'Select * from [table]'
df = pd.read_sql(query, conn)

# Template file paths
template1_path = 'template1.docx'  # Template for rows with condition == 0
template2_path = 'template2.docx'  # Template for rows with condition == 1

# Output directory
output_dir = 'output'
os.makedirs(output_dir, exist_ok=True)

# Iterate through each row in the dataframe
for idx, row in df.iterrows():
    # Select the appropriate template based on the "condition" column
    if row['condition'] == 1:
        doc = DocxTemplate(template2_path)
    else:
        doc = DocxTemplate(template1_path)
    
    # Convert row to dictionary and render it in the template
    context = row.to_dict()
    doc.render(context)
    
    # Create the document name using "ColumnA" and "ColumnB" columns
    doc_name = f"{row['ColumnA']}_{row['ColumnB']}"
    
    # Save the document
    output_docx_path = os.path.join(output_dir, f'{doc_name}.docx')
    doc.save(output_docx_path)
    
    # Convert to PDF using IronPDF
    doc_pdf_path = os.path.join(output_dir, f'{doc_name}_doc.pdf')
    run(['ironpdf', 'html-to-pdf', '--input-file', output_docx_path, '--output-file', doc_pdf_path])
    
    # Assuming the existing PDF is also converted using IronPDF or already present
    existing_pdf_path = os.path.join(output_dir, f'{doc_name}.pdf')  # Adjust as necessary
    run(['ironpdf', 'html-to-pdf', '--input-file', output_docx_path, '--output-file', existing_pdf_path])
    
    # Create a combined PDF
    combined_pdf_path = os.path.join(output_dir, f'{doc_name}_combined.pdf')
    doc_pdf = fitz.open(doc_pdf_path)
    existing_pdf = fitz.open(existing_pdf_path)
    combined_pdf = fitz.open()

    # Append pages of the two PDFs
    combined_pdf.insert_pdf(doc_pdf)
    combined_pdf.insert_pdf(existing_pdf)

    # Save the combined PDF
    combined_pdf.save(combined_pdf_path)
    combined_pdf.close()
    doc_pdf.close()
    existing_pdf.close()
    
    # Print the combined PDF
    run(['lp', combined_pdf_path])
