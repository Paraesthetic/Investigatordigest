import os
import subprocess
import sys
import shutil
import pandas as pd
from docx2pdf import convert
from openpyxl import load_workbook
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import PyPDF2
from easygui import diropenbox

# Function to install dependencies
def install(package):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
    except Exception as e:
        print(f"Failed to install {package}: {e}")

# Check and install dependencies
def check_and_install_dependencies():
    dependencies = {
        "pandas": "pandas",
        "openpyxl": "openpyxl",
        "reportlab": "reportlab",
        "docx2pdf": "docx2pdf",
        "PyPDF2": "PyPDF2",
        "eml_parser": "eml_parser",
        "extract_msg": "extract_msg",
        "easygui": "easygui"
    }
    
    for package_name, import_name in dependencies.items():
        try:
            __import__(import_name)
            print(f"{package_name} is already installed.")
        except ImportError:
            print(f"{package_name} is missing. Installing it now.")
            install(package_name)

check_and_install_dependencies()

# Function to convert .eml files to PDF
def convert_eml_to_pdf(eml_file, output_pdf_path):
    try:
        import eml_parser
        with open(eml_file, 'rb') as f:
            raw_email = f.read()
        ep = eml_parser.eml_parser.EmlParser(include_attachment_data=True)
        parsed_email = ep.decode_email_bytes(raw_email)

        # Extracting email body content
        email_body = parsed_email['body'][0]['content'] if parsed_email['body'] else "No content"
        email_subject = parsed_email.get('header', {}).get('subject', 'No Subject')

        # Create PDF with the email content
        pdf = canvas.Canvas(output_pdf_path, pagesize=letter)
        width, height = letter
        pdf.drawString(40, height - 40, f"Subject: {email_subject}")
        pdf.drawString(40, height - 60, "Body:")
        pdf.drawString(40, height - 80, email_body[:1000])  # Limit the body to avoid issues with long text
        pdf.save()

        print(f"Converted .eml to PDF: {eml_file}")
    except Exception as e:
        print(f"Error converting .eml to PDF: {e}")

# Function to convert .msg files to PDF
def convert_msg_to_pdf(msg_file, output_pdf_path):
    try:
        import extract_msg
        msg = extract_msg.Message(msg_file)
        msg_subject = msg.subject or "No Subject"
        msg_body = msg.body or "No Body"

        # Create PDF with the email content
        pdf = canvas.Canvas(output_pdf_path, pagesize=letter)
        width, height = letter
        pdf.drawString(40, height - 40, f"Subject: {msg_subject}")
        pdf.drawString(40, height - 60, "Body:")
        pdf.drawString(40, height - 80, msg_body[:1000])  # Limit the body to avoid issues with long text
        pdf.save()

        print(f"Converted .msg to PDF: {msg_file}")
    except Exception as e:
        print(f"Error converting .msg to PDF: {e}")

# Function to convert Excel to PDF using pandas and reportlab
def convert_excel_to_pdf(excel_file, output_pdf_path):
    try:
        df = pd.read_excel(excel_file)

        # Set up the ReportLab PDF canvas
        pdf = canvas.Canvas(output_pdf_path, pagesize=letter)
        width, height = letter

        # Write each row of the Excel file to the PDF
        text = pdf.beginText(40, height - 40)
        for column in df.columns:
            text.textLine(f"{column}: {df[column].values}")
        
        pdf.drawText(text)
        pdf.save()

        print(f"Converted Excel to PDF: {excel_file}")
    except Exception as e:
        print(f"Error converting Excel to PDF: {e}")

# Function to add a numbered page to PDF
def add_page_number(input_pdf, output_pdf, page_number):
    try:
        pdf_reader = PyPDF2.PdfReader(input_pdf)
        pdf_writer = PyPDF2.PdfWriter()
        
        for page_num, page in enumerate(pdf_reader.pages):
            pdf_writer.add_page(page)

        with open(output_pdf, 'wb') as out:
            pdf_writer.write(out)

        print(f"Successfully added page number to {input_pdf}")
        
    except FileNotFoundError:
        print(f"Error: File not found - {input_pdf}")
    except PyPDF2.errors.PdfReadError:
        print(f"Error: Unable to read PDF - {input_pdf}")
    except Exception as e:
        print(f"An unexpected error occurred while processing {input_pdf}: {e}")

# Function to convert and maintain folder structure
def convert_files_to_pdf(input_folder, output_folder):
    file_count = 1

    for root, dirs, files in os.walk(input_folder):
        # Replicate the subdirectory structure in the output folder
        relative_path = os.path.relpath(root, input_folder)
        target_dir = os.path.join(output_folder, relative_path)
        os.makedirs(target_dir, exist_ok=True)

        for file in files:
            file_path = os.path.join(root, file)
            file_name, file_extension = os.path.splitext(file)

            # Append the old extension to the new PDF (unless it's already a PDF)
            if file_extension.lower() == '.pdf':
                output_pdf_name = f"{file_count}. {file_name}.pdf"
            else:
                output_pdf_name = f"{file_count}. {file_name}{file_extension}.pdf"

            output_pdf_path = os.path.join(target_dir, output_pdf_name)

            try:
                if file.endswith('.docx'):
                    # Convert .docx to PDF
                    convert(file_path, output_pdf_path)
                elif file.endswith('.pdf'):
                    # Simply copy the PDF file and number it
                    shutil.copy(file_path, output_pdf_path)
                elif file.endswith('.eml'):
                    # Convert .eml to PDF
                    convert_eml_to_pdf(file_path, output_pdf_path)
                elif file.endswith('.msg'):
                    # Convert .msg to PDF
                    convert_msg_to_pdf(file_path, output_pdf_path)
                elif file.endswith(('.xlsx', '.xls')):
                    # Convert Excel file to PDF
                    convert_excel_to_pdf(file_path, output_pdf_path)
                else:
                    print(f"Skipping unsupported file type: {file}")
                    continue

                # Add a page number to the PDF
                add_page_number(output_pdf_path, output_pdf_path, file_count)
                file_count += 1

            except Exception as e:
                print(f"Error processing {file_path}: {e}")

# Function to select directories and start conversion
def main():
    input_folder = diropenbox(msg="Select the Input Folder")
    output_folder = diropenbox(msg="Select the Output Folder")
    
    if input_folder and output_folder:
        convert_files_to_pdf(input_folder, output_folder)
        print("Conversion completed.")
    else:
        print("No folder selected.")

if __name__ == "__main__":
    main()
