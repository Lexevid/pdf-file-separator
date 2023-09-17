import os
import PyPDF2
from openpyxl import load_workbook


def extract_pages_from_pdf(pdf_file, excel_file):
    workbook = load_workbook(excel_file)
    sheet = workbook.active

    names = [cell.value for cell in sheet['A'][1:]]

    with open(pdf_file, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        num_pages = len(reader.pages)

        for i, name in enumerate(names):
            writer = PyPDF2.PdfWriter()
            writer.add_page(reader.pages[i])

          # New file name
            new_filename = f"{name}.pdf"
            with open(new_filename, 'wb') as new_file:
                writer.write(new_file)

            print(f"Page {i + 1} saved as {new_filename}")

    print("Extraction completed.")

# Usage example
# Type the directory of your PDF file into pdf_file
# type the directory of your excel file into excel_file
pdf_file = "L:/Tests/pdf_file.pdf"
excel_file = "L:/Tests/Names.xlsx"
extract_pages_from_pdf(pdf_file, excel_file)
