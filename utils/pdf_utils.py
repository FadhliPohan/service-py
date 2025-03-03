from PyPDF2 import PdfReader, PdfWriter
import os

def split_pdf(file_path, start_page, end_page):
    if start_page < 1 or end_page < 1:
        raise ValueError("Start and end page numbers must be greater than 0")

    reader = PdfReader(file_path)
    total_pages = len(reader.pages)
    if start_page > total_pages or end_page > total_pages:
        raise ValueError("Start or end page number exceeds total pages in the document")

    writer = PdfWriter()
    for page_num in range(start_page - 1, end_page):
        writer.add_page(reader.pages[page_num])

    split_file_path = f"split_{os.path.basename(file_path)}"
    with open(split_file_path, "wb") as output_pdf:
        writer.write(output_pdf)

    return split_file_path
