import os
import subprocess
from docx import Document
from pdfkit import from_file

def convert_to_pdf(docx_path, start, end):
    # Buat file PDF sementara
    pdf_path = os.path.splitext(docx_path)[0] + '.pdf'
    
    # Konversi DOCX ke PDF menggunakan pdfkit
    try:
        from_file(docx_path, pdf_path)
    except Exception as e:
        raise Exception(f"Gagal mengkonversi DOCX ke PDF: {str(e)}")
    
    return pdf_path