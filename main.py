from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from utils.pdf_utils import split_pdf
from utils.docx_utils import convert_to_pdf
import os
from datetime import datetime
from docx2pdf import convert
from pdf2docx import Converter

from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import os
import tempfile

from PyPDF2 import PdfReader, PdfWriter

app = FastAPI(
    title="Document Splitter API",
    description="API untuk memecah dokumen PDF, DOCX, dan DOC berdasarkan rentang halaman yang ditentukan.",
    version="1.0.0"
)

# Pastikan folder hasil-split ada
os.makedirs("hasil-split", exist_ok=True)

@app.post("/split-pdf/")
async def split_document(file: UploadFile = File(...), start_page: int = 1, end_page: int = 1):
    # Validasi tipe file
    valid_extensions = [".pdf", ".docx", ".doc"]
    file_extension = os.path.splitext(file.filename)[1].lower()
    if file_extension not in valid_extensions:
        raise HTTPException(status_code=400, detail="Unsupported file type")

    # Simpan file yang diunggah
    file_location = f"temp_{file.filename}"
    with open(file_location, "wb+") as file_object:
        file_object.write(file.file.read())

    try:
        # Memecah file berdasarkan tipe
        if file_extension == ".pdf":
            split_file_location = split_pdf(file_location, start_page, end_page)
        elif file_extension == ".docx":
            split_file_location = split_docx(file_location, start_page, end_page)
        else:
            raise HTTPException(status_code=400, detail="Unsupported file type for splitting")

        # Buat nama file hasil dengan timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        split_file_name = f"hasil-split/{timestamp}_{os.path.basename(file.filename)}"
        
        # Pindahkan file hasil ke folder hasil-split
        os.rename(split_file_location, split_file_name)

        # Mengirimkan file yang telah dipisah
        return FileResponse(path=split_file_name, filename=os.path.basename(split_file_name))

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        # Hapus file sementara
        os.remove(file_location)
        if os.path.exists(split_file_location):
            os.remove(split_file_location)

@app.post("/convert-docx-to-pdf/")
async def convert_docx_to_pdf(file: UploadFile = File(...)):
    # Validasi input
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="File harus berformat DOCX")

    # Simpan file DOCX yang diunggah ke direktori sementara
    temp_dir = tempfile.mkdtemp()
    temp_docx_path = os.path.join(temp_dir, file.filename)
    
    with open(temp_docx_path, "wb") as buffer:
        buffer.write(await file.read())

    # Tentukan path untuk file PDF yang dihasilkan
    pdf_file_path = os.path.join(temp_dir, f"{os.path.splitext(file.filename)[0]}.pdf")

    # Konversi DOCX ke PDF
    try:
        convert(temp_docx_path, pdf_file_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Terjadi kesalahan saat mengonversi file: {str(e)}")

    # Kembalikan file PDF sebagai respons
    return FileResponse(pdf_file_path, media_type='application/pdf', filename=os.path.basename(pdf_file_path))

@app.post("/slice-pdf/")
async def slice_pdf(file: UploadFile = File(...), start_page: int = 1, end_page: int = 1):
    # Validasi input
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="File harus berformat DOCX")
    
    if start_page < 1 or end_page < 1 or start_page > end_page:
        raise HTTPException(status_code=400, detail="Start dan end page harus lebih besar dari 0 dan start_page harus kurang dari atau sama dengan end_page")

    # Simpan file DOCX yang diunggah ke direktori sementara
    temp_dir = tempfile.mkdtemp()
    temp_docx_path = os.path.join(temp_dir, file.filename)
    
    with open(temp_docx_path, "wb") as buffer:
        buffer.write(await file.read())

    # Konversi DOCX ke PDF
    temp_pdf_path = os.path.join(temp_dir, f"{file.filename[:-5]}.pdf")
    convert(temp_docx_path, temp_pdf_path)

    # Baca PDF dan potong halaman
    pdf_reader = PdfReader(temp_pdf_path)
    pdf_writer = PdfWriter()

    # Tambahkan halaman yang diinginkan ke PDF baru
    for page in range(start_page - 1, min(end_page, len(pdf_reader.pages))):
        pdf_writer.add_page(pdf_reader.pages[page])

    # Simpan PDF yang dipotong
    sliced_pdf_path = os.path.join(temp_dir, f"sliced_{file.filename[:-5]}.pdf")
    with open(sliced_pdf_path, "wb") as output_pdf:
        pdf_writer.write(output_pdf)

    # Kembalikan file PDF yang dipotong
    return FileResponse(sliced_pdf_path, media_type='application/pdf', filename=os.path.basename(sliced_pdf_path))




# --------------------------------------------------------------------------------------------

@app.post("/slice-pdf-docx/")
async def slice_pdf_docx(file: UploadFile = File(...), start_page: int = 1, end_page: int = 1, output_type: str = 'pdf'):
    # Validate input file type
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="File must be a DOCX format")
    
    # Validate page numbers
    if start_page < 1 or end_page < 1 or start_page > end_page:
        raise HTTPException(status_code=400, detail="Start and end page numbers must be greater than 0 and start_page must be less than or equal to end_page")

    # Save the uploaded DOCX file to a temporary directory
    temp_dir = tempfile.mkdtemp()
    temp_docx_path = os.path.join(temp_dir, file.filename)
    
    with open(temp_docx_path, "wb") as buffer:
        buffer.write(await file.read())

    # Convert DOCX to PDF
    temp_pdf_path = os.path.join(temp_dir, f"{file.filename[:-5]}.pdf")
    convert(temp_docx_path, temp_pdf_path)

    # Split the PDF into the specified page range
    sliced_pdf_path = slice_pdf(temp_pdf_path, start_page, end_page)

    # Return the output in the requested format
    if output_type == 'pdf':
        media_type = 'application/pdf'
        return FileResponse(sliced_pdf_path, media_type=media_type, filename=os.path.basename(sliced_pdf_path))
    elif output_type == 'docx':
        # Convert the sliced PDF back to DOCX
        sliced_docx_path = convert_pdf_to_docx(sliced_pdf_path)
        media_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        return FileResponse(sliced_docx_path, media_type=media_type, filename=os.path.basename(sliced_docx_path))

def slice_pdf(file_path, start_page, end_page):
    reader = PdfReader(file_path)
    total_pages = len(reader.pages)
    
    if start_page > total_pages or end_page > total_pages:
        raise ValueError("Start or end page number exceeds total pages in the document")

    writer = PdfWriter()
    for page_num in range(start_page - 1, end_page):
        writer.add_page(reader.pages[page_num])

    sliced_file_path = os.path.join(os.path.dirname(file_path), f"sliced_{os.path.basename(file_path)}")
    with open(sliced_file_path, "wb") as output_pdf:
        writer.write(output_pdf)

    return sliced_file_path

def convert_pdf_to_docx(pdf_path):
    docx_path = os.path.join(os.path.dirname(pdf_path), f"{os.path.basename(pdf_path)[:-4]}.docx")
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()
    return docx_path
    docx_path = os.path.join(os.path.dirname(pdf_path), f"{os.path.basename(pdf_path)[:-4]}.docx")
    cv = PdfToDocxConverter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()
    return docx_path