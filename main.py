from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from utils.pdf_utils import split_pdf
from utils.docx_utils import convert_to_pdf
import os
from datetime import datetime


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

@app.post("/split-doc/")
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
        # Konversi ke PDF jika file adalah DOCX atau DOC
        if file_extension in [".docx", ".doc"]:
            pdf_location = f"temp_{os.path.splitext(file.filename)[0]}.pdf"
            convert_to_pdf(file_location, pdf_location)

            # Pisahkan PDF
            split_pdf_location = split_pdf(pdf_location, start_page, end_page)

            # Konversi kembali ke DOCX
            final_docx_location = f"hasil-split/split_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            convert_to_docx(split_pdf_location, final_docx_location)

            return FileResponse(path=final_docx_location, filename=os.path.basename(final_docx_location))

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        # Hapus file sementara
        os.remove(file_location)
        if os.path.exists(pdf_location):
            os.remove(pdf_location)
        if os.path.exists(split_pdf_location):
            os.remove(split_pdf_location)
    # Validasi tipe file
    valid_extensions = [".pdf", ".docx", ".doc"]
    file_extension = os.path.splitext(file.filename)[1].lower()
    if file_extension not in valid_extensions:
        raise HTTPException(status_code=400, detail="Unsupported file type")

    # Simpan file yang diunggah
    file_location = f"hasil-split/temp_{file.filename}"
    with open(file_location, "wb+") as file_object:
        file_object.write(file.file.read())

    try:
        # Konversi ke PDF jika file adalah DOCX atau DOC
        if file_extension in [".docx", ".doc"]:
            pdf_location = f"hasil-split/temp_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            convert_to_pdf(file_location, pdf_location)
            split_file_location = split_pdf(pdf_location, start_page, end_page)
        else:
            split_file_location = split_pdf(file_location, start_page, end_page)

        # Konversi kembali ke DOCX jika file awal adalah DOCX
        if file_extension in [".docx", ".doc"]:
            final_docx_location = f"hasil-split/split_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            convert_to_docx(split_file_location, final_docx_location)
            return FileResponse(path=final_docx_location, filename=os.path.basename(final_docx_location))

        return FileResponse(path=split_file_location, filename=os.path.basename(split_file_location))

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        # Hapus file sementara
        os.remove(file_location)
        if 'pdf_location' in locals() and os.path.exists(pdf_location):
            os.remove(pdf_location)
        if 'split_file_location' in locals() and os.path.exists(split_file_location):
            os.remove(split_file_location)
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
        # Konversi ke PDF jika file adalah DOCX atau DOC
        if file_extension in [".docx", ".doc"]:
            pdf_location = f"temp_{os.path.splitext(file.filename)[0]}.pdf"
            convert_to_pdf(file_location, pdf_location)

            # Memisahkan PDF
            split_pdf_location = split_pdf(pdf_location, start_page, end_page)

            # Konversi kembali ke DOCX
            final_docx_location = f"hasil-split/{datetime.now().strftime('%Y%m%d_%H%M%S')}_{os.path.splitext(file.filename)[0]}.docx"
            convert_to_docx(split_pdf_location, final_docx_location)

            return FileResponse(path=final_docx_location, filename=os.path.basename(final_docx_location))

        elif file_extension == ".pdf":
            split_pdf_location = split_pdf(file_location, start_page, end_page)
            return FileResponse(path=split_pdf_location, filename=os.path.basename(split_pdf_location))

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        # Hapus file sementara
        os.remove(file_location)
        if 'pdf_location' in locals() and os.path.exists(pdf_location):
            os.remove(pdf_location)
        if 'split_pdf_location' in locals() and os.path.exists(split_pdf_location):
            os.remove(split_pdf_location)
    # Validasi tipe file
    valid_extensions = [".pdf", ".docx", ".doc"]
    file_extension = os.path.splitext(file.filename)[1].lower()
    if file_extension not in valid_extensions:
        raise HTTPException(status_code=400, detail="Unsupported file type")

    # Simpan file yang diunggah
    file_location = f"temp_{file.filename}"
    with open(file_location, "wb+") as file_object:
        file_object.write(file.file.read())

    split_pdf_location = None  # Inisialisasi variabel

    try:
        # Konversi ke PDF jika file bukan PDF
        if file_extension in [".docx", ".doc"]:
            pdf_file_location = convert_to_pdf(file_location)
            split_pdf_location = split_pdf(pdf_file_location, start_page, end_page)
            # Konversi kembali ke DOCX
            final_file_location = convert_to_docx(split_pdf_location)
        else:
            split_pdf_location = split_pdf(file_location, start_page, end_page)
            final_file_location = split_pdf_location  # Jika sudah PDF

        # Mengirimkan file yang telah dipisah
        return FileResponse(path=final_file_location, filename=f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_split.docx")

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        # Hapus file sementara
        os.remove(file_location)
        if split_pdf_location and os.path.exists(split_pdf_location):
            os.remove(split_pdf_location)
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
        # Konversi ke PDF jika file adalah DOCX atau DOC
        if file_extension in [".docx", ".doc"]:
            pdf_location = f"temp_{os.path.splitext(file.filename)[0]}.pdf"
            convert_to_pdf(file_location, pdf_location)

            # Split PDF
            split_pdf_location = split_pdf(pdf_location, start_page, end_page)

            # Konversi kembali ke DOCX
            docx_location = f"hasil-split/split_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            convert_to_docx(split_pdf_location, docx_location)

            return FileResponse(path=docx_location, filename=os.path.basename(docx_location))

        elif file_extension == ".pdf":
            # Split PDF langsung
            split_pdf_location = split_pdf(file_location, start_page, end_page)

            # Konversi kembali ke DOCX
            docx_location = f"hasil-split/split_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            convert_to_docx(split_pdf_location, docx_location)

            return FileResponse(path=docx_location, filename=os.path.basename(docx_location))

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        # Hapus file sementara
        os.remove(file_location)
        if os.path.exists(pdf_location):
            os.remove(pdf_location)
        if os.path.exists(split_pdf_location):
            os.remove(split_pdf_location)
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
        # Konversi ke PDF jika file bukan PDF
        if file_extension != ".pdf":
            pdf_file_location = convert_to_pdf(file_location)
        else:
            pdf_file_location = file_location

        # Memisahkan PDF
        split_pdf_location = split_pdf(pdf_file_location, start_page, end_page)

        # Konversi kembali ke DOCX
        final_docx_location = convert_to_docx(split_pdf_location)

        # Mengirimkan file yang telah dipisah
        return FileResponse(path=final_docx_location, filename=f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_split.docx")

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        # Hapus file sementara
        os.remove(file_location)
        if 'pdf_file_location' in locals():
            os.remove(pdf_file_location)
        if 'split_pdf_location' in locals():
            os.remove(split_pdf_location)
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
        # Konversi ke PDF
        pdf_file_location = convert_to_pdf(file_location)

        # Memisahkan PDF
        split_pdf_location = split_pdf(pdf_file_location, start_page, end_page)

        # Konversi kembali ke DOCX
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_docx_location = f"hasil-split/{timestamp}_split_document.docx"
        convert_to_docx(split_pdf_location, output_docx_location)

        return FileResponse(path=output_docx_location, filename=os.path.basename(output_docx_location))

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        # Hapus file sementara
        os.remove(file_location)
        if os.path.exists(pdf_file_location):
            os.remove(pdf_file_location)
        if os.path.exists(split_pdf_location):
            os.remove(split_pdf_location)
    # Validasi tipe file
    valid_extensions = [".docx", ".doc"]
    file_extension = os.path.splitext(file.filename)[1].lower()
    if file_extension not in valid_extensions:
        raise HTTPException(status_code=400, detail="Unsupported file type")

    # Simpan file yang diunggah
    file_location = f"temp_{file.filename}"
    with open(file_location, "wb+") as file_object:
        file_object.write(file.file.read())

    try:
        # Memecah file berdasarkan tipe
        split_file_location = split_docx(file_location, start_page, end_page)

        # Mengirimkan file yang telah dipisah
        return FileResponse(path=split_file_location, filename=f"split_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{file.filename}")

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        # Hapus file sementara
        os.remove(file_location)
    # Validasi tipe file
    valid_extensions = [".docx", ".doc"]
    file_extension = os.path.splitext(file.filename)[1].lower()
    if file_extension not in valid_extensions:
        raise HTTPException(status_code=400, detail="Unsupported file type")

    # Simpan file yang diunggah
    file_location = f"temp_{file.filename}"
    with open(file_location, "wb+") as file_object:
        file_object.write(file.file.read())

    split_file_location = None  # Inisialisasi variabel

    try:
        # Memecah file berdasarkan tipe
        if file_extension in [".docx", ".doc"]:
            split_file_location = split_docx(file_location, start_page, end_page)
        else:
            raise HTTPException(status_code=400, detail="Unsupported file type for splitting")

        # Mengirimkan file yang telah dipisah
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        new_filename = f"{timestamp}_{os.path.basename(split_file_location)}"
        return FileResponse(path=split_file_location, filename=new_filename)

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        # Hapus file sementara
        os.remove(file_location)
        if split_file_location and os.path.exists(split_file_location):
            os.remove(split_file_location)

@app.post("/convert-docx-to-pdf/")
async def convert_docx_to_pdf(file: UploadFile = File(...), start: int = 1, end: int = 1):
    # Validasi input
    if start < 1 or end < 1:
        raise HTTPException(status_code=400, detail="Start dan end harus lebih besar dari 0")

    # Simpan file DOCX yang diunggah ke direktori sementara
    temp_dir = tempfile.mkdtemp(dir="temp")
    temp_docx_path = os.path.join(temp_dir, file.filename)
    with open(temp_docx_path, "wb") as buffer:
        buffer.write(await file.read())

    # Konversi DOCX ke PDF
    try:
        temp_pdf_path = convert_to_pdf(temp_docx_path, start, end)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Konversi gagal: {str(e)}")

    # Hapus file DOCX sementara
    os.remove(temp_docx_path)

    # Kembalikan file PDF sebagai respons
    return FileResponse(temp_pdf_path, filename=f"{os.path.splitext(file.filename)[0]}.pdf", media_type="application/pdf")

