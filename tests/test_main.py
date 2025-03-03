import pytest
from utils.pdf_utils import split_pdf
import os

def test_split_pdf():
    # Buat file PDF sementara
    file_path = "test.pdf"
    with open(file_path, "wb") as f:
        f.write(b"%PDF-1.4\n")

    # Memecah PDF
    split_file_path = split_pdf(file_path, 1, 1)

    # Periksa apakah file hasil pemisahan ada
    assert os.path.exists(split_file_path)

    # Hapus file sementara
    os.remove(file_path)
    os.remove(split_file_path)

def test_split_pdf_invalid_page():
    # Buat file PDF sementara
    file_path = "test.pdf"
    with open(file_path, "wb") as f:
        f.write(b"%PDF-1.4\n")

    # Coba memecah dengan nomor halaman tidak valid
    with pytest.raises(ValueError):
        split_pdf(file_path, 0, 1)

    # Hapus file sementara
    os.remove(file_path)
