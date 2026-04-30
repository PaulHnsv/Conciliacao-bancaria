import os
import pytest
from parsers.detector import detect_format

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")


def test_detecta_ofx_pela_extensao():
    path = os.path.join(FIXTURES, "sample.ofx")
    assert detect_format(path) == "ofx"


def test_detecta_csv_pela_extensao():
    path = os.path.join(FIXTURES, "nubank.csv")
    assert detect_format(path) == "csv"


def test_detecta_ofx_pelo_conteudo(tmp_path):
    f = tmp_path / "extrato.txt"
    f.write_text("OFXHEADER:100\nDATA:OFXSGML\n<OFX>\n")
    assert detect_format(str(f)) == "ofx"


def test_detecta_pdf_pela_extensao(tmp_path):
    f = tmp_path / "extrato.pdf"
    f.write_bytes(b"%PDF-1.4 fake content")
    assert detect_format(str(f)) == "pdf"


def test_detecta_excel_pela_extensao(tmp_path):
    import openpyxl
    wb = openpyxl.Workbook()
    path = str(tmp_path / "extrato.xlsx")
    wb.save(path)
    assert detect_format(path) == "excel"


def test_formato_desconhecido_levanta_erro(tmp_path):
    f = tmp_path / "extrato.zip"
    f.write_bytes(b"PK\x03\x04")
    with pytest.raises(ValueError, match="Formato não suportado"):
        detect_format(str(f))
