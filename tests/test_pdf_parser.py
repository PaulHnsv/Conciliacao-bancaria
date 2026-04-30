import pytest
from unittest.mock import patch, MagicMock
from decimal import Decimal
from datetime import date
from parsers.pdf_parser import parse_pdf, _extrair_transacoes_do_texto

TEXTO_PDF = """
15/01/2024  PIX recebido Joao                   1.000,00
16/01/2024  iFood*Restaurante                     -45,90
17/01/2024  Boleto energia                        -120,00
"""


def test_extrair_transacoes_de_texto():
    transacoes = _extrair_transacoes_do_texto(TEXTO_PDF)
    assert len(transacoes) == 3


def test_extrair_credito():
    transacoes = _extrair_transacoes_do_texto(TEXTO_PDF)
    credito = next(t for t in transacoes if t.tipo == "Crédito")
    assert credito.valor == Decimal("1000.00")
    assert credito.data == date(2024, 1, 15)


def test_extrair_debito():
    transacoes = _extrair_transacoes_do_texto(TEXTO_PDF)
    debitos = [t for t in transacoes if t.tipo == "Débito"]
    assert len(debitos) == 2


def test_parse_pdf_usa_pdfplumber(tmp_path):
    with patch("parsers.pdf_parser.pdfplumber") as mock_plumber:
        mock_page = MagicMock()
        mock_page.extract_text.return_value = TEXTO_PDF
        mock_pdf = MagicMock()
        mock_pdf.__enter__ = MagicMock(return_value=mock_pdf)
        mock_pdf.__exit__ = MagicMock(return_value=False)
        mock_pdf.pages = [mock_page]
        mock_plumber.open.return_value = mock_pdf
        f = tmp_path / "extrato.pdf"
        f.write_bytes(b"%PDF fake")
        result = parse_pdf(str(f))
        assert len(result) == 3
