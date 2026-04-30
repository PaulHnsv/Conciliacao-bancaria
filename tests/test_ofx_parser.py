import os
import pytest
from decimal import Decimal
from datetime import date
from parsers.ofx_parser import parse_ofx

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")


def test_parse_ofx_retorna_lista_de_transacoes():
    path = os.path.join(FIXTURES, "sample.ofx")
    transacoes = parse_ofx(path)
    assert len(transacoes) == 2


def test_parse_ofx_credito():
    path = os.path.join(FIXTURES, "sample.ofx")
    transacoes = parse_ofx(path)
    credito = next(t for t in transacoes if t.tipo == "Crédito")
    assert credito.valor == Decimal("1000.00")
    assert credito.descricao == "PIX RECEBIDO - JOAO"
    assert credito.data == date(2024, 1, 15)


def test_parse_ofx_debito():
    path = os.path.join(FIXTURES, "sample.ofx")
    transacoes = parse_ofx(path)
    debito = next(t for t in transacoes if t.tipo == "Débito")
    assert debito.valor == Decimal("-150.00")
    assert debito.descricao == "COMPRA SUPERMERCADO"
    assert debito.data == date(2024, 1, 16)


def test_parse_ofx_valores_decimal():
    path = os.path.join(FIXTURES, "sample.ofx")
    transacoes = parse_ofx(path)
    for t in transacoes:
        assert isinstance(t.valor, Decimal)
