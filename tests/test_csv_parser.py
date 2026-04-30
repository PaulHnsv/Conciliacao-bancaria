import os
from decimal import Decimal
from datetime import date
from parsers.csv_parser import parse_csv

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")


def test_parse_nubank():
    path = os.path.join(FIXTURES, "nubank.csv")
    t = parse_csv(path)
    assert len(t) == 2
    assert t[0].banco_detectado == "nubank"
    assert t[0].tipo == "Crédito"
    assert t[0].valor == Decimal("1000.00")
    assert t[0].data == date(2024, 1, 15)


def test_parse_inter():
    path = os.path.join(FIXTURES, "inter.csv")
    t = parse_csv(path)
    assert len(t) == 2
    assert t[0].banco_detectado == "inter"
    assert t[0].tipo == "Crédito"
    assert t[0].valor == Decimal("1000.00")
    debito = t[1]
    assert debito.tipo == "Débito"
    assert debito.valor == Decimal("-200.00")


def test_parse_itau():
    path = os.path.join(FIXTURES, "itau.csv")
    t = parse_csv(path)
    assert len(t) == 2
    assert t[0].banco_detectado == "itau"
    assert t[0].valor == Decimal("1000.00")
    assert t[1].valor == Decimal("-150.00")


def test_parse_bradesco():
    path = os.path.join(FIXTURES, "bradesco.csv")
    t = parse_csv(path)
    assert len(t) == 2
    assert t[0].banco_detectado == "bradesco"
    assert t[1].valor == Decimal("-200.00")


def test_parse_bb():
    from decimal import Decimal
    path = os.path.join(FIXTURES, "bb.csv")
    t = parse_csv(path)
    assert len(t) == 2
    assert t[0].banco_detectado == "bb"
    assert t[0].valor == Decimal("1000.00")
    assert t[1].valor == Decimal("-200.00")


def test_parse_santander():
    from decimal import Decimal
    path = os.path.join(FIXTURES, "santander.csv")
    t = parse_csv(path)
    assert len(t) == 2
    assert t[0].banco_detectado == "santander"
    assert t[0].valor == Decimal("1000.00")
    assert t[1].valor == Decimal("-200.00")
