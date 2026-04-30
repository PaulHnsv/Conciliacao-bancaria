import pytest
from decimal import Decimal
from datetime import date
from models.transaction import Transaction


def test_transaction_credito():
    t = Transaction(
        data=date(2024, 1, 15),
        descricao="PIX RECEBIDO - JOAO",
        valor=Decimal("1000.00"),
        tipo="Crédito",
        saldo=Decimal("5000.00"),
        banco_detectado="inter"
    )
    assert t.valor > 0
    assert t.tipo == "Crédito"


def test_transaction_debito():
    t = Transaction(
        data=date(2024, 1, 16),
        descricao="COMPRA SUPERMERCADO",
        valor=Decimal("-150.00"),
        tipo="Débito",
    )
    assert t.valor < 0
    assert t.saldo is None
    assert t.banco_detectado is None


def test_transaction_campos_obrigatorios():
    with pytest.raises(TypeError):
        Transaction()  # sem argumentos obrigatórios
