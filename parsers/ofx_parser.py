from decimal import Decimal
from datetime import date
from typing import List
from ofxparse import OfxParser
from models.transaction import Transaction


def parse_ofx(filepath: str) -> List[Transaction]:
    """Le arquivo OFX/OFC e retorna lista de Transaction."""
    with open(filepath, "rb") as f:
        ofx = OfxParser.parse(f)

    transacoes = []
    for account in ofx.accounts:
        for t in account.statement.transactions:
            valor = Decimal(str(t.amount))
            tipo = "Crédito" if valor >= 0 else "Débito"
            dt = t.date.date() if hasattr(t.date, "date") else t.date
            transacoes.append(Transaction(
                data=dt,
                descricao=t.memo or t.payee or "",
                valor=valor,
                tipo=tipo,
                banco_detectado=None,
            ))

    return sorted(transacoes, key=lambda t: t.data)
