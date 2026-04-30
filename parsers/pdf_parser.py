import re
import pdfplumber
from decimal import Decimal, InvalidOperation
from datetime import datetime
from typing import List
from models.transaction import Transaction

_PADRAO_LINHA = re.compile(
    r"(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([-+]?\d{1,3}(?:\.\d{3})*(?:,\d{2}))\s*$"
)


def _parse_valor_br(s: str) -> Decimal:
    s = s.strip().replace(".", "").replace(",", ".")
    return Decimal(s)


def _extrair_transacoes_do_texto(texto: str) -> List[Transaction]:
    transacoes = []
    for linha in texto.splitlines():
        linha = linha.strip()
        m = _PADRAO_LINHA.search(linha)
        if not m:
            continue
        data_str = m.group(1)
        descricao = m.group(2).strip()
        valor_str = m.group(3)
        try:
            dt = datetime.strptime(data_str, "%d/%m/%Y").date()
            valor = _parse_valor_br(valor_str)
            tipo = "Crédito" if valor >= 0 else "Débito"
            transacoes.append(Transaction(
                data=dt,
                descricao=descricao,
                valor=valor,
                tipo=tipo,
                banco_detectado="pdf",
            ))
        except (ValueError, InvalidOperation):
            continue
    return sorted(transacoes, key=lambda t: t.data)


def parse_pdf(filepath: str) -> List[Transaction]:
    """Extrai transacoes de PDF bancario usando pdfplumber."""
    texto_completo = []
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                texto_completo.append(t)
    return _extrair_transacoes_do_texto("\n".join(texto_completo))
