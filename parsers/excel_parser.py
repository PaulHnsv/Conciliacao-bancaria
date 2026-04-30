import re
import pandas as pd
from decimal import Decimal
from typing import List
from models.transaction import Transaction
from parsers.csv_parser import parse_decimal_br, parse_date_str


def _normalizar_col(nome: str) -> str:
    import unicodedata
    s = "".join(c for c in unicodedata.normalize("NFD", str(nome)) if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]", "", s.lower())


def parse_excel(filepath: str) -> List[Transaction]:
    """Le planilha Excel (xls/xlsx) de extrato bancario."""
    df = pd.read_excel(filepath, dtype=str).dropna(how="all")
    df.columns = df.columns.str.strip()

    col_map = {}
    for c in df.columns:
        chave = _normalizar_col(c)
        if chave == "data":
            col_map["data"] = c
        elif chave in ("historico", "descricao", "lancamento", "lancamentos"):
            col_map["descricao"] = c
        elif "valor" in chave and "saldo" not in chave:
            col_map["valor"] = c
        elif "saldo" in chave:
            col_map["saldo"] = c
        elif chave.startswith("credit") or "credito" in chave:
            col_map["credito"] = c
        elif chave.startswith("debit") or "debito" in chave:
            col_map["debito"] = c

    if "data" not in col_map:
        raise ValueError("Coluna de data nao encontrada na planilha")

    transacoes = []
    for _, row in df.iterrows():
        try:
            dt = parse_date_str(row[col_map["data"]])
        except (ValueError, KeyError):
            continue
        descricao = str(row.get(col_map.get("descricao", ""), "")).strip()
        if "valor" in col_map:
            valor = parse_decimal_br(row[col_map["valor"]])
        elif "credito" in col_map and "debito" in col_map:
            valor = parse_decimal_br(row[col_map["credito"]]) - parse_decimal_br(row[col_map["debito"]])
        else:
            valor = Decimal("0")
        tipo = "Crédito" if valor >= 0 else "Débito"
        saldo_col = col_map.get("saldo")
        saldo = parse_decimal_br(row[saldo_col]) if saldo_col else None
        transacoes.append(Transaction(
            data=dt,
            descricao=descricao,
            valor=valor,
            tipo=tipo,
            saldo=saldo,
        ))
    return sorted(transacoes, key=lambda t: t.data)
