import re
import pandas as pd
from decimal import Decimal
from datetime import datetime, date
from typing import List, Optional, Dict
from models.transaction import Transaction


def _sem_acentos(texto: str) -> str:
    import unicodedata
    return "".join(
        c for c in unicodedata.normalize("NFD", str(texto))
        if unicodedata.category(c) != "Mn"
    )


def _normalizar_col(nome: str) -> str:
    """Remove acentos, caracteres especiais e lowercase."""
    s = _sem_acentos(nome)
    s = re.sub(r"[^a-z0-9]", "", s.lower())
    return s


def _mapear_colunas(df_cols: list) -> Dict[str, str]:
    """Mapeia nome real da coluna para chave normalizada."""
    return {_normalizar_col(c): c for c in df_cols}


def parse_decimal_br(valor_str) -> Decimal:
    """Converte string monetaria br para Decimal. Suporta 1.000,00 e 1000.00."""
    if pd.isna(valor_str) or str(valor_str).strip() == "":
        return Decimal("0")
    s = str(valor_str).strip().replace(" ", "")
    # Detecta formato: se tem virgula e ponto, virgula eh decimal
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    # else: ponto como decimal, sem alteracao
    try:
        return Decimal(s)
    except Exception:
        return Decimal("0")


def parse_date_str(valor_str: str) -> date:
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return datetime.strptime(str(valor_str).strip(), fmt).date()
        except ValueError:
            continue
    raise ValueError("Data nao reconhecida: " + str(valor_str))


def _tentar_leitura(filepath: str, sep: str, encoding: str):
    df = pd.read_csv(filepath, sep=sep, encoding=encoding, dtype=str).dropna(how="all")
    df.columns = df.columns.str.strip()
    return df


def _detectar_banco(col_keys: set) -> Optional[str]:
    """Detecta o banco pelo conjunto de chaves de colunas normalizadas."""
    # Itau: tem colunas "creditoR$" e "debitoR$" — normalizadas ficam "creditors" "debitors"
    if "creditors" in col_keys or any("creditor" in k for k in col_keys):
        return "itau"
    # Inter / BB: tem "credito" E "debito" separados
    if "credito" in col_keys and "debito" in col_keys:
        if "lancamento" in col_keys:
            return "bb"
        return "inter"
    # Santander: tem "descricao" + "valor" + "saldo" (com ponto-e-virgula)
    if "descricao" in col_keys and "valor" in col_keys and "saldo" in col_keys and "historico" not in col_keys:
        return "santander"
    # Nubank: tem "descricao" e "valor" (sem saldo, usa virgula como separador)
    if "descricao" in col_keys and "valor" in col_keys and "historico" not in col_keys:
        return "nubank"
    # Bradesco: tem "historico" + "valor" + "saldo"
    if "historico" in col_keys and "valor" in col_keys and "saldo" in col_keys:
        return "bradesco"
    return None


def _extrair_transacoes(df: pd.DataFrame, col_keys: Dict[str, str], banco: str) -> List[Transaction]:
    transacoes = []
    # Mapeia as chaves reais a partir das normalizadas
    def col(chave_norm: str) -> Optional[str]:
        return col_keys.get(chave_norm)

    col_data = col("data")
    col_descricao = col("descricao") or col("historico") or col("lancamento")
    col_valor = col("valor")
    col_credito = col("credito") or next((col_keys[k] for k in col_keys if k.startswith("creditor")), None)
    col_debito = col("debito") or next((col_keys[k] for k in col_keys if k.startswith("debitor")), None)
    col_saldo = col("saldo")

    if not col_data:
        return []

    for _, row in df.iterrows():
        try:
            dt = parse_date_str(row[col_data])
        except (ValueError, KeyError):
            continue
        descricao = str(row.get(col_descricao, "")).strip() if col_descricao else ""
        if col_valor:
            valor = parse_decimal_br(row[col_valor])
        elif col_credito and col_debito:
            cred = parse_decimal_br(row.get(col_credito, ""))
            deb = parse_decimal_br(row.get(col_debito, ""))
            valor = cred - deb
        else:
            valor = Decimal("0")
        tipo = "Crédito" if valor >= 0 else "Débito"
        saldo = parse_decimal_br(row[col_saldo]) if col_saldo and not pd.isna(row.get(col_saldo, float("nan"))) else None
        transacoes.append(Transaction(
            data=dt,
            descricao=descricao,
            valor=valor,
            tipo=tipo,
            saldo=saldo,
            banco_detectado=banco,
        ))
    return sorted(transacoes, key=lambda t: t.data)


def parse_csv(filepath: str) -> List[Transaction]:
    """Detecta o banco pelo header do CSV e normaliza as transacoes."""
    for sep in (";", ",", "\t"):
        for encoding in ("utf-8", "latin-1", "utf-8-sig"):
            try:
                df = _tentar_leitura(filepath, sep, encoding)
                if len(df.columns) < 2:
                    continue
                col_keys = _mapear_colunas(df.columns.tolist())
                banco = _detectar_banco(set(col_keys.keys()))
                if banco is None:
                    continue
                result = _extrair_transacoes(df, col_keys, banco)
                if result:
                    return result
            except Exception:
                continue
    raise ValueError("Formato de CSV nao reconhecido: " + filepath)
