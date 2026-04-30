"""
Parser do arquivo de configuracao de limites por categoria.
"""
import os
import re
from typing import Dict

_DEFAULT_LIMITES = {
    "PIX Enviado": 30,
    "Cartao de Credito": 25,
    "Debito Automatico": 20,
    "Boleto": 15,
    "Transferencia": 20,
    "Saque/IOF": 5,
    "Outros Debitos": 10,
}

_ARQUIVO = os.path.join(os.path.dirname(__file__), "limites.txt")


def carregar_limites() -> Dict[str, float]:
    """Carrega limites do arquivo config/limites.txt. Usa defaults se nao existir."""
    limites = dict(_DEFAULT_LIMITES)
    if not os.path.exists(_ARQUIVO):
        return limites
    with open(_ARQUIVO, encoding="utf-8") as f:
        for linha in f:
            linha = linha.strip()
            if not linha or linha.startswith("#"):
                continue
            if "=" in linha:
                chave, _, valor = linha.partition("=")
                chave = chave.strip()
                try:
                    limites[chave] = float(valor.strip())
                except ValueError:
                    pass
    return limites
