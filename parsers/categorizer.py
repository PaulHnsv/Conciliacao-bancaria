"""
Categorizador automatico de transacoes bancarias.
Agrupa descricoes por palavras-chave em categorias de gastos.
"""
import re
from typing import List, Tuple
from models.transaction import Transaction

# Ordem importa: primeiro match ganha
_REGRAS_CREDITO = [
    ("Salario",         [r"salari", r"credito de sal", r"pagamento de salari"]),
    ("Rendimentos",     [r"remuneracao aplicacao", r"rendimento", r"juros", r"cdb", r"lci", r"lca"]),
    ("PIX Recebido",    [r"pix recebido", r"pix enviado.*recebido"]),
    ("TED/DOC Recebido",[r"ted recebida", r"doc recebido", r"transferencia recebida"]),
    ("Estorno",         [r"estorno", r"devolu"]),
    ("Outros Creditos", [r"."]),  # catchall
]

_REGRAS_DEBITO = [
    ("Cartao de Credito",  [r"pagamento cartao", r"pagto cartao", r"fatura cartao", r"cartao visa.*pagamento", r"pagamento de cartao"]),
    ("Debito Automatico",  [r"debito autom", r"debito visa", r"mensalidade", r"assinatura", r"recorrente"]),
    ("Boleto",             [r"boleto", r"pagamento de boleto", r"cedente", r"concessionaria"]),
    ("PIX Enviado",        [r"pix enviado", r"pix.*enviado"]),
    ("Transferencia",      [r"ted enviada", r"doc enviado", r"transferencia enviada", r"transf enviada"]),
    ("Saque/IOF",          [r"saque", r"iof", r"tarifa"]),
    ("Outros Debitos",     [r"."]),  # catchall
]


def _sem_acentos(texto: str) -> str:
    import unicodedata
    return "".join(
        c for c in unicodedata.normalize("NFD", str(texto))
        if unicodedata.category(c) != "Mn"
    ).lower()


def categorizar(transacao: Transaction) -> str:
    """Retorna a categoria da transacao baseado na descricao."""
    desc = _sem_acentos(transacao.descricao)
    regras = _REGRAS_CREDITO if transacao.tipo == "Credito" else _REGRAS_DEBITO
    for categoria, padroes in regras:
        for padrao in padroes:
            if re.search(padrao, desc):
                return categoria
    return "Outros Creditos" if transacao.tipo == "Credito" else "Outros Debitos"


def agrupar_por_categoria(transacoes: List[Transaction]) -> dict:
    """
    Retorna dict: { categoria: {"total": Decimal, "count": int, "tipo": str} }
    Apenas transacoes do tipo Debito sao agrupadas para analise de gastos.
    """
    from decimal import Decimal
    from collections import defaultdict
    grupos = defaultdict(lambda: {"total": Decimal("0"), "count": 0, "tipo": ""})
    for t in transacoes:
        cat = categorizar(t)
        grupos[cat]["total"] += abs(t.valor)
        grupos[cat]["count"] += 1
        grupos[cat]["tipo"] = t.tipo
    return dict(grupos)
