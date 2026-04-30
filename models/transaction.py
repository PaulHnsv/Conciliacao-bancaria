from dataclasses import dataclass
from decimal import Decimal
from datetime import date
from typing import Optional


@dataclass
class Transaction:
    data: date
    descricao: str
    valor: Decimal        # positivo = crédito, negativo = débito
    tipo: str             # "Crédito" | "Débito"
    saldo: Optional[Decimal] = None
    banco_detectado: Optional[str] = None
