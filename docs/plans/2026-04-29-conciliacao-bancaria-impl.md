# Plano de Implementação — Conciliação Bancária Automática
**Data:** 2026-04-29  
**Spec de referência:** `docs/specs/2026-04-29-conciliacao-bancaria-design.md`  
**Status:** Pronto para execução

---

## Objetivo

Ferramenta Python + `.bat` para Windows que:
1. Abre seletor de arquivo nativo
2. Detecta formato automaticamente (OFX, CSV, Excel, PDF)
3. Normaliza as transações
4. Gera relatório Excel formatado na mesma pasta do arquivo de entrada

---

## Tech Stack

- **Python 3.8+**
- `ofxparse` — leitura de OFX/OFC
- `pandas` — leitura e normalização de CSV/Excel
- `openpyxl` — geração e leitura de `.xlsx`
- `pdfplumber` — extração de texto de PDF
- `tkinter` — seletor de arquivo (nativo, já incluso no Python)
- `pytest` — testes

---

## Estrutura de Arquivos

```
conciliacao-bancaria/
├── conciliar.bat               # Ponto de entrada Windows
├── main.py                     # Orquestrador principal
├── requirements.txt            # Dependências
├── models/
│   ├── __init__.py
│   └── transaction.py          # Dataclass Transaction
├── parsers/
│   ├── __init__.py
│   ├── detector.py             # Detecta formato do arquivo
│   ├── ofx_parser.py           # Parser OFX/OFC
│   ├── csv_parser.py           # Parser CSV (multi-banco)
│   ├── excel_parser.py         # Parser XLS/XLSX de entrada
│   └── pdf_parser.py           # Parser PDF (pdfplumber)
├── report/
│   ├── __init__.py
│   └── excel_report.py         # Gerador do relatório Excel
└── tests/
    ├── __init__.py
    ├── fixtures/
    │   ├── sample.ofx
    │   ├── nubank.csv
    │   ├── inter.csv
    │   ├── itau.csv
    │   └── bradesco.csv
    ├── test_transaction.py
    ├── test_detector.py
    ├── test_ofx_parser.py
    ├── test_csv_parser.py
    ├── test_excel_parser.py
    ├── test_pdf_parser.py
    └── test_excel_report.py
```

---

## Task 1: Setup do Projeto + Modelo Transaction

**Files:** `requirements.txt`, `models/__init__.py`, `models/transaction.py`, `tests/test_transaction.py`

**Steps:**

- [ ] Criar `requirements.txt`:
  ```
  ofxparse==0.21
  pandas==2.2.2
  openpyxl==3.1.2
  pdfplumber==0.11.0
  pytest==8.2.0
  ```

- [ ] Criar `models/__init__.py` (vazio):
  ```python
  ```

- [ ] Criar `tests/test_transaction.py`:
  ```python
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
  ```

- [ ] Rodar: `python -m pytest tests/test_transaction.py -v` — esperar: **FAIL** (módulo não existe)

- [ ] Criar `models/transaction.py`:
  ```python
  from dataclasses import dataclass, field
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
  ```

- [ ] Rodar: `python -m pytest tests/test_transaction.py -v` — esperar: **PASS**

- [ ] Commit: `git commit -m "feat: Transaction dataclass com campos data, descricao, valor, tipo, saldo, banco_detectado"`

---

## Task 2: Detector de Formato

**Files:** `parsers/__init__.py`, `parsers/detector.py`, `tests/test_detector.py`

**Steps:**

- [ ] Criar `parsers/__init__.py` (vazio)

- [ ] Criar `tests/fixtures/sample.ofx`:
  ```
  OFXHEADER:100
  DATA:OFXSGML
  VERSION:102
  SECURITY:NONE
  ENCODING:UTF-8
  CHARSET:1252
  COMPRESSION:NONE
  OLDFILEUID:NONE
  NEWFILEUID:NONE

  <OFX>
  <SIGNONMSGSRSV1>
  <SONRS>
  <STATUS><CODE>0<SEVERITY>INFO</STATUS>
  <DTSERVER>20240115120000
  <LANGUAGE>POR
  </SONRS>
  </SIGNONMSGSRSV1>
  <BANKMSGSRSV1>
  <STMTTRNRS>
  <TRNUID>1
  <STMTRS>
  <CURDEF>BRL
  <BANKACCTFROM><BANKID>341<ACCTID>12345-6<ACCTTYPE>CHECKING</BANKACCTFROM>
  <BANKTRANLIST>
  <DTSTART>20240101
  <DTEND>20240131
  <STMTTRN>
  <TRNTYPE>CREDIT
  <DTPOSTED>20240115
  <TRNAMT>1000.00
  <FITID>001
  <MEMO>PIX RECEBIDO - JOAO
  </STMTTRN>
  <STMTTRN>
  <TRNTYPE>DEBIT
  <DTPOSTED>20240116
  <TRNAMT>-150.00
  <FITID>002
  <MEMO>COMPRA SUPERMERCADO
  </STMTTRN>
  </BANKTRANLIST>
  <LEDGERBAL><BALAMT>4850.00<DTASOF>20240131</LEDGERBAL>
  </STMTRS>
  </STMTTRNRS>
  </BANKMSGSRSV1>
  </OFX>
  ```

- [ ] Criar `tests/fixtures/nubank.csv`:
  ```
  Data,Descrição,Valor
  2024-01-15,"PIX recebido",1000.00
  2024-01-16,"iFood*RESTAURANTE",-45.90
  ```

- [ ] Criar `tests/test_detector.py`:
  ```python
  import os
  import pytest
  from parsers.detector import detect_format

  FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")

  def test_detecta_ofx_pela_extensao():
      path = os.path.join(FIXTURES, "sample.ofx")
      assert detect_format(path) == "ofx"

  def test_detecta_csv_pela_extensao():
      path = os.path.join(FIXTURES, "nubank.csv")
      assert detect_format(path) == "csv"

  def test_detecta_ofx_pelo_conteudo(tmp_path):
      f = tmp_path / "extrato.txt"
      f.write_text("OFXHEADER:100\nDATA:OFXSGML\n<OFX>\n")
      assert detect_format(str(f)) == "ofx"

  def test_detecta_pdf_pela_extensao(tmp_path):
      f = tmp_path / "extrato.pdf"
      f.write_bytes(b"%PDF-1.4 fake content")
      assert detect_format(str(f)) == "pdf"

  def test_detecta_excel_pela_extensao(tmp_path):
      import openpyxl
      wb = openpyxl.Workbook()
      path = str(tmp_path / "extrato.xlsx")
      wb.save(path)
      assert detect_format(path) == "excel"

  def test_formato_desconhecido_levanta_erro(tmp_path):
      f = tmp_path / "extrato.zip"
      f.write_bytes(b"PK\x03\x04")
      with pytest.raises(ValueError, match="Formato não suportado"):
          detect_format(str(f))
  ```

- [ ] Rodar: `python -m pytest tests/test_detector.py -v` — esperar: **FAIL**

- [ ] Criar `parsers/detector.py`:
  ```python
  import os

  _EXTENSOES = {
      ".ofx": "ofx",
      ".ofc": "ofx",
      ".csv": "csv",
      ".xls": "excel",
      ".xlsx": "excel",
      ".pdf": "pdf",
  }

  _ASSINATURAS_CONTEUDO = [
      (b"OFXHEADER", "ofx"),
      (b"<OFX>", "ofx"),
      (b"%PDF", "pdf"),
  ]

  def detect_format(filepath: str) -> str:
      """Detecta o formato do arquivo de extrato bancário."""
      ext = os.path.splitext(filepath)[1].lower()
      if ext in _EXTENSOES:
          return _EXTENSOES[ext]

      # Fallback: inspeciona conteúdo
      try:
          with open(filepath, "rb") as f:
              head = f.read(512)
          for assinatura, formato in _ASSINATURAS_CONTEUDO:
              if assinatura in head:
                  return formato
          # Tenta como CSV se for texto legível
          head.decode("utf-8")
          return "csv"
      except UnicodeDecodeError:
          pass

      raise ValueError(
          f"Formato não suportado para o arquivo: {os.path.basename(filepath)}\n"
          "Formatos aceitos: .ofx, .ofc, .csv, .xls, .xlsx, .pdf"
      )
  ```

- [ ] Rodar: `python -m pytest tests/test_detector.py -v` — esperar: **PASS**

- [ ] Commit: `git commit -m "feat: detector de formato (OFX, CSV, Excel, PDF) por extensão e conteúdo"`

---

## Task 3: Parser OFX

**Files:** `parsers/ofx_parser.py`, `tests/test_ofx_parser.py`

**Steps:**

- [ ] Criar `tests/test_ofx_parser.py`:
  ```python
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
  ```

- [ ] Rodar: `python -m pytest tests/test_ofx_parser.py -v` — esperar: **FAIL**

- [ ] Instalar dependências: `pip install ofxparse openpyxl pandas pdfplumber pytest --break-system-packages -q`

- [ ] Criar `parsers/ofx_parser.py`:
  ```python
  from decimal import Decimal
  from datetime import date
  from typing import List
  from ofxparse import OfxParser
  from models.transaction import Transaction

  def parse_ofx(filepath: str) -> List[Transaction]:
      """Lê arquivo OFX/OFC e retorna lista de Transaction."""
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
  ```

- [ ] Rodar: `python -m pytest tests/test_ofx_parser.py -v` — esperar: **PASS**

- [ ] Commit: `git commit -m "feat: parser OFX/OFC usando ofxparse, retorna lista de Transaction ordenada por data"`

---

## Task 4: Parser CSV (Nubank, Inter, Itaú, Bradesco)

**Files:** `parsers/csv_parser.py`, `tests/fixtures/inter.csv`, `tests/fixtures/itau.csv`, `tests/fixtures/bradesco.csv`, `tests/test_csv_parser.py`

**Steps:**

- [ ] Criar `tests/fixtures/inter.csv`:
  ```
  Data;Histórico;Crédito;Débito;Saldo
  15/01/2024;PIX RECEBIDO - JOAO;1000,00;;5000,00
  16/01/2024;BOLETO CONCESSIONARIA;;200,00;4800,00
  ```

- [ ] Criar `tests/fixtures/itau.csv`:
  ```
  Data;Histórico;Docto.;Crédito (R$);Débito (R$);Saldo (R$)
  15/01/2024;PIX RECEBIDO - JOAO;;1000,00;;5000,00
  16/01/2024;COMPRA SUPERMERCADO;;;150,00;4850,00
  ```

- [ ] Criar `tests/fixtures/bradesco.csv`:
  ```
  Data;Histórico;Valor;Saldo
  15/01/2024;TED RECEBIDA;1000,00;5000,00
  16/01/2024;DÉBITO EM CONTA;-200,00;4800,00
  ```

- [ ] Criar `tests/test_csv_parser.py`:
  ```python
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
  ```

- [ ] Rodar: `python -m pytest tests/test_csv_parser.py -v` — esperar: **FAIL**

- [ ] Criar `parsers/csv_parser.py`:
  ```python
  import pandas as pd
  from decimal import Decimal
  from datetime import date
  from typing import List, Optional
  from models.transaction import Transaction

  # Perfis de bancos: (colunas_necessárias, separador, encoding, banco_id)
  _PERFIS = [
      {
          "banco": "nubank",
          "sep": ",",
          "encoding": "utf-8",
          "colunas": {"data": "Data", "descricao": "Descrição", "valor": "Valor"},
          "detectar_por": ["Descrição", "Valor"],
      },
      {
          "banco": "inter",
          "sep": ";",
          "encoding": "utf-8",
          "colunas": {"data": "Data", "descricao": "Histórico",
                      "credito": "Crédito", "debito": "Débito", "saldo": "Saldo"},
          "detectar_por": ["Crédito", "Débito"],
      },
      {
          "banco": "itau",
          "sep": ";",
          "encoding": "latin-1",
          "colunas": {"data": "Data", "descricao": "Histórico",
                      "credito": "Crédito (R$)", "debito": "Débito (R$)", "saldo": "Saldo (R$)"},
          "detectar_por": ["Crédito (R$)", "Débito (R$)"],
      },
      {
          "banco": "bradesco",
          "sep": ";",
          "encoding": "latin-1",
          "colunas": {"data": "Data", "descricao": "Histórico",
                      "valor": "Valor", "saldo": "Saldo"},
          "detectar_por": ["Histórico", "Valor", "Saldo"],
      },
  ]

  def _parse_decimal(valor_str: str) -> Decimal:
      if pd.isna(valor_str) or str(valor_str).strip() == "":
          return Decimal("0")
      s = str(valor_str).strip().replace(".", "").replace(",", ".")
      return Decimal(s)

  def _parse_date(valor_str: str) -> date:
      for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y"):
          try:
              from datetime import datetime
              return datetime.strptime(str(valor_str).strip(), fmt).date()
          except ValueError:
              continue
      raise ValueError(f"Data não reconhecida: {valor_str}")

  def _detectar_perfil(df: pd.DataFrame) -> Optional[dict]:
      colunas = set(df.columns.tolist())
      for perfil in _PERFIS:
          if all(c in colunas for c in perfil["detectar_por"]):
              return perfil
      return None

  def parse_csv(filepath: str) -> List[Transaction]:
      """Detecta o banco pelo header do CSV e normaliza as transações."""
      for perfil in _PERFIS:
          try:
              df = pd.read_csv(
                  filepath,
                  sep=perfil["sep"],
                  encoding=perfil["encoding"],
                  dtype=str,
              ).dropna(how="all")
              df.columns = df.columns.str.strip()
              perfil_detectado = _detectar_perfil(df)
              if perfil_detectado is None:
                  continue
              return _normalizar(df, perfil_detectado)
          except Exception:
              continue

      raise ValueError(f"Formato de CSV não reconhecido: {filepath}")

  def _normalizar(df: pd.DataFrame, perfil: dict) -> List[Transaction]:
      col = perfil["colunas"]
      banco = perfil["banco"]
      transacoes = []

      for _, row in df.iterrows():
          data = _parse_date(row[col["data"]])
          descricao = str(row.get(col.get("descricao", ""), "")).strip()

          if "valor" in col:
              valor = _parse_decimal(row[col["valor"]])
          else:
              credito = _parse_decimal(row.get(col.get("credito", ""), ""))
              debito = _parse_decimal(row.get(col.get("debito", ""), ""))
              valor = credito - debito

          tipo = "Crédito" if valor >= 0 else "Débito"
          saldo_str = row.get(col.get("saldo", ""), None)
          saldo = _parse_decimal(saldo_str) if saldo_str and not pd.isna(saldo_str) else None

          transacoes.append(Transaction(
              data=data,
              descricao=descricao,
              valor=valor,
              tipo=tipo,
              saldo=saldo,
              banco_detectado=banco,
          ))

      return sorted(transacoes, key=lambda t: t.data)
  ```

- [ ] Rodar: `python -m pytest tests/test_csv_parser.py -v` — esperar: **PASS**

- [ ] Commit: `git commit -m "feat: parser CSV multi-banco (Nubank, Inter, Itaú, Bradesco) com detecção automática por header"`

---

## Task 5: Parser Excel de Entrada

**Files:** `parsers/excel_parser.py`, `tests/test_excel_parser.py`

**Steps:**

- [ ] Criar `tests/test_excel_parser.py`:
  ```python
  import pytest
  import openpyxl
  from decimal import Decimal
  from datetime import date
  from parsers.excel_parser import parse_excel

  def test_parse_excel_basico(tmp_path):
      wb = openpyxl.Workbook()
      ws = wb.active
      ws.append(["Data", "Histórico", "Valor", "Saldo"])
      ws.append(["15/01/2024", "PIX RECEBIDO", "1000,00", "5000,00"])
      ws.append(["16/01/2024", "COMPRA", "-150,00", "4850,00"])
      path = str(tmp_path / "extrato.xlsx")
      wb.save(path)

      t = parse_excel(path)
      assert len(t) == 2
      assert t[0].valor == Decimal("1000.00")
      assert t[0].tipo == "Crédito"
      assert t[1].valor == Decimal("-150.00")
      assert t[1].tipo == "Débito"

  def test_parse_excel_datas(tmp_path):
      wb = openpyxl.Workbook()
      ws = wb.active
      ws.append(["Data", "Histórico", "Valor"])
      ws.append(["2024-01-15", "PIX", "500,00"])
      path = str(tmp_path / "extrato.xlsx")
      wb.save(path)

      t = parse_excel(path)
      assert t[0].data == date(2024, 1, 15)
  ```

- [ ] Rodar: `python -m pytest tests/test_excel_parser.py -v` — esperar: **FAIL**

- [ ] Criar `parsers/excel_parser.py`:
  ```python
  import pandas as pd
  from decimal import Decimal
  from typing import List
  from models.transaction import Transaction
  from parsers.csv_parser import _parse_decimal, _parse_date

  def parse_excel(filepath: str) -> List[Transaction]:
      """Lê planilha Excel (xls/xlsx) de extrato bancário."""
      df = pd.read_excel(filepath, dtype=str).dropna(how="all")
      df.columns = df.columns.str.strip()

      # Mapeia nomes de colunas comuns
      col_map = {}
      for c in df.columns:
          cl = c.lower()
          if "data" in cl:
              col_map["data"] = c
          elif "hist" in cl or "descri" in cl or "lançam" in cl or "lancam" in cl:
              col_map["descricao"] = c
          elif "valor" in cl and "saldo" not in cl:
              col_map["valor"] = c
          elif "saldo" in cl:
              col_map["saldo"] = c
          elif "crédit" in cl or "credit" in cl:
              col_map["credito"] = c
          elif "débit" in cl or "debit" in cl:
              col_map["debito"] = c

      if "data" not in col_map:
          raise ValueError("Coluna de data não encontrada na planilha")

      transacoes = []
      for _, row in df.iterrows():
          data = _parse_date(row[col_map["data"]])
          descricao = str(row.get(col_map.get("descricao", ""), "")).strip()

          if "valor" in col_map:
              valor = _parse_decimal(row[col_map["valor"]])
          elif "credito" in col_map and "debito" in col_map:
              valor = _parse_decimal(row[col_map["credito"]]) - _parse_decimal(row[col_map["debito"]])
          else:
              valor = Decimal("0")

          tipo = "Crédito" if valor >= 0 else "Débito"
          saldo = _parse_decimal(row[col_map["saldo"]]) if "saldo" in col_map else None

          transacoes.append(Transaction(
              data=data,
              descricao=descricao,
              valor=valor,
              tipo=tipo,
              saldo=saldo,
          ))

      return sorted(transacoes, key=lambda t: t.data)
  ```

- [ ] Rodar: `python -m pytest tests/test_excel_parser.py -v` — esperar: **PASS**

- [ ] Commit: `git commit -m "feat: parser Excel de entrada com mapeamento automático de colunas"`

---

## Task 6: Parser PDF

**Files:** `parsers/pdf_parser.py`, `tests/test_pdf_parser.py`

**Steps:**

- [ ] Criar `tests/test_pdf_parser.py`:
  ```python
  import pytest
  from unittest.mock import patch, MagicMock
  from decimal import Decimal
  from datetime import date
  from parsers.pdf_parser import parse_pdf, _extrair_transacoes_do_texto

  TEXTO_NUBANK_PDF = """
  15/01/2024  PIX recebido João                   1.000,00
  16/01/2024  iFood*Restaurante                     -45,90
  17/01/2024  Boleto energia                        -120,00
  """

  def test_extrair_transacoes_de_texto():
      transacoes = _extrair_transacoes_do_texto(TEXTO_NUBANK_PDF)
      assert len(transacoes) == 3

  def test_extrair_credito():
      transacoes = _extrair_transacoes_do_texto(TEXTO_NUBANK_PDF)
      credito = next(t for t in transacoes if t.tipo == "Crédito")
      assert credito.valor == Decimal("1000.00")
      assert credito.data == date(2024, 1, 15)

  def test_extrair_debito():
      transacoes = _extrair_transacoes_do_texto(TEXTO_NUBANK_PDF)
      debitos = [t for t in transacoes if t.tipo == "Débito"]
      assert len(debitos) == 2

  def test_parse_pdf_usa_pdfplumber(tmp_path):
      with patch("parsers.pdf_parser.pdfplumber") as mock_plumber:
          mock_page = MagicMock()
          mock_page.extract_text.return_value = TEXTO_NUBANK_PDF
          mock_pdf = MagicMock()
          mock_pdf.__enter__ = MagicMock(return_value=mock_pdf)
          mock_pdf.__exit__ = MagicMock(return_value=False)
          mock_pdf.pages = [mock_page]
          mock_plumber.open.return_value = mock_pdf

          f = tmp_path / "extrato.pdf"
          f.write_bytes(b"%PDF fake")
          result = parse_pdf(str(f))
          assert len(result) == 3
  ```

- [ ] Rodar: `python -m pytest tests/test_pdf_parser.py -v` — esperar: **FAIL**

- [ ] Criar `parsers/pdf_parser.py`:
  ```python
  import re
  import pdfplumber
  from decimal import Decimal, InvalidOperation
  from datetime import datetime
  from typing import List
  from models.transaction import Transaction

  # Padrão: DD/MM/AAAA + descrição + valor (com possível sinal)
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
          data_str, descricao, valor_str = m.group(1), m.group(2).strip(), m.group(3)
          try:
              data = datetime.strptime(data_str, "%d/%m/%Y").date()
              valor = _parse_valor_br(valor_str)
              tipo = "Crédito" if valor >= 0 else "Débito"
              transacoes.append(Transaction(
                  data=data,
                  descricao=descricao,
                  valor=valor,
                  tipo=tipo,
                  banco_detectado="pdf",
              ))
          except (ValueError, InvalidOperation):
              continue
      return sorted(transacoes, key=lambda t: t.data)

  def parse_pdf(filepath: str) -> List[Transaction]:
      """Extrai transações de PDF bancário usando pdfplumber."""
      texto_completo = []
      with pdfplumber.open(filepath) as pdf:
          for page in pdf.pages:
              t = page.extract_text()
              if t:
                  texto_completo.append(t)
      return _extrair_transacoes_do_texto("\n".join(texto_completo))
  ```

- [ ] Rodar: `python -m pytest tests/test_pdf_parser.py -v` — esperar: **PASS**

- [ ] Commit: `git commit -m "feat: parser PDF usando pdfplumber com regex para extrair data/descricao/valor"`

---

## Task 7: Gerador de Relatório Excel — Aba Extrato

**Files:** `report/__init__.py`, `report/excel_report.py`, `tests/test_excel_report.py` (parcial)

**Steps:**

- [ ] Criar `report/__init__.py` (vazio)

- [ ] Criar `tests/test_excel_report.py`:
  ```python
  import os
  import pytest
  import openpyxl
  from decimal import Decimal
  from datetime import date
  from models.transaction import Transaction
  from report.excel_report import gerar_relatorio

  TRANSACOES = [
      Transaction(date(2024, 1, 15), "PIX RECEBIDO - JOAO", Decimal("1000.00"), "Crédito", Decimal("5000.00")),
      Transaction(date(2024, 1, 16), "COMPRA SUPERMERCADO", Decimal("-150.00"), "Débito", Decimal("4850.00")),
      Transaction(date(2024, 1, 17), "SALÁRIO", Decimal("3000.00"), "Crédito", Decimal("7850.00")),
      Transaction(date(2024, 1, 20), "ALUGUEL", Decimal("-1200.00"), "Débito", Decimal("6650.00")),
  ]

  def test_gerar_relatorio_cria_arquivo(tmp_path):
      output = str(tmp_path / "relatorio.xlsx")
      gerar_relatorio(TRANSACOES, output)
      assert os.path.exists(output)

  def test_relatorio_tem_duas_abas(tmp_path):
      output = str(tmp_path / "relatorio.xlsx")
      gerar_relatorio(TRANSACOES, output)
      wb = openpyxl.load_workbook(output)
      assert "Extrato" in wb.sheetnames
      assert "Resumo" in wb.sheetnames

  def test_extrato_tem_cabecalho_correto(tmp_path):
      output = str(tmp_path / "relatorio.xlsx")
      gerar_relatorio(TRANSACOES, output)
      wb = openpyxl.load_workbook(output)
      ws = wb["Extrato"]
      cabecalho = [ws.cell(1, c).value for c in range(1, 6)]
      assert cabecalho == ["Data", "Descrição", "Débitos (R$)", "Créditos (R$)", "Saldo (R$)"]

  def test_extrato_tem_quantidade_correta_de_linhas(tmp_path):
      output = str(tmp_path / "relatorio.xlsx")
      gerar_relatorio(TRANSACOES, output)
      wb = openpyxl.load_workbook(output)
      ws = wb["Extrato"]
      # 1 cabeçalho + 4 transações + 1 total = 6
      assert ws.max_row == 6

  def test_celulas_debito_tem_fonte_vermelha(tmp_path):
      output = str(tmp_path / "relatorio.xlsx")
      gerar_relatorio(TRANSACOES, output)
      wb = openpyxl.load_workbook(output)
      ws = wb["Extrato"]
      # Linha 3 é o primeiro débito (linha 2 = primeira transação = crédito)
      debito_row = 3
      cor = ws.cell(debito_row, 3).font.color.rgb  # coluna Débitos
      assert cor == "FFFF0000"  # vermelho

  def test_resumo_total_creditos(tmp_path):
      output = str(tmp_path / "relatorio.xlsx")
      gerar_relatorio(TRANSACOES, output)
      wb = openpyxl.load_workbook(output)
      ws = wb["Resumo"]
      # Procura célula com valor 4000.00 (1000 + 3000)
      valores = [ws.cell(r, 2).value for r in range(1, ws.max_row + 1)]
      assert 4000.0 in valores or Decimal("4000.00") in [v for v in valores if v]
  ```

- [ ] Rodar: `python -m pytest tests/test_excel_report.py -v` — esperar: **FAIL**

- [ ] Criar `report/excel_report.py`:
  ```python
  from decimal import Decimal
  from datetime import date
  from typing import List, Optional
  import openpyxl
  from openpyxl.styles import (
      PatternFill, Font, Alignment, Border, Side, numbers
  )
  from openpyxl.utils import get_column_letter
  from models.transaction import Transaction

  # ── Cores ──────────────────────────────────────────────────────────────────
  COR_CABECALHO_FUNDO   = "1F3864"   # azul escuro
  COR_CABECALHO_FONTE   = "FFFFFF"   # branco
  COR_LINHA_PAR         = "EEF2FF"   # azul muito claro
  COR_LINHA_IMPAR       = "FFFFFF"   # branco
  COR_DEBITO            = "FF0000"   # vermelho
  COR_CREDITO           = "006100"   # verde escuro
  COR_TOTAL_FUNDO       = "D9E1F2"   # azul claro
  COR_TOTAL_FONTE       = "1F3864"   # azul escuro

  FMT_MOEDA = 'R$ #,##0.00;[RED]-R$ #,##0.00'
  FMT_DATA  = 'DD/MM/YYYY'

  def _borda_fina():
      lado = Side(style="thin", color="CCCCCC")
      return Border(left=lado, right=lado, top=lado, bottom=lado)

  def _fill(hex_color: str) -> PatternFill:
      return PatternFill("solid", fgColor=hex_color)

  def _ajustar_largura(ws):
      for col in ws.columns:
          max_len = 0
          col_letter = get_column_letter(col[0].column)
          for cell in col:
              if cell.value:
                  max_len = max(max_len, len(str(cell.value)))
          ws.column_dimensions[col_letter].width = min(max_len + 4, 50)

  def _escrever_aba_extrato(ws, transacoes: List[Transaction]):
      # Cabeçalho
      cabecalho = ["Data", "Descrição", "Débitos (R$)", "Créditos (R$)", "Saldo (R$)"]
      for c, titulo in enumerate(cabecalho, 1):
          cell = ws.cell(1, c, titulo)
          cell.fill = _fill(COR_CABECALHO_FUNDO)
          cell.font = Font(bold=True, color=COR_CABECALHO_FONTE, size=11)
          cell.alignment = Alignment(horizontal="center", vertical="center")
          cell.border = _borda_fina()
      ws.row_dimensions[1].height = 22

      # Dados
      for i, t in enumerate(transacoes):
          row = i + 2
          fundo = COR_LINHA_PAR if i % 2 == 0 else COR_LINHA_IMPAR

          debito  = abs(t.valor) if t.tipo == "Débito"  else None
          credito = t.valor      if t.tipo == "Crédito" else None

          valores = [
              t.data,
              t.descricao,
              float(debito)  if debito  is not None else None,
              float(credito) if credito is not None else None,
              float(t.saldo) if t.saldo is not None else None,
          ]
          for c, val in enumerate(valores, 1):
              cell = ws.cell(row, c, val)
              cell.fill = _fill(fundo)
              cell.border = _borda_fina()
              cell.alignment = Alignment(vertical="center")

              if c == 1:  # Data
                  cell.number_format = FMT_DATA
                  cell.alignment = Alignment(horizontal="center", vertical="center")
              elif c == 3 and val is not None:  # Débito
                  cell.number_format = FMT_MOEDA
                  cell.font = Font(color=COR_DEBITO)
                  cell.alignment = Alignment(horizontal="right", vertical="center")
              elif c in (4, 5) and val is not None:  # Crédito / Saldo
                  cell.number_format = FMT_MOEDA
                  cell.font = Font(color=COR_CREDITO if c == 4 else "000000")
                  cell.alignment = Alignment(horizontal="right", vertical="center")

      # Linha de totais
      total_row = len(transacoes) + 2
      total_debitos  = sum(abs(t.valor) for t in transacoes if t.tipo == "Débito")
      total_creditos = sum(t.valor for t in transacoes if t.tipo == "Crédito")

      for c in range(1, 6):
          cell = ws.cell(total_row, c)
          cell.fill = _fill(COR_TOTAL_FUNDO)
          cell.font = Font(bold=True, color=COR_TOTAL_FONTE)
          cell.border = _borda_fina()

      ws.cell(total_row, 2, "TOTAL").alignment = Alignment(horizontal="right")
      ws.cell(total_row, 3, float(total_debitos)).number_format  = FMT_MOEDA
      ws.cell(total_row, 4, float(total_creditos)).number_format = FMT_MOEDA

      _ajustar_largura(ws)
      ws.freeze_panes = "A2"

  def _escrever_aba_resumo(ws, transacoes: List[Transaction]):
      if not transacoes:
          ws.cell(1, 1, "Nenhuma transação encontrada")
          return

      datas = [t.data for t in transacoes]
      total_creditos = sum(t.valor for t in transacoes if t.tipo == "Crédito")
      total_debitos  = sum(abs(t.valor) for t in transacoes if t.tipo == "Débito")
      saldo_liquido  = total_creditos - total_debitos
      quantidade     = len(transacoes)

      def linha(r, label, valor, fmt=None):
          label_cell = ws.cell(r, 1, label)
          label_cell.font = Font(bold=True, size=11)
          label_cell.fill = _fill(COR_LINHA_PAR)
          label_cell.border = _borda_fina()
          val_cell = ws.cell(r, 2, valor)
          val_cell.border = _borda_fina()
          val_cell.alignment = Alignment(horizontal="right")
          if fmt:
              val_cell.number_format = fmt

      ws.cell(1, 1, "RESUMO DO PERÍODO").font = Font(
          bold=True, size=13, color=COR_CABECALHO_FONTE
      )
      ws.cell(1, 1).fill = _fill(COR_CABECALHO_FUNDO)
      ws.merge_cells("A1:B1")

      linha(3,  "Período inicial",          datas[0],               FMT_DATA)
      linha(4,  "Período final",             datas[-1],              FMT_DATA)
      linha(5,  "Total de transações",       quantidade)
      linha(7,  "Total de entradas (R$)",    float(total_creditos),  FMT_MOEDA)
      linha(8,  "Total de saídas (R$)",      float(total_debitos),   FMT_MOEDA)
      linha(9,  "Saldo líquido do período",  float(saldo_liquido),   FMT_MOEDA)

      ws.cell(9, 2).font = Font(
          bold=True,
          color=COR_CREDITO if saldo_liquido >= 0 else COR_DEBITO
      )

      # Tabela mensal (se mais de 1 mês)
      from collections import defaultdict
      meses: dict = defaultdict(lambda: {"credito": Decimal(0), "debito": Decimal(0)})
      for t in transacoes:
          chave = f"{t.data.year}-{t.data.month:02d}"
          if t.tipo == "Crédito":
              meses[chave]["credito"] += t.valor
          else:
              meses[chave]["debito"] += abs(t.valor)

      if len(meses) > 1:
          r = 12
          ws.cell(r, 1, "Resumo Mensal").font = Font(bold=True, size=11,
                                                       color=COR_CABECALHO_FONTE)
          ws.cell(r, 1).fill = _fill(COR_CABECALHO_FUNDO)
          ws.merge_cells(f"A{r}:D{r}")
          r += 1
          for col, titulo in enumerate(["Mês", "Entradas", "Saídas", "Saldo"], 1):
              c = ws.cell(r, col, titulo)
              c.font = Font(bold=True, color=COR_CABECALHO_FONTE)
              c.fill = _fill(COR_CABECALHO_FUNDO)
              c.border = _borda_fina()
          r += 1
          for mes, v in sorted(meses.items()):
              saldo_mes = v["credito"] - v["debito"]
              for col, val in enumerate([mes, float(v["credito"]),
                                         float(v["debito"]), float(saldo_mes)], 1):
                  cell = ws.cell(r, col, val)
                  cell.border = _borda_fina()
                  if col > 1:
                      cell.number_format = FMT_MOEDA
              r += 1

      _ajustar_largura(ws)

  def gerar_relatorio(transacoes: List[Transaction], output_path: str) -> str:
      """Gera relatório Excel com abas Extrato e Resumo."""
      wb = openpyxl.Workbook()
      ws_extrato = wb.active
      ws_extrato.title = "Extrato"
      ws_resumo = wb.create_sheet("Resumo")

      _escrever_aba_extrato(ws_extrato, transacoes)
      _escrever_aba_resumo(ws_resumo, transacoes)

      wb.save(output_path)
      return output_path
  ```

- [ ] Rodar: `python -m pytest tests/test_excel_report.py -v` — esperar: **PASS**

- [ ] Commit: `git commit -m "feat: gerador de relatório Excel com abas Extrato (cores débito/crédito) e Resumo (totais + mensal)"`

---

## Task 8: Orquestrador Principal

**Files:** `main.py`, `tests/test_main.py` (integração)

**Steps:**

- [ ] Criar `tests/test_main.py`:
  ```python
  import os
  import pytest
  import openpyxl
  from main import processar_extrato

  FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")

  def test_processar_extrato_ofx(tmp_path):
      input_path  = os.path.join(FIXTURES, "sample.ofx")
      output_path = str(tmp_path / "relatorio.xlsx")
      resultado   = processar_extrato(input_path, output_path)
      assert os.path.exists(resultado)
      wb = openpyxl.load_workbook(resultado)
      assert "Extrato" in wb.sheetnames
      assert "Resumo" in wb.sheetnames

  def test_processar_extrato_csv_nubank(tmp_path):
      input_path  = os.path.join(FIXTURES, "nubank.csv")
      output_path = str(tmp_path / "relatorio.xlsx")
      resultado   = processar_extrato(input_path, output_path)
      assert os.path.exists(resultado)
      wb = openpyxl.load_workbook(resultado)
      ws = wb["Extrato"]
      assert ws.max_row >= 3  # cabeçalho + 2 transações + total

  def test_processar_extrato_formato_invalido(tmp_path):
      f = tmp_path / "arquivo.zip"
      f.write_bytes(b"PK")
      with pytest.raises(ValueError):
          processar_extrato(str(f), str(tmp_path / "out.xlsx"))
  ```

- [ ] Rodar: `python -m pytest tests/test_main.py -v` — esperar: **FAIL**

- [ ] Criar `main.py`:
  ```python
  import os
  import sys
  import subprocess
  from datetime import datetime
  from typing import Optional

  def _selecionar_arquivo() -> Optional[str]:
      """Abre diálogo de seleção de arquivo nativo do Windows via tkinter."""
      try:
          import tkinter as tk
          from tkinter import filedialog, messagebox
          root = tk.Tk()
          root.withdraw()
          root.attributes("-topmost", True)
          path = filedialog.askopenfilename(
              title="Selecionar extrato bancário",
              filetypes=[
                  ("Extratos bancários", "*.ofx *.ofc *.csv *.xlsx *.xls *.pdf"),
                  ("OFX / OFC", "*.ofx *.ofc"),
                  ("CSV", "*.csv"),
                  ("Excel", "*.xlsx *.xls"),
                  ("PDF", "*.pdf"),
                  ("Todos os arquivos", "*.*"),
              ]
          )
          root.destroy()
          return path if path else None
      except Exception as e:
          print(f"Erro ao abrir seletor de arquivo: {e}")
          return None

  def _gerar_nome_saida(input_path: str) -> str:
      pasta = os.path.dirname(os.path.abspath(input_path))
      timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
      nome = f"Conciliacao_{timestamp}.xlsx"
      return os.path.join(pasta, nome)

  def processar_extrato(input_path: str, output_path: Optional[str] = None) -> str:
      """Processa um arquivo de extrato e gera relatório Excel."""
      from parsers.detector import detect_format
      from parsers.ofx_parser import parse_ofx
      from parsers.csv_parser import parse_csv
      from parsers.excel_parser import parse_excel
      from parsers.pdf_parser import parse_pdf
      from report.excel_report import gerar_relatorio

      formato = detect_format(input_path)

      parsers = {
          "ofx":   parse_ofx,
          "csv":   parse_csv,
          "excel": parse_excel,
          "pdf":   parse_pdf,
      }

      print(f"Formato detectado: {formato.upper()}")
      transacoes = parsers[formato](input_path)
      print(f"Transações encontradas: {len(transacoes)}")

      if output_path is None:
          output_path = _gerar_nome_saida(input_path)

      gerar_relatorio(transacoes, output_path)
      print(f"Relatório gerado: {output_path}")
      return output_path

  def main():
      if len(sys.argv) > 1:
          input_path = sys.argv[1]
      else:
          input_path = _selecionar_arquivo()
          if not input_path:
              print("Nenhum arquivo selecionado. Encerrando.")
              sys.exit(0)

      try:
          output = processar_extrato(input_path)
          # Abrir o arquivo Excel gerado automaticamente
          if sys.platform == "win32":
              os.startfile(output)
          else:
              subprocess.run(["xdg-open", output])
      except Exception as e:
          print(f"\nErro ao processar o extrato: {e}")
          input("\nPressione ENTER para fechar...")
          sys.exit(1)

  if __name__ == "__main__":
      main()
  ```

- [ ] Rodar: `python -m pytest tests/test_main.py -v` — esperar: **PASS**

- [ ] Rodar todos os testes: `python -m pytest tests/ -v --tb=short` — esperar: **PASS** em todos

- [ ] Commit: `git commit -m "feat: orquestrador main.py com seletor de arquivo tkinter e abertura automática do relatório"`

---

## Task 9: Arquivo .bat de Lançamento

**Files:** `conciliar.bat`

> Esta task não usa TDD — a verificação é manual, pois envolve interação com o Windows.

**Steps:**

- [ ] Criar `conciliar.bat`:
  ```batch
  @echo off
  chcp 65001 >nul
  title Conciliação Bancária

  echo ============================================
  echo    CONCILIAÇÃO BANCÁRIA AUTOMÁTICA
  echo ============================================
  echo.

  :: Verifica se Python está instalado
  python --version >nul 2>&1
  if errorlevel 1 (
      echo ERRO: Python não encontrado.
      echo.
      echo Por favor, instale o Python 3.8+ em:
      echo   https://www.python.org/downloads/
      echo.
      echo Marque a opção "Add Python to PATH" durante a instalação.
      pause
      exit /b 1
  )

  :: Vai para o diretório do .bat
  cd /d "%~dp0"

  :: Instala dependências (silencioso se já instalado)
  echo Verificando dependências...
  python -m pip install -r requirements.txt --quiet --disable-pip-version-check
  if errorlevel 1 (
      echo AVISO: Não foi possível verificar as dependências.
      echo Tentando continuar mesmo assim...
      echo.
  )

  :: Executa o script principal
  echo Abrindo seletor de arquivo...
  echo.
  python main.py

  :: Se terminar com erro, mostra mensagem
  if errorlevel 1 (
      echo.
      echo O programa encerrou com um erro.
      pause
  )
  ```

- [ ] Verificação manual:
  - Dar duplo-clique em `conciliar.bat`
  - Verificar que o seletor de arquivo abre
  - Selecionar um dos arquivos em `tests/fixtures/`
  - Verificar que o Excel é gerado na mesma pasta do arquivo selecionado
  - Verificar que o Excel abre automaticamente

- [ ] Commit: `git commit -m "feat: conciliar.bat com verificação de Python, instalação automática de deps e lançamento do seletor"`

---

## Task 10: Perfis CSV Adicionais (BB, Santander, C6, Sicoob)

**Files:** `parsers/csv_parser.py`, `tests/fixtures/bb.csv`, `tests/fixtures/santander.csv`, `tests/test_csv_parser.py` (adição)

**Steps:**

- [ ] Criar `tests/fixtures/bb.csv`:
  ```
  Data;Lançamento;Documento;Crédito;Débito;Saldo
  15/01/2024;PIX RECEBIDO - JOAO;;1.000,00;;5.000,00
  16/01/2024;PAGAMENTO BOLETO;;; 200,00;4.800,00
  ```

- [ ] Criar `tests/fixtures/santander.csv`:
  ```
  Data;Descrição;Valor;Saldo
  15/01/2024;Crédito PIX;1.000,00;5.000,00
  16/01/2024;Débito PIX;-200,00;4.800,00
  ```

- [ ] Adicionar ao `tests/test_csv_parser.py`:
  ```python
  def test_parse_bb():
      path = os.path.join(FIXTURES, "bb.csv")
      t = parse_csv(path)
      assert len(t) == 2
      assert t[0].banco_detectado == "bb"
      assert t[0].valor == Decimal("1000.00")
      assert t[1].valor == Decimal("-200.00")

  def test_parse_santander():
      path = os.path.join(FIXTURES, "santander.csv")
      t = parse_csv(path)
      assert len(t) == 2
      assert t[0].banco_detectado == "santander"
      assert t[0].valor == Decimal("1000.00")
  ```

- [ ] Rodar: `python -m pytest tests/test_csv_parser.py::test_parse_bb tests/test_csv_parser.py::test_parse_santander -v` — esperar: **FAIL**

- [ ] Adicionar perfis ao início de `_PERFIS` em `parsers/csv_parser.py`:
  ```python
  {
      "banco": "bb",
      "sep": ";",
      "encoding": "latin-1",
      "colunas": {"data": "Data", "descricao": "Lançamento",
                  "credito": "Crédito", "debito": "Débito", "saldo": "Saldo"},
      "detectar_por": ["Lançamento", "Crédito", "Débito"],
  },
  {
      "banco": "santander",
      "sep": ";",
      "encoding": "latin-1",
      "colunas": {"data": "Data", "descricao": "Descrição",
                  "valor": "Valor", "saldo": "Saldo"},
      "detectar_por": ["Descrição", "Valor", "Saldo"],
  },
  ```

- [ ] Rodar: `python -m pytest tests/test_csv_parser.py -v` — esperar: **PASS** em todos

- [ ] Rodar suite completa: `python -m pytest tests/ -v` — esperar: **PASS** em todos

- [ ] Commit: `git commit -m "feat: adiciona perfis CSV para Banco do Brasil e Santander"`

---

## Verificação Final

- [ ] Rodar: `python -m pytest tests/ -v --tb=short` — esperar todos **PASS**
- [ ] Testar manualmente com `conciliar.bat` + arquivo OFX real
- [ ] Confirmar que o relatório Excel abre com formatação correta (cores, totais, aba Resumo)
- [ ] Confirmar que o arquivo é salvo na mesma pasta do extrato de entrada

---

## Ordem de Execução Recomendada

```
Task 1 → Task 2 → Task 3 → Task 4 → Task 5 → Task 6 → Task 7 → Task 8 → Task 9 → Task 10
```

Cada task é independente das seguintes, mas depende das anteriores.
