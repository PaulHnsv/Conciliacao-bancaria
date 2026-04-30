# Design Spec — Conciliação Bancária Automática
**Data:** 2026-04-29  
**Status:** Aprovado para implementação

---

## Objetivo

Criar uma ferramenta de linha de comando (Python + .bat) que:
1. Aceite um arquivo de extrato bancário via seletor de arquivo nativo do Windows
2. Detecte automaticamente o formato (OFX, CSV, XLSX, PDF)
3. Normalize os dados para uma estrutura comum
4. Gere um relatório Excel visualmente formatado, pronto para apresentação

Interação mínima com o usuário: apenas selecionar o arquivo de entrada.

---

## Escopo

### Bancos suportados (via OFX — padrão)
- Itaú, Bradesco, Banco do Brasil, Santander, Caixa Econômica Federal, Nubank, Inter, Sicredi, Banrisul, BTG Pactual

### Formatos de entrada suportados
| Formato | Extensão | Biblioteca | Observação |
|---------|----------|-----------|-----------|
| OFX / OFC | `.ofx`, `.ofc` | `ofxparse` | Padrão de todos os bancos |
| CSV | `.csv` | `pandas` | Detecção automática de separador e encoding |
| Excel | `.xls`, `.xlsx` | `openpyxl` / `pandas` | |
| PDF | `.pdf` | `pdfplumber` | Melhor esforço — qualidade depende do banco |

---

## Arquitetura do Sistema

```
conciliacao/
├── conciliar.bat          # Ponto de entrada — instala deps, abre seletor, chama main.py
├── main.py                # Orquestrador principal
├── parsers/
│   ├── __init__.py
│   ├── detector.py        # Detecta formato pelo conteúdo/extensão
│   ├── ofx_parser.py      # Parser OFX/OFC
│   ├── csv_parser.py      # Parser CSV com detecção de banco
│   ├── excel_parser.py    # Parser XLS/XLSX
│   └── pdf_parser.py      # Parser PDF (pdfplumber)
├── models/
│   └── transaction.py     # Modelo de dados: data, descrição, valor, tipo, saldo
├── report/
│   └── excel_report.py    # Gerador do relatório Excel formatado
└── requirements.txt       # ofxparse, pandas, openpyxl, pdfplumber
```

---

## Fluxo de Execução

```
[Duplo-clique em conciliar.bat]
        ↓
[Verifica Python instalado]
        ↓
[pip install -r requirements.txt --quiet]
        ↓
[Abre seletor de arquivo nativo (tkinter)]
        ↓
[detector.py → identifica formato]
        ↓
[parser correspondente → lista de Transaction]
        ↓
[excel_report.py → gera relatório]
        ↓
[Abre o Excel gerado automaticamente]
        ↓
[Janela do .bat fecha]
```

---

## Modelo de Dados

```python
@dataclass
class Transaction:
    data: date
    descricao: str
    valor: Decimal       # positivo = crédito, negativo = débito
    tipo: str            # "Crédito" | "Débito"
    saldo: Decimal | None
    banco_detectado: str | None
```

---

## Relatório Excel — Estrutura

### Aba 1: "Extrato"
- Tabela com todas as transações
- Colunas: Data | Descrição | Débitos | Créditos | Saldo
- Débitos em vermelho, créditos em verde
- Linha de totais no rodapé

### Aba 2: "Resumo"
- Total de entradas (créditos)
- Total de saídas (débitos)
- Saldo líquido do período
- Quantidade de transações
- Período (data inicial → data final)
- Mini tabela mensal (se o extrato cobrir mais de 1 mês)

### Formatação
- Cabeçalhos com fundo azul escuro, texto branco, negrito
- Linhas alternadas em cinza claro / branco
- Valores monetários formatados em R$
- Datas no formato DD/MM/AAAA
- Largura de colunas ajustada automaticamente
- Bordas nas células

---

## Detecção Automática de Formato

O `detector.py` usa a seguinte lógica em cascata:

1. **Extensão do arquivo:** `.ofx`/`.ofc` → OFX; `.pdf` → PDF; `.xls`/`.xlsx` → Excel
2. **Conteúdo do arquivo (para `.csv` e ambíguos):** lê os primeiros bytes e procura por `<OFX>`, `OFXHEADER`, `%PDF`, etc.
3. **Fallback:** tenta CSV com detecção de separador via `csv.Sniffer`

---

## Detecção de Banco no CSV

Cada banco exporta CSV com estrutura diferente. O `csv_parser.py` detecta o banco pelo:
- Nome das colunas (header)
- Padrão da primeira linha
- Encoding e separador típico

Bancos mapeados inicialmente: Nubank, Inter, Itaú, Bradesco, BB, Santander, C6 Bank.

---

## Requisitos Técnicos

- **Python:** 3.8+
- **OS:** Windows 10/11 (o .bat usa PowerShell para verificações)
- **Dependências:** `ofxparse`, `pandas`, `openpyxl`, `pdfplumber`
- **Sem servidor, sem internet após instalação inicial**

---

## Arquivo de Saída

- Nome: `Conciliacao_AAAA-MM-DD_HH-MM.xlsx`
- Salvo na mesma pasta do arquivo de entrada
- Aberto automaticamente após geração

---

## Decisões de Design

| Decisão | Escolha | Motivo |
|---------|---------|--------|
| Interface | `.bat` + `tkinter` | Zero instalação extra, nativo no Windows |
| Saída | Excel `.xlsx` | Mais familiar para o usuário final, editável |
| OFX como prioritário | Sim | Padrão oficial dos bancos brasileiros, mais confiável |
| PDF como fallback | Sim, best-effort | Qualidade varia por banco, mas melhor que nada |
| Seletor de arquivo | `tkinter.filedialog` | Nativo, sem dependência extra |
