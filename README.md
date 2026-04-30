# Conciliação Bancária Automática

Ferramenta Python para Windows que processa extratos bancários e gera relatórios Excel visuais com análise financeira, categorização de gastos e alertas inteligentes.

## Como usar

1. Dê **duplo-clique em `conciliar.bat`**
2. Escolha o **modo de relatório**: Simples ou Detalhado
3. **Selecione o arquivo** de extrato bancário
4. O Excel é gerado automaticamente na mesma pasta do extrato

> Na primeira execução o `.bat` instala as dependências automaticamente.

---

## Modos de Relatório

| Modo | Descrição | Abas |
|------|-----------|------|
| **Simples** | Visão executiva rápida — ideal para apresentar | Resumo (saúde financeira + categorias) |
| **Detalhado** | Análise completa — para uso pessoal | Extrato + Resumo + Evolução Mensal |

---

## Indicadores de Saúde Financeira

- **Banner verde**: gastos dentro das receitas — tudo OK
- **Banner vermelho**: gastos superam receitas — atenção necessária
- **% gasto da renda**: mostra quanto da renda foi consumido no período
- **Alertas por categoria** (amarelo): quando uma categoria ultrapassa o limite configurado

---

## Categorias Detectadas Automaticamente

| Categoria (débito) | Exemplos de transação |
|--------------------|----------------------|
| Salário | CREDITO DE SALARIO, PAGAMENTO DE SALARIO |
| PIX Enviado | PIX ENVIADO, PIX TRANSFERENCIA |
| Cartão de Crédito | PAGAMENTO CARTAO, FATURA CARTAO |
| Débito Automático | DEBITO AUTOM, MENSALIDADE, DEBITO VISA |
| Boleto | BOLETO, PAGAMENTO BOLETO, CONCESSIONARIA |
| Rendimentos | REMUNERACAO APLICACAO, RENDIMENTO, CDB |
| Transferência | TED ENVIADA, DOC ENVIADO |

---

## Configurando Limites de Gastos

Edite o arquivo `config/limites.txt` para definir o percentual máximo de cada categoria em relação à renda:

```ini
# Limite de 30% da renda para PIX Enviado
PIX Enviado = 30

# Limite de 25% para Cartão de Crédito
Cartao de Credito = 25
```

Quando uma categoria ultrapassar o limite, o relatório destaca em amarelo com o status **ALTO**.

---

## Formatos de Extrato Aceitos

| Formato | Extensão | Bancos |
|---------|----------|--------|
| OFX / OFC | `.ofx`, `.ofc` | Itaú, Bradesco, BB, Santander, CEF, Nubank, Inter, BTG, Sicredi |
| CSV | `.csv` | Nubank, Inter, Itaú, Bradesco, BB, Santander |
| Excel | `.xlsx`, `.xls` | Qualquer banco com exportação em planilha |
| PDF | `.pdf` | Melhor esforço — qualidade varia por banco |

---

## Suporte a Períodos Longos

O sistema suporta extratos de qualquer período. Quando o arquivo contém mais de um mês:
- O relatório gera automaticamente a tabela de **Evolução Mensal**
- Cada mês aparece com totais de entradas, saídas e saldo
- Meses com saldo negativo são destacados em vermelho

---

## Requisitos

- **Python 3.8+** ([download](https://www.python.org/downloads/))
  - Marque **"Add Python to PATH"** durante a instalação
- As demais dependências são instaladas automaticamente pelo `.bat`

### Dependências Python
```
ofxparse >= 0.21
pandas >= 1.5.0
openpyxl >= 3.0.0
pdfplumber >= 0.7.0
```

---

## Estrutura do Projeto

```
conciliacao-bancaria/
├── conciliar.bat              # Launcher principal (Windows)
├── main.py                    # Orquestrador + seleção de modo
├── requirements.txt
├── config/
│   └── limites.txt            # Limites de gastos por categoria (editável)
├── models/
│   └── transaction.py         # Modelo de dados
├── parsers/
│   ├── detector.py            # Detecção automática de formato
│   ├── ofx_parser.py          # Parser OFX/OFC
│   ├── csv_parser.py          # Parser CSV (multi-banco)
│   ├── excel_parser.py        # Parser Excel de entrada
│   ├── pdf_parser.py          # Parser PDF
│   └── categorizer.py         # Categorização automática de transações
├── report/
│   └── excel_report.py        # Gerador de relatório Excel
└── tests/                     # 35 testes automatizados
```

---

## Rodando os Testes

```bash
pip install pytest --prefer-binary
python -m pytest tests/ -v
```

---

## Licença

MIT — use, modifique e distribua livremente.
