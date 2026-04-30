from collections import defaultdict
from decimal import Decimal
from typing import List
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from models.transaction import Transaction

COR_CAB_FUNDO  = "1F3864"
COR_CAB_FONTE  = "FFFFFF"
COR_PAR        = "EEF2FF"
COR_IMPAR      = "FFFFFF"
COR_DEBITO     = "FF0000"
COR_CREDITO    = "006100"
COR_TOT_FUNDO  = "D9E1F2"
COR_TOT_FONTE  = "1F3864"

FMT_MOEDA = 'R$ #,##0.00;[RED]-R$ #,##0.00'
FMT_DATA  = 'DD/MM/YYYY'


def _borda():
    l = Side(style="thin", color="CCCCCC")
    return Border(left=l, right=l, top=l, bottom=l)


def _fill(cor):
    return PatternFill("solid", fgColor=cor)


def _ajustar_largura(ws):
    for col in ws.columns:
        maxlen = max((len(str(c.value)) for c in col if c.value), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(maxlen + 4, 50)


def _aba_extrato(ws, transacoes: List[Transaction]):
    cabecalho = ["Data", "Descricao", "Debitos (R$)", "Creditos (R$)", "Saldo (R$)"]
    for c, titulo in enumerate(cabecalho, 1):
        cell = ws.cell(1, c, titulo)
        cell.fill = _fill(COR_CAB_FUNDO)
        cell.font = Font(bold=True, color=COR_CAB_FONTE, size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = _borda()
    ws.row_dimensions[1].height = 22

    for i, t in enumerate(transacoes):
        row = i + 2
        fundo = COR_PAR if i % 2 == 0 else COR_IMPAR
        debito  = float(abs(t.valor)) if t.tipo == "Debito"  else None
        credito = float(t.valor)      if t.tipo == "Credito" else None
        saldo   = float(t.saldo)      if t.saldo is not None else None
        vals = [t.data, t.descricao, debito, credito, saldo]
        for c, val in enumerate(vals, 1):
            cell = ws.cell(row, c, val)
            cell.fill = _fill(fundo)
            cell.border = _borda()
            cell.alignment = Alignment(vertical="center")
            if c == 1:
                cell.number_format = FMT_DATA
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c == 3 and val is not None:
                cell.number_format = FMT_MOEDA
                cell.font = Font(color=COR_DEBITO)
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c in (4, 5) and val is not None:
                cell.number_format = FMT_MOEDA
                cell.font = Font(color=COR_CREDITO if c == 4 else "000000")
                cell.alignment = Alignment(horizontal="right", vertical="center")

    total_row = len(transacoes) + 2
    total_deb = sum(abs(t.valor) for t in transacoes if t.tipo == "Debito")
    total_cred = sum(t.valor for t in transacoes if t.tipo == "Credito")
    for c in range(1, 6):
        cell = ws.cell(total_row, c)
        cell.fill = _fill(COR_TOT_FUNDO)
        cell.font = Font(bold=True, color=COR_TOT_FONTE)
        cell.border = _borda()
    ws.cell(total_row, 2, "TOTAL").alignment = Alignment(horizontal="right")
    c3 = ws.cell(total_row, 3, float(total_deb))
    c3.number_format = FMT_MOEDA
    c3.font = Font(bold=True, color=COR_TOT_FONTE)
    c4 = ws.cell(total_row, 4, float(total_cred))
    c4.number_format = FMT_MOEDA
    c4.font = Font(bold=True, color=COR_TOT_FONTE)

    _ajustar_largura(ws)
    ws.freeze_panes = "A2"


def _aba_resumo(ws, transacoes: List[Transaction]):
    if not transacoes:
        ws.cell(1, 1, "Nenhuma transacao encontrada")
        return
    datas = [t.data for t in transacoes]
    tot_cred = sum(t.valor for t in transacoes if t.tipo == "Credito")
    tot_deb  = sum(abs(t.valor) for t in transacoes if t.tipo == "Debito")
    saldo    = tot_cred - tot_deb
    qtd      = len(transacoes)

    ws.cell(1, 1, "RESUMO DO PERIODO")
    ws.cell(1, 1).fill = _fill(COR_CAB_FUNDO)
    ws.cell(1, 1).font = Font(bold=True, size=13, color=COR_CAB_FONTE)
    ws.merge_cells("A1:B1")

    def linha(r, label, valor, fmt=None):
        lc = ws.cell(r, 1, label)
        lc.font = Font(bold=True, size=11)
        lc.fill = _fill(COR_PAR)
        lc.border = _borda()
        vc = ws.cell(r, 2, valor)
        vc.border = _borda()
        vc.alignment = Alignment(horizontal="right")
        if fmt:
            vc.number_format = fmt

    linha(3, "Periodo inicial",         datas[0],         FMT_DATA)
    linha(4, "Periodo final",            datas[-1],        FMT_DATA)
    linha(5, "Total de transacoes",      qtd)
    linha(7, "Total de entradas (R$)",   float(tot_cred),  FMT_MOEDA)
    linha(8, "Total de saidas (R$)",     float(tot_deb),   FMT_MOEDA)
    linha(9, "Saldo liquido do periodo", float(saldo),     FMT_MOEDA)
    ws.cell(9, 2).font = Font(bold=True, color=COR_CREDITO if saldo >= 0 else COR_DEBITO)

    meses = defaultdict(lambda: {"cred": Decimal(0), "deb": Decimal(0)})
    for t in transacoes:
        k = f"{t.data.year}-{t.data.month:02d}"
        if t.tipo == "Credito":
            meses[k]["cred"] += t.valor
        else:
            meses[k]["deb"] += abs(t.valor)
    if len(meses) > 1:
        r = 12
        ws.cell(r, 1, "Resumo Mensal").fill = _fill(COR_CAB_FUNDO)
        ws.cell(r, 1).font = Font(bold=True, size=11, color=COR_CAB_FONTE)
        ws.merge_cells(f"A{r}:D{r}")
        r += 1
        for cidx, titulo in enumerate(["Mes", "Entradas", "Saidas", "Saldo"], 1):
            c = ws.cell(r, cidx, titulo)
            c.font = Font(bold=True, color=COR_CAB_FONTE)
            c.fill = _fill(COR_CAB_FUNDO)
            c.border = _borda()
        r += 1
        for mes, v in sorted(meses.items()):
            sal_mes = v["cred"] - v["deb"]
            for cidx, val in enumerate([mes, float(v["cred"]), float(v["deb"]), float(sal_mes)], 1):
                cell = ws.cell(r, cidx, val)
                cell.border = _borda()
                if cidx > 1:
                    cell.number_format = FMT_MOEDA
            r += 1
    _ajustar_largura(ws)


def gerar_relatorio(transacoes: List[Transaction], output_path: str) -> str:
    """Gera relatorio Excel com abas Extrato e Resumo."""
    # Normaliza tipos para uso interno sem acentos (para comparacao robusta)
    for t in transacoes:
        if t.tipo not in ("Debito", "Credito"):
            t.tipo = "Credito" if t.valor >= 0 else "Debito"
    wb = openpyxl.Workbook()
    ws_ext = wb.active
    ws_ext.title = "Extrato"
    ws_res = wb.create_sheet("Resumo")
    _aba_extrato(ws_ext, transacoes)
    _aba_resumo(ws_res, transacoes)
    wb.save(output_path)
    return output_path
