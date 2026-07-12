#!/usr/bin/env python3
"""
Gerador da Planilha de Tesouraria Profissional para Igreja.
Cria arquivo Excel com 9 abas, fórmulas, gráficos, validações e proteção.
"""

from __future__ import annotations

import calendar
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

# ── Paleta de cores ──────────────────────────────────────────────────────────
NAVY = "1B2A4A"
GOLD = "C9A227"
WHITE = "FFFFFF"
LIGHT_YELLOW = "FFFACD"
LIGHT_GRAY = "F5F5F5"
LIGHT_BLUE = "E8EEF4"
GOLD_LIGHT = "F5E6B8"

DATA_START = 3
DATA_END = 502
YEAR = 2026

# ── Estilos reutilizáveis ────────────────────────────────────────────────────
def _font(bold=False, size=11, color="000000", italic=False):
    return Font(name="Calibri", bold=bold, size=size, color=color, italic=italic)


def _fill(hex_color: str):
    return PatternFill("solid", fgColor=hex_color)


def _border():
    thin = Side(style="thin", color="CCCCCC")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


TITLE_FONT = _font(bold=True, size=18, color=WHITE)
HEADER_FONT = _font(bold=True, size=11, color=WHITE)
KPI_FONT = _font(bold=True, size=14, color=NAVY)
KPI_VALUE_FONT = _font(bold=True, size=28, color=GOLD)
SECTION_FONT = _font(bold=True, size=12, color=NAVY)

TITLE_FILL = _fill(NAVY)
HEADER_FILL = _fill(NAVY)
GOLD_FILL = _fill(GOLD)
INPUT_FILL = _fill(LIGHT_YELLOW)
FORMULA_FILL = _fill(LIGHT_BLUE)
ALT_FILL = _fill(LIGHT_GRAY)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")

CURRENCY_FMT = 'R$ #,##0.00'
DATE_FMT = "DD/MM/YYYY"
PCT_FMT = "0%"

MONTHS_PT = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
]

RECEITAS = [
    "Dízimos", "Doações", "Missas", "Campanhas",
    "Festas", "Eventos", "Rifas", "Outros",
]

DESPESAS = [
    "Água", "Energia", "Internet", "Limpeza", "Material Litúrgico",
    "Flores", "Manutenção", "Construção", "Equipamentos",
    "Secretaria", "Eventos", "Transporte", "Outros",
]

GASTOS_CATEGORIAS = [
    "Água", "Energia", "Limpeza", "Material Litúrgico", "Hóstias", "Vinho",
    "Flores", "Som", "Manutenção", "Obras", "Secretaria", "Eventos",
    "Alimentação", "Transporte", "Outros",
]

FORMAS_PAGAMENTO = [
    "Dinheiro", "PIX", "Transferência", "Cartão Débito",
    "Cartão Crédito", "Cheque", "Depósito",
]

RESPONSAVEIS = [
    "Tesoureiro(a)", "Pároco", "Secretário(a)", "Vice-Tesoureiro(a)", "Outro",
]

DOACAO_TIPOS = [
    "Monetária", "Material", "Alimentos", "Vestuário", "Equipamento", "Outros",
]


def sundays_of_year(year: int) -> list[tuple[date, int]]:
    """Retorna lista de (data, número_do_domingo_no_ano)."""
    result: list[tuple[date, int]] = []
    d = date(year, 1, 1)
    while d.weekday() != 6:
        d += timedelta(days=1)
    n = 1
    while d.year == year:
        result.append((d, n))
        d += timedelta(days=7)
        n += 1
    return result


def style_title_row(ws, title: str, cols: int, icon: str = ""):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=cols)
    cell = ws.cell(row=1, column=1, value=f"{icon}  {title}" if icon else title)
    cell.font = TITLE_FONT
    cell.fill = TITLE_FILL
    cell.alignment = CENTER
    ws.row_dimensions[1].height = 40


def style_header_row(ws, headers: list[str], row: int = 2):
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = _border()
    ws.row_dimensions[row].height = 28


def apply_input_style(ws, row: int, col: int):
    cell = ws.cell(row=row, column=col)
    cell.fill = INPUT_FILL
    cell.border = _border()
    cell.protection = Protection(locked=False)


def apply_formula_style(ws, row: int, col: int):
    cell = ws.cell(row=row, column=col)
    cell.fill = FORMULA_FILL
    cell.border = _border()
    cell.protection = Protection(locked=True)


def add_autofilter(ws, last_col: int, header_row: int = 2):
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(last_col)}{header_row}"


def add_list_validation(ws, col: str, start_row: int, end_row: int, list_range: str):
    dv = DataValidation(
        type="list",
        formula1=f"={list_range}",
        allow_blank=True,
        showDropDown=False,
    )
    dv.error = "Selecione um valor da lista."
    dv.errorTitle = "Valor inválido"
    ws.add_data_validation(dv)
    dv.add(f"{col}{start_row}:{col}{end_row}")


def set_column_widths(ws, widths: dict[int, float]):
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width


def create_config_sheet(wb: Workbook):
    ws = wb.create_sheet("Configurações")
    style_title_row(ws, "Configurações — Listas e Parâmetros", 4, "⚙️")

    ws["A3"] = "Ano de Referência"
    ws["A3"].font = _font(bold=True, color=NAVY)
    ws["B3"] = YEAR
    apply_input_style(ws, 3, 2)
    ws["B3"].alignment = CENTER

    ws["A4"] = "Saldo Inicial do Ano"
    ws["A4"].font = _font(bold=True, color=NAVY)
    ws["B4"] = 0
    apply_input_style(ws, 4, 2)
    ws["B4"].number_format = CURRENCY_FMT

    sections = [
        ("A10", "Meses", MONTHS_PT, "B"),
        ("D10", "Formas de Pagamento", FORMAS_PAGAMENTO, "E"),
        ("G10", "Responsáveis", RESPONSAVEIS, "H"),
        ("A25", "Tipos de Receitas", RECEITAS, "B"),
        ("D25", "Tipos de Despesas", DESPESAS, "E"),
        ("G25", "Categorias de Gastos", GASTOS_CATEGORIAS, "H"),
        ("A40", "Tipos de Doação", DOACAO_TIPOS, "B"),
    ]

    for title_cell, title, items, col_letter in sections:
        ws[title_cell] = title
        ws[title_cell].font = SECTION_FONT
        ws[title_cell].fill = GOLD_FILL
        ws[title_cell].alignment = CENTER
        start_row = int(title_cell[1:]) + 1
        col_idx = ord(col_letter) - ord("A") + 1
        for i, item in enumerate(items):
            r = start_row + i
            ws.cell(row=r, column=col_idx, value=item).border = _border()

    # Nomes definidos via referências nas fórmulas
    set_column_widths(ws, {1: 22, 2: 18, 4: 22, 5: 18, 7: 22, 8: 18})
    ws.sheet_properties.tabColor = NAVY
    return ws


def create_dizimos_sheet(wb: Workbook):
    ws = wb.create_sheet("Dízimos Mensais")
    headers = ["Data", "Nome", "Telefone", "Mês", "Valor", "Forma de Pagamento", "Observação"]
    style_title_row(ws, "Dízimos Mensais", len(headers), "💰")
    style_header_row(ws, headers)

    for row in range(DATA_START, DATA_END + 1):
        apply_input_style(ws, row, 1)
        ws.cell(row=row, column=1).number_format = DATE_FMT
        for col in (2, 3, 6, 7):
            apply_input_style(ws, row, col)
        # Mês automático a partir da data
        ws.cell(row=row, column=4).value = (
            f'=SE(A{row}="";"";ÍNDICE(Configurações!$B$11:$B$22;MÊS(A{row})))'
        )
        apply_formula_style(ws, row, 4)
        apply_input_style(ws, row, 5)
        ws.cell(row=row, column=5).number_format = CURRENCY_FMT

    add_list_validation(ws, "F", DATA_START, DATA_END, "Configurações!$E$11:$E$17")
    add_autofilter(ws, len(headers))
    set_column_widths(ws, {1: 14, 2: 28, 3: 16, 4: 14, 5: 14, 6: 20, 7: 30})
    ws.sheet_properties.tabColor = GOLD
    return ws


def create_doacoes_sheet(wb: Workbook):
    ws = wb.create_sheet("Doações")
    headers = ["Data", "Doador", "Tipo da Doação", "Valor", "Destinação", "Observação"]
    style_title_row(ws, "Doações", len(headers), "🎁")
    style_header_row(ws, headers)

    for row in range(DATA_START, DATA_END + 1):
        for col in (1, 2, 4, 5, 6):
            apply_input_style(ws, row, col)
        ws.cell(row=row, column=1).number_format = DATE_FMT
        ws.cell(row=row, column=4).number_format = CURRENCY_FMT

    add_list_validation(ws, "C", DATA_START, DATA_END, "Configurações!$B$41:$B$46")
    add_autofilter(ws, len(headers))
    set_column_widths(ws, {1: 14, 2: 28, 3: 18, 4: 14, 5: 22, 6: 30})
    ws.sheet_properties.tabColor = GOLD
    return ws


def create_missas_sheet(wb: Workbook):
    ws = wb.create_sheet("Arrecadação das Missas")
    headers = ["Data", "Domingo", "Mês", "Valor", "Responsável", "Observação"]
    style_title_row(ws, "Arrecadação das Missas", len(headers), "⛪")
    style_header_row(ws, headers)

    sundays = sundays_of_year(YEAR)
    row = DATA_START
    for sunday_date, num in sundays:
        ws.cell(row=row, column=1, value=sunday_date)
        apply_input_style(ws, row, 1)
        ws.cell(row=row, column=1).number_format = DATE_FMT

        ws.cell(row=row, column=2, value=f"Domingo {num}")
        apply_formula_style(ws, row, 2)

        ws.cell(row=row, column=3).value = (
            f'=SE(A{row}="";"";ÍNDICE(Configurações!$B$11:$B$22;MÊS(A{row})))'
        )
        apply_formula_style(ws, row, 3)

        apply_input_style(ws, row, 4)
        ws.cell(row=row, column=4).number_format = CURRENCY_FMT
        apply_input_style(ws, row, 5)
        apply_input_style(ws, row, 6)
        row += 1

    # Linhas extras para missas especiais
    for extra_row in range(row, DATA_END + 1):
        apply_input_style(ws, extra_row, 1)
        ws.cell(row=extra_row, column=1).number_format = DATE_FMT
        apply_input_style(ws, extra_row, 2)
        ws.cell(row=extra_row, column=3).value = (
            f'=SE(A{extra_row}="";"";ÍNDICE(Configurações!$B$11:$B$22;MÊS(A{extra_row})))'
        )
        apply_formula_style(ws, extra_row, 3)
        apply_input_style(ws, extra_row, 4)
        ws.cell(row=extra_row, column=4).number_format = CURRENCY_FMT
        apply_input_style(ws, extra_row, 5)
        apply_input_style(ws, extra_row, 6)

    add_list_validation(ws, "E", DATA_START, DATA_END, "Configurações!$H$11:$H$15")
    add_autofilter(ws, len(headers))
    set_column_widths(ws, {1: 14, 2: 16, 3: 14, 4: 14, 5: 20, 6: 30})
    ws.sheet_properties.tabColor = GOLD
    return ws


def create_gastos_sheet(wb: Workbook):
    ws = wb.create_sheet("Gastos")
    headers = [
        "Data", "Categoria", "Descrição", "Fornecedor",
        "Forma de Pagamento", "Valor", "Responsável", "Observação",
    ]
    style_title_row(ws, "Gastos", len(headers), "💸")
    style_header_row(ws, headers)

    for row in range(DATA_START, DATA_END + 1):
        for col in (1, 2, 3, 4, 5, 6, 7, 8):
            apply_input_style(ws, row, col)
        ws.cell(row=row, column=1).number_format = DATE_FMT
        ws.cell(row=row, column=6).number_format = CURRENCY_FMT

    add_list_validation(ws, "B", DATA_START, DATA_END, "Configurações!$H$26:$H$40")
    add_list_validation(ws, "E", DATA_START, DATA_END, "Configurações!$E$11:$E$17")
    add_list_validation(ws, "G", DATA_START, DATA_END, "Configurações!$H$11:$H$15")
    add_autofilter(ws, len(headers))
    set_column_widths(ws, {1: 14, 2: 18, 3: 28, 4: 22, 5: 20, 6: 14, 7: 20, 8: 28})
    ws.sheet_properties.tabColor = GOLD
    return ws


def create_resumo_sheet(wb: Workbook):
    ws = wb.create_sheet("Resumo Financeiro")
    headers = ["Mês", "Dízimos", "Doações", "Missas", "Total Entradas", "Gastos", "Saldo do Mês"]
    style_title_row(ws, "Resumo Financeiro Anual", len(headers), "📊")
    style_header_row(ws, headers)

    for i, month in enumerate(MONTHS_PT):
        row = DATA_START + i
        ws.cell(row=row, column=1, value=month)
        ws.cell(row=row, column=1).font = _font(bold=True, color=NAVY)
        ws.cell(row=row, column=1).border = _border()

        month_num = i + 1
        # Dízimos por nome do mês
        ws.cell(row=row, column=2).value = (
            f'=SOMASE(\'Dízimos Mensais\'!$D${DATA_START}:$D${DATA_END};A{row};'
            f'\'Dízimos Mensais\'!$E${DATA_START}:$E${DATA_END})'
        )
        # Doações por intervalo de datas do mês
        ws.cell(row=row, column=3).value = (
            f'=SOMASES(Doações!$D${DATA_START}:$D${DATA_END};'
            f'Doações!$A${DATA_START}:$A${DATA_END};">="&DATA(Configurações!$B$3;{month_num};1);'
            f'Doações!$A${DATA_START}:$A${DATA_END};"<="&FIM.MÊS(DATA(Configurações!$B$3;{month_num};1)))'
        )
        # Missas por nome do mês
        ws.cell(row=row, column=4).value = (
            f'=SOMASE(\'Arrecadação das Missas\'!$C${DATA_START}:$C${DATA_END};A{row};'
            f'\'Arrecadação das Missas\'!$D${DATA_START}:$D${DATA_END})'
        )
        ws.cell(row=row, column=5).value = f"=SOMA(B{row}:D{row})"
        ws.cell(row=row, column=6).value = (
            f'=SOMASES(Gastos!$F${DATA_START}:$F${DATA_END};'
            f'Gastos!$A${DATA_START}:$A${DATA_END};">="&DATA(Configurações!$B$3;{month_num};1);'
            f'Gastos!$A${DATA_START}:$A${DATA_END};"<="&FIM.MÊS(DATA(Configurações!$B$3;{month_num};1)))'
        )
        ws.cell(row=row, column=7).value = f"=E{row}-F{row}"

        for col in range(2, 8):
            apply_formula_style(ws, row, col)
            ws.cell(row=row, column=col).number_format = CURRENCY_FMT
            ws.cell(row=row, column=col).alignment = RIGHT

    # Linha de totais
    total_row = DATA_START + 12
    ws.cell(row=total_row, column=1, value="TOTAL ANUAL")
    ws.cell(row=total_row, column=1).font = _font(bold=True, size=12, color=WHITE)
    ws.cell(row=total_row, column=1).fill = TITLE_FILL
    for col in range(2, 8):
        letter = get_column_letter(col)
        ws.cell(row=total_row, column=col).value = f"=SOMA({letter}{DATA_START}:{letter}{DATA_START + 11})"
        ws.cell(row=total_row, column=col).font = _font(bold=True, color=WHITE)
        ws.cell(row=total_row, column=col).fill = TITLE_FILL
        ws.cell(row=total_row, column=col).number_format = CURRENCY_FMT
        ws.cell(row=total_row, column=col).alignment = RIGHT

    add_autofilter(ws, len(headers))
    set_column_widths(ws, {1: 16, 2: 14, 3: 14, 4: 14, 5: 16, 6: 14, 7: 16})
    ws.sheet_properties.tabColor = NAVY
    return ws


def create_plano_contas_sheet(wb: Workbook):
    ws = wb.create_sheet("Plano de Contas")
    style_title_row(ws, "Plano de Contas", 4, "📋")

    ws.merge_cells("A3:B3")
    ws["A3"] = "RECEITAS"
    ws["A3"].font = _font(bold=True, size=13, color=WHITE)
    ws["A3"].fill = _fill("2E7D32")
    ws["A3"].alignment = CENTER

    ws.merge_cells("D3:E3")
    ws["D3"] = "DESPESAS"
    ws["D3"].font = _font(bold=True, size=13, color=WHITE)
    ws["D3"].fill = _fill("C62828")
    ws["D3"].alignment = CENTER

    ws["A4"] = "Código"
    ws["B4"] = "Descrição"
    ws["D4"] = "Código"
    ws["E4"] = "Descrição"
    for cell_ref in ("A4", "B4", "D4", "E4"):
        ws[cell_ref].font = HEADER_FONT
        ws[cell_ref].fill = HEADER_FILL
        ws[cell_ref].alignment = CENTER

    for i, receita in enumerate(RECEITAS, 1):
        row = 4 + i
        ws.cell(row=row, column=1, value=f"R{i:02d}").alignment = CENTER
        ws.cell(row=row, column=2, value=receita)
        ws.cell(row=row, column=1).border = _border()
        ws.cell(row=row, column=2).border = _border()

    for i, despesa in enumerate(DESPESAS, 1):
        row = 4 + i
        ws.cell(row=row, column=4, value=f"D{i:02d}").alignment = CENTER
        ws.cell(row=row, column=5, value=despesa)
        ws.cell(row=row, column=4).border = _border()
        ws.cell(row=row, column=5).border = _border()

    ws["A20"] = "ℹ️ Utilize este plano de contas como referência ao registrar receitas e despesas."
    ws["A20"].font = _font(italic=True, color="666666")
    ws.merge_cells("A20:E20")

    set_column_widths(ws, {1: 10, 2: 24, 4: 10, 5: 24})
    ws.sheet_properties.tabColor = NAVY
    return ws


def create_aux_fluxo_sheet(wb: Workbook):
    """Aba auxiliar (oculta) que consolida lançamentos para o Fluxo de Caixa."""
    ws = wb.create_sheet("_Consolidado")
    headers = ["Data", "Descrição", "Entrada", "Saída", "Origem"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)

    row_out = 2
    # Dízimos
    for r in range(DATA_START, DATA_END + 1):
        ws.cell(row=row_out, column=1).value = f"=SE('Dízimos Mensais'!A{r}=\"\";\"\";'Dízimos Mensais'!A{r})"
        ws.cell(row=row_out, column=2).value = f"=SE('Dízimos Mensais'!A{r}=\"\";\"\";\"Dízimo - \"&'Dízimos Mensais'!B{r})"
        ws.cell(row=row_out, column=3).value = f"=SE('Dízimos Mensais'!A{r}=\"\";\"\";'Dízimos Mensais'!E{r})"
        ws.cell(row=row_out, column=4).value = f"=SE('Dízimos Mensais'!A{r}=\"\";\"\";0)"
        ws.cell(row=row_out, column=5, value="Dízimos")
        row_out += 1

    # Doações
    for r in range(DATA_START, DATA_END + 1):
        ws.cell(row=row_out, column=1).value = f"=SE(Doações!A{r}=\"\";\"\";Doações!A{r})"
        ws.cell(row=row_out, column=2).value = f"=SE(Doações!A{r}=\"\";\"\";\"Doação - \"&Doações!B{r})"
        ws.cell(row=row_out, column=3).value = f"=SE(Doações!A{r}=\"\";\"\";Doações!D{r})"
        ws.cell(row=row_out, column=4).value = f"=SE(Doações!A{r}=\"\";\"\";0)"
        ws.cell(row=row_out, column=5, value="Doações")
        row_out += 1

    # Missas
    for r in range(DATA_START, DATA_END + 1):
        ws.cell(row=row_out, column=1).value = f"=SE('Arrecadação das Missas'!A{r}=\"\";\"\";'Arrecadação das Missas'!A{r})"
        ws.cell(row=row_out, column=2).value = f"=SE('Arrecadação das Missas'!A{r}=\"\";\"\";\"Missas - \"&'Arrecadação das Missas'!B{r})"
        ws.cell(row=row_out, column=3).value = f"=SE('Arrecadação das Missas'!A{r}=\"\";\"\";'Arrecadação das Missas'!D{r})"
        ws.cell(row=row_out, column=4).value = f"=SE('Arrecadação das Missas'!A{r}=\"\";\"\";0)"
        ws.cell(row=row_out, column=5, value="Missas")
        row_out += 1

    # Gastos
    for r in range(DATA_START, DATA_END + 1):
        ws.cell(row=row_out, column=1).value = f"=SE(Gastos!A{r}=\"\";\"\";Gastos!A{r})"
        ws.cell(row=row_out, column=2).value = f"=SE(Gastos!A{r}=\"\";\"\";Gastos!B{r}&\" - \"&Gastos!C{r})"
        ws.cell(row=row_out, column=3).value = f"=SE(Gastos!A{r}=\"\";\"\";0)"
        ws.cell(row=row_out, column=4).value = f"=SE(Gastos!A{r}=\"\";\"\";Gastos!F{r})"
        ws.cell(row=row_out, column=5, value="Gastos")
        row_out += 1

    ws.sheet_state = "hidden"
    return ws


def create_fluxo_sheet(wb: Workbook):
    ws = wb.create_sheet("Fluxo de Caixa")
    headers = ["Data", "Descrição", "Entrada", "Saída", "Saldo"]
    style_title_row(ws, "Fluxo de Caixa", len(headers), "💵")
    style_header_row(ws, headers)

    ws["G2"] = "Saldo Inicial:"
    ws["G2"].font = _font(bold=True, color=NAVY)
    ws["H2"] = "=Configurações!B4"
    ws["H2"].number_format = CURRENCY_FMT
    apply_formula_style(ws, 2, 8)

    # Linha de saldo inicial no fluxo
    ws.cell(row=DATA_START, column=2, value="Saldo Inicial")
    ws.cell(row=DATA_START, column=2).font = _font(bold=True, color=NAVY)
    ws.cell(row=DATA_START, column=5).value = "=Configurações!B4"
    apply_formula_style(ws, DATA_START, 5)
    ws.cell(row=DATA_START, column=5).number_format = CURRENCY_FMT

    fluxo_start = DATA_START + 1
    max_fluxo_rows = 500
    for i in range(max_fluxo_rows):
        fluxo_row = fluxo_start + i
        aux_row = 2 + i
        ws.cell(row=fluxo_row, column=1).value = f"=_Consolidado!A{aux_row}"
        ws.cell(row=fluxo_row, column=2).value = f"=_Consolidado!B{aux_row}"
        ws.cell(row=fluxo_row, column=3).value = f"=_Consolidado!C{aux_row}"
        ws.cell(row=fluxo_row, column=4).value = f"=_Consolidado!D{aux_row}"
        prev = fluxo_row - 1
        ws.cell(row=fluxo_row, column=5).value = (
            f'=SE(E{prev}="";"";SE(A{fluxo_row}="";E{prev};E{prev}+SEERRO(C{fluxo_row};0)-SEERRO(D{fluxo_row};0)))'
        )
        for col in range(1, 6):
            apply_formula_style(ws, fluxo_row, col)
        ws.cell(row=fluxo_row, column=1).number_format = DATE_FMT
        ws.cell(row=fluxo_row, column=3).number_format = CURRENCY_FMT
        ws.cell(row=fluxo_row, column=4).number_format = CURRENCY_FMT
        ws.cell(row=fluxo_row, column=5).number_format = CURRENCY_FMT

    ws["G3"] = "💡 Dica: Ordene a coluna Data (A) para visualizar cronologicamente."
    ws["G3"].font = _font(italic=True, color="666666")
    ws.merge_cells("G3:H3")

    add_autofilter(ws, len(headers))
    set_column_widths(ws, {1: 14, 2: 40, 3: 14, 4: 14, 5: 16, 7: 14, 8: 14})
    ws.sheet_properties.tabColor = NAVY
    return ws


def create_dashboard_sheet(wb: Workbook):
    ws = wb.create_sheet("Dashboard")
    style_title_row(ws, "Tesouraria — Dashboard Financeiro", 8, "📑")

    # ── KPIs ─────────────────────────────────────────────────────────────────
    kpi_row = 3
    ws.merge_cells(start_row=kpi_row, start_column=3, end_row=kpi_row + 2, end_column=6)
    saldo_cell = ws.cell(row=kpi_row, column=3, value="🟢 SALDO ATUAL")
    saldo_cell.font = _font(bold=True, size=12, color=NAVY)
    saldo_cell.alignment = CENTER

    ws.merge_cells(start_row=kpi_row + 3, start_column=3, end_row=kpi_row + 5, end_column=6)
    saldo_val = ws.cell(row=kpi_row + 3, column=3)
    saldo_val.value = (
        f"=Configurações!B4+SOMA('Resumo Financeiro'!E{DATA_START}:E{DATA_START + 11})"
        f"-SOMA('Resumo Financeiro'!F{DATA_START}:F{DATA_START + 11})"
    )
    saldo_val.font = KPI_VALUE_FONT
    saldo_val.alignment = CENTER
    saldo_val.number_format = CURRENCY_FMT
    saldo_val.fill = GOLD_FILL

    kpis = [
        (1, kpi_row, "💰 Total de Entradas", f"='Resumo Financeiro'!E{DATA_START + 12}"),
        (8, kpi_row, "💸 Total de Gastos", f"='Resumo Financeiro'!F{DATA_START + 12}"),
        (1, kpi_row + 4, "📈 Resultado do Ano", f"='Resumo Financeiro'!G{DATA_START + 12}"),
        (8, kpi_row + 4, "📅 Entradas do Mês", (
            f"=ÍNDICE('Resumo Financeiro'!E{DATA_START}:E{DATA_START + 11};MÊS(HOJE()))"
        )),
        (1, kpi_row + 8, "📅 Gastos do Mês", (
            f"=ÍNDICE('Resumo Financeiro'!F{DATA_START}:F{DATA_START + 11};MÊS(HOJE()))"
        )),
    ]

    for col, row, label, formula in kpis:
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
        lbl = ws.cell(row=row, column=col, value=label)
        lbl.font = KPI_FONT
        lbl.fill = _fill(LIGHT_BLUE)
        lbl.alignment = CENTER
        lbl.border = _border()

        ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + 1)
        val = ws.cell(row=row + 1, column=col, value=formula)
        val.font = _font(bold=True, size=16, color=NAVY)
        val.number_format = CURRENCY_FMT
        val.alignment = CENTER
        val.fill = _fill(WHITE)
        val.border = _border()

    # ── Área de dados para gráficos (oculta visualmente à direita) ───────────
    chart_data_col = 10  # coluna J
    ws.cell(row=2, column=chart_data_col, value="Mês").font = _font(bold=True)
    headers_chart = ["Entradas", "Gastos", "Saldo Acum.", "Dízimos", "Doações", "Missas"]
    for i, h in enumerate(headers_chart):
        ws.cell(row=2, column=chart_data_col + 1 + i, value=h).font = _font(bold=True)

    for i, month in enumerate(MONTHS_PT):
        r = 3 + i
        ws.cell(row=r, column=chart_data_col, value=month[:3])
        ws.cell(row=r, column=chart_data_col + 1).value = f"='Resumo Financeiro'!E{DATA_START + i}"
        ws.cell(row=r, column=chart_data_col + 2).value = f"='Resumo Financeiro'!F{DATA_START + i}"
        ws.cell(row=r, column=chart_data_col + 3).value = (
            f"=Configurações!$B$4+SOMA($K$3:K{r})-SOMA($L$3:L{r})"
        )
        ws.cell(row=r, column=chart_data_col + 4).value = f"='Resumo Financeiro'!B{DATA_START + i}"
        ws.cell(row=r, column=chart_data_col + 5).value = f"='Resumo Financeiro'!C{DATA_START + i}"
        ws.cell(row=r, column=chart_data_col + 6).value = f"='Resumo Financeiro'!D{DATA_START + i}"

    # Ranking meses (col P)
    ws.cell(row=2, column=16, value="Ranking").font = _font(bold=True)
    ws.cell(row=2, column=17, value="Mês").font = _font(bold=True)
    ws.cell(row=2, column=18, value="Entradas").font = _font(bold=True)
    for i in range(12):
        r = 3 + i
        ws.cell(row=r, column=16, value=i + 1)
        ws.cell(row=r, column=17).value = (
            f"=ÍNDICE($J$3:$J$14;CORRESP(MAIOR($K$3:$K$14;{i + 1});$K$3:$K$14;0))"
        )
        ws.cell(row=r, column=18).value = (
            f"=ÍNDICE($K$3:$K$14;CORRESP(MAIOR($K$3:$K$14;{i + 1});$K$3:$K$14;0))"
        )

    # Gastos por categoria (col S)
    ws.cell(row=2, column=19, value="Categoria").font = _font(bold=True)
    ws.cell(row=2, column=20, value="Total").font = _font(bold=True)
    for i, cat in enumerate(GASTOS_CATEGORIAS):
        r = 3 + i
        ws.cell(row=r, column=19, value=cat)
        ws.cell(row=r, column=20).value = (
            f'=SOMASE(Gastos!$B${DATA_START}:$B${DATA_END};S{r};Gastos!$F${DATA_START}:$F${DATA_END})'
        )

    chart_top = 14

    # Gráfico 1: Entradas x Gastos por mês
    chart1 = BarChart()
    chart1.type = "col"
    chart1.grouping = "clustered"
    chart1.title = "📊 Entradas x Gastos por Mês"
    chart1.y_axis.title = "Valor (R$)"
    chart1.style = 10
    chart1.width = 18
    chart1.height = 10
    cats1 = Reference(ws, min_col=chart_data_col, min_row=3, max_row=14)
    data1 = Reference(ws, min_col=chart_data_col + 1, min_row=2, max_col=chart_data_col + 2, max_row=14)
    chart1.add_data(data1, titles_from_data=True)
    chart1.set_categories(cats1)
    ws.add_chart(chart1, f"A{chart_top}")

    # Gráfico 2: Evolução do saldo
    chart2 = LineChart()
    chart2.title = "📈 Evolução do Saldo no Ano"
    chart2.y_axis.title = "Saldo (R$)"
    chart2.style = 10
    chart2.width = 18
    chart2.height = 10
    data2 = Reference(ws, min_col=chart_data_col + 3, min_row=2, max_row=14)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats1)
    ws.add_chart(chart2, f"J{chart_top}")

    # Gráfico 3: Origem das receitas (pizza)
    chart3 = PieChart()
    chart3.title = "🍕 Origem das Receitas"
    chart3.style = 10
    chart3.width = 14
    chart3.height = 10
    # Totais anuais
    ws.cell(row=16, column=chart_data_col, value="Dízimos")
    ws.cell(row=17, column=chart_data_col, value="Doações")
    ws.cell(row=18, column=chart_data_col, value="Missas")
    ws.cell(row=16, column=chart_data_col + 1).value = f"='Resumo Financeiro'!B{DATA_START + 12}"
    ws.cell(row=17, column=chart_data_col + 1).value = f"='Resumo Financeiro'!C{DATA_START + 12}"
    ws.cell(row=18, column=chart_data_col + 1).value = f"='Resumo Financeiro'!D{DATA_START + 12}"
    cats3 = Reference(ws, min_col=chart_data_col, min_row=16, max_row=18)
    data3 = Reference(ws, min_col=chart_data_col + 1, min_row=16, max_row=18)
    chart3.add_data(data3)
    chart3.set_categories(cats3)
    chart3.dataLabels = DataLabelList()
    chart3.dataLabels.showPercent = True
    ws.add_chart(chart3, f"A{chart_top + 14}")

    # Gráfico 4: Gastos por categoria
    chart4 = PieChart()
    chart4.title = "🍕 Gastos por Categoria"
    chart4.style = 10
    chart4.width = 14
    chart4.height = 10
    cats4 = Reference(ws, min_col=19, min_row=3, max_row=3 + len(GASTOS_CATEGORIAS) - 1)
    data4 = Reference(ws, min_col=20, min_row=3, max_row=3 + len(GASTOS_CATEGORIAS) - 1)
    chart4.add_data(data4)
    chart4.set_categories(cats4)
    chart4.dataLabels = DataLabelList()
    chart4.dataLabels.showPercent = True
    ws.add_chart(chart4, f"J{chart_top + 14}")

    # Gráfico 5: Comparativo mensal (empilhado)
    chart5 = BarChart()
    chart5.type = "col"
    chart5.grouping = "stacked"
    chart5.title = "📊 Comparativo Mensal de Receitas"
    chart5.style = 10
    chart5.width = 18
    chart5.height = 10
    data5 = Reference(ws, min_col=chart_data_col + 4, min_row=2, max_col=chart_data_col + 6, max_row=14)
    chart5.add_data(data5, titles_from_data=True)
    chart5.set_categories(cats1)
    ws.add_chart(chart5, f"A{chart_top + 28}")

    # Gráfico 6: Ranking top 5 meses
    chart6 = BarChart()
    chart6.type = "bar"
    chart6.title = "🏆 Top 5 Meses — Maior Arrecadação"
    chart6.style = 10
    chart6.width = 18
    chart6.height = 10
    cats6 = Reference(ws, min_col=17, min_row=3, max_row=7)
    data6 = Reference(ws, min_col=18, min_row=3, max_row=7)
    chart6.add_data(data6)
    chart6.set_categories(cats6)
    ws.add_chart(chart6, f"J{chart_top + 28}")

    set_column_widths(ws, {1: 14, 2: 14, 3: 14, 4: 14, 5: 14, 6: 14, 7: 14, 8: 14})
    ws.sheet_properties.tabColor = GOLD

    # Ocultar colunas de dados auxiliares
    for col_idx in range(chart_data_col, 21):
        ws.column_dimensions[get_column_letter(col_idx)].hidden = True

    return ws


def protect_sheets(wb: Workbook):
    """Protege abas: células de fórmula bloqueadas, entradas liberadas."""
    password = ""  # sem senha — usuário pode desproteger facilmente
    for ws in wb.worksheets:
        if ws.title == "_Consolidado":
            ws.protection.sheet = True
            continue
        ws.protection.sheet = True
        ws.protection.formatCells = False
        ws.protection.formatColumns = False
        ws.protection.formatRows = False
        ws.protection.insertColumns = False
        ws.protection.insertRows = False
        ws.protection.deleteColumns = False
        ws.protection.deleteRows = False
        ws.protection.sort = True
        ws.protection.autoFilter = True
        ws.protection.selectLockedCells = True
        ws.protection.selectUnlockedCells = True


def generate(output_path: Path) -> Path:
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    create_dizimos_sheet(wb)
    create_doacoes_sheet(wb)
    create_missas_sheet(wb)
    create_gastos_sheet(wb)
    create_resumo_sheet(wb)
    create_plano_contas_sheet(wb)
    create_config_sheet(wb)
    create_aux_fluxo_sheet(wb)
    create_fluxo_sheet(wb)
    create_dashboard_sheet(wb)

    sheet_order = [
        "Dashboard",
        "Dízimos Mensais",
        "Doações",
        "Arrecadação das Missas",
        "Gastos",
        "Resumo Financeiro",
        "Plano de Contas",
        "Fluxo de Caixa",
        "Configurações",
        "_Consolidado",
    ]
    for target_idx, name in enumerate(sheet_order):
        current_idx = wb.sheetnames.index(name)
        if current_idx != target_idx:
            wb.move_sheet(wb[name], target_idx - current_idx)

    protect_sheets(wb)

    wb.active = wb["Dashboard"]
    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    out = Path(__file__).parent / "Tesouraria_Igreja.xlsx"
    path = generate(out)
    print(f"Planilha gerada: {path}")
