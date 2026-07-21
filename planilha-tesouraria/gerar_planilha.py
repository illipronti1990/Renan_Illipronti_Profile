#!/usr/bin/env python3
"""
Gerador da Planilha de Tesouraria Profissional para Igreja — Sistema Completo.
Inclui: saldo caixa/banco, conciliação, relatórios, alertas, impressão A4,
cadastro de dizimistas, comparativo anual, metas e macro Novo Exercício.
"""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ── Paleta ───────────────────────────────────────────────────────────────────
NAVY = "1B2A4A"
GOLD = "C9A227"
WHITE = "FFFFFF"
LIGHT_YELLOW = "FFFACD"
LIGHT_GRAY = "F5F5F5"
LIGHT_BLUE = "E8EEF4"
RED_LIGHT = "FFC7CE"
GREEN_LIGHT = "C6EFCE"
YELLOW_LIGHT = "FFEB9C"
ORANGE_LIGHT = "FCE4D6"

DATA_START = 3
DATA_END = 502
YEAR = 2026

# ── Referências da aba Configurações ───────────────────────────────────────
CFG_ANO = "$B$3"
CFG_IGREJA = "$B$4"
CFG_CAIXA = "$B$8"
CFG_BANCO = "$B$9"
CFG_SALDO_TOTAL = "$B$10"
CFG_METAS = "$B$14:$M$14"
CFG_MESES = "$B$35:$B$46"
CFG_FILTRO_MESES = "$B$34:$B$46"
CFG_FORMAS = "$E$36:$E$42"
CFG_RESP = "$H$36:$H$40"
CFG_CONTAS = "$J$36:$J$37"
CFG_RECEITAS = "$B$51:$B$58"
CFG_DOACAO_TIPOS = "$B$70:$B$75"
CFG_CATEGORIAS = "$H$51:$H$65"
CFG_LIMITES_CAT = "$I$51:$I$65"

MONTHS_PT = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
]
FILTER_MONTHS = ["Ano Inteiro"] + MONTHS_PT

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
CONTAS = ["Caixa", "Banco"]

RESUMO_TOTAL_ROW = DATA_START + 12


def _font(bold=False, size=11, color="000000", italic=False):
    return Font(name="Calibri", bold=bold, size=size, color=color, italic=italic)


def _fill(hex_color: str):
    return PatternFill("solid", fgColor=hex_color)


def _border():
    thin = Side(style="thin", color="CCCCCC")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


TITLE_FONT = _font(bold=True, size=18, color=WHITE)
HEADER_FONT = _font(bold=True, size=11, color=WHITE)
KPI_FONT = _font(bold=True, size=13, color=NAVY)
KPI_VALUE_FONT = _font(bold=True, size=26, color=GOLD)
SECTION_FONT = _font(bold=True, size=12, color=NAVY)
TITLE_FILL = _fill(NAVY)
HEADER_FILL = _fill(NAVY)
GOLD_FILL = _fill(GOLD)
INPUT_FILL = _fill(LIGHT_YELLOW)
FORMULA_FILL = _fill(LIGHT_BLUE)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")
CURRENCY_FMT = 'R$ #,##0.00'
DATE_FMT = "DD/MM/YYYY"
PCT_FMT = "0.0%"


def sundays_of_year(year: int) -> list[tuple[date, int]]:
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
    dv = DataValidation(type="list", formula1=f"={list_range}", allow_blank=True, showDropDown=False)
    dv.error = "Selecione um valor da lista."
    dv.errorTitle = "Valor inválido"
    ws.add_data_validation(dv)
    dv.add(f"{col}{start_row}:{col}{end_row}")


def add_date_validation(ws, col: str, start_row: int, end_row: int):
    dv = DataValidation(
        type="date", operator="between",
        formula1="DATE(2020,1,1)", formula2="DATE(2035,12,31)",
        allow_blank=True,
    )
    dv.error = "Informe uma data válida (dd/mm/aaaa)."
    ws.add_data_validation(dv)
    dv.add(f"{col}{start_row}:{col}{end_row}")


def set_column_widths(ws, widths: dict[int, float]):
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width


def setup_print_a4(ws, title: str = ""):
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    if title:
        ws.oddHeader.center.text = title
    ws.oddFooter.center.text = "Gerado pela Planilha de Tesouraria — Página &P de &N"


def add_signature_block(ws, row: int, col_start: int = 1, col_end: int = 6):
    ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_start + 2)
    ws.merge_cells(start_row=row, start_column=col_end - 2, end_row=row, end_column=col_end)
    ws.cell(row=row, column=col_start, value="_" * 30).alignment = CENTER
    ws.cell(row=row, column=col_end - 2, value="_" * 30).alignment = CENTER
    ws.cell(row=row + 1, column=col_start, value="Tesoureiro(a)").alignment = CENTER
    ws.cell(row=row + 1, column=col_end - 2, value="Pároco / Presidente").alignment = CENTER


def print_report_header(ws, row: int, titulo: str, periodo_formula: str, cols: int = 6):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    ws.cell(row=row, column=1, value=f"=Configurações!{CFG_IGREJA}").font = _font(bold=True, size=16, color=NAVY)
    ws.cell(row=row, column=1).alignment = CENTER
    ws.merge_cells(start_row=row + 1, start_column=1, end_row=row + 1, end_column=cols)
    ws.cell(row=row + 1, column=1, value=titulo).font = _font(bold=True, size=13, color=NAVY)
    ws.cell(row=row + 1, column=1).alignment = CENTER
    ws.cell(row=row + 2, column=1, value="Período:").font = _font(bold=True)
    ws.cell(row=row + 2, column=2, value=periodo_formula)
    ws.cell(row=row + 2, column=4, value="Emissão:").font = _font(bold=True)
    ws.cell(row=row + 2, column=5, value="=HOJE()")
    ws.cell(row=row + 2, column=5).number_format = DATE_FMT


def create_config_sheet(wb: Workbook, year: int = YEAR):
    ws = wb.create_sheet("Configurações")
    style_title_row(ws, "Configurações — Parâmetros e Listas", 8, "⚙️")

    fields = [
        (3, "Ano de Referência", year, False),
        (4, "Nome da Igreja", "Paróquia / Comunidade", True),
    ]
    for row, label, val, is_text in fields:
        ws.cell(row=row, column=1, value=label).font = _font(bold=True, color=NAVY)
        ws.cell(row=row, column=2, value=val)
        apply_input_style(ws, row, 2)

    ws.merge_cells("A6:B6")
    ws["A6"] = "💰 SALDO INICIAL (antes dos lançamentos do ano)"
    ws["A6"].font = SECTION_FONT
    ws["A6"].fill = GOLD_FILL
    ws["A6"].alignment = CENTER

    ws["A7"], ws["B7"] = "Conta", "Saldo Inicial"
    for c in ("A7", "B7"):
        ws[c].font = HEADER_FONT
        ws[c].fill = HEADER_FILL
        ws[c].alignment = CENTER

    saldos = [("Caixa", 0), ("Banco", 0)]
    for i, (conta, val) in enumerate(saldos, 8):
        ws.cell(row=i, column=1, value=conta).font = _font(bold=True, color=NAVY)
        ws.cell(row=i, column=2, value=val)
        apply_input_style(ws, i, 2)
        ws.cell(row=i, column=2).number_format = CURRENCY_FMT

    ws.cell(row=10, column=1, value="Total").font = _font(bold=True, size=12, color=WHITE)
    ws.cell(row=10, column=1).fill = TITLE_FILL
    ws.cell(row=10, column=2, value=f"=SOMA(B8:B9)")
    apply_formula_style(ws, 10, 2)
    ws.cell(row=10, column=2).number_format = CURRENCY_FMT
    ws.cell(row=10, column=2).font = _font(bold=True, size=12, color=NAVY)

    ws.merge_cells("A12:M12")
    ws["A12"] = "🎯 METAS MENSAIS DE ARRECADAÇÃO"
    ws["A12"].font = SECTION_FONT
    ws["A12"].fill = GOLD_FILL
    ws["A12"].alignment = CENTER
    for i, m in enumerate(MONTHS_PT):
        col = 2 + i
        ws.cell(row=13, column=col, value=m[:3]).font = _font(bold=True, color=NAVY)
        ws.cell(row=13, column=col).alignment = CENTER
        ws.cell(row=14, column=col, value=0)
        apply_input_style(ws, 14, col)
        ws.cell(row=14, column=col).number_format = CURRENCY_FMT

    ws.merge_cells("A16:C16")
    ws["A16"] = "⚠️ LIMITES MENSAIS DE GASTOS POR CATEGORIA"
    ws["A16"].font = SECTION_FONT
    ws["A16"].fill = GOLD_FILL

    ws["H50"] = "Categoria"
    ws["I50"] = "Limite Mensal"
    for c in ("H50", "I50"):
        ws[c].font = HEADER_FONT
        ws[c].fill = HEADER_FILL

    ws["A34"] = "📅 Meses (filtros e listas)"
    ws["A34"].font = SECTION_FONT
    ws["A34"].fill = GOLD_FILL
    ws["B34"] = "Ano Inteiro"
    ws["B34"].font = _font(bold=True, color=NAVY)
    for i, m in enumerate(MONTHS_PT):
        ws.cell(row=35 + i, column=2, value=m)

    sections = [
        ("D35", "Formas de Pagamento", FORMAS_PAGAMENTO, "E"),
        ("G35", "Responsáveis", RESPONSAVEIS, "H"),
        ("J35", "Contas", CONTAS, "J"),
        ("A50", "Tipos de Receitas", RECEITAS, "B"),
        ("D50", "Tipos de Despesas", DESPESAS, "E"),
        ("A69", "Tipos de Doação", DOACAO_TIPOS, "B"),
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
            if col_letter == "H" and title == "Responsáveis":
                pass
            elif col_letter == "H" and title != "Responsáveis":
                ws.cell(row=r, column=9, value=0)
                apply_input_style(ws, r, 9)
                ws.cell(row=r, column=9).number_format = CURRENCY_FMT

    for i, cat in enumerate(GASTOS_CATEGORIAS):
        r = 51 + i
        ws.cell(row=r, column=8, value=cat).border = _border()
        ws.cell(row=r, column=9, value=0)
        apply_input_style(ws, r, 9)
        ws.cell(row=r, column=9).number_format = CURRENCY_FMT

    ws["A77"] = "💾 Dica: salve no OneDrive ou Google Drive para backup automático na nuvem."
    ws["A77"].font = _font(italic=True, color="666666")
    ws.merge_cells("A77:H77")

    set_column_widths(ws, {1: 24, 2: 16, 4: 22, 5: 18, 7: 22, 8: 22, 9: 16, 10: 14})
    ws.sheet_properties.tabColor = NAVY
    return ws


def create_cadastro_sheet(wb: Workbook):
    ws = wb.create_sheet("Cadastro Dizimistas")
    headers = ["Código", "Nome", "Telefone", "Endereço", "Aniversário", "Observações"]
    style_title_row(ws, "Cadastro de Dizimistas", len(headers), "👥")
    style_header_row(ws, headers)

    for row in range(DATA_START, DATA_END + 1):
        ws.cell(row=row, column=1).value = f"=SE(B{row}=\"\";\"\";LIN()-2)"
        apply_formula_style(ws, row, 1)
        for col in (2, 3, 4, 5, 6):
            apply_input_style(ws, row, col)
        ws.cell(row=row, column=5).number_format = DATE_FMT

    add_date_validation(ws, "E", DATA_START, DATA_END)
    add_autofilter(ws, len(headers))
    set_column_widths(ws, {1: 10, 2: 30, 3: 16, 4: 35, 5: 14, 6: 30})
    ws.sheet_properties.tabColor = GOLD
    return ws


def create_pesquisa_sheet(wb: Workbook):
    ws = wb.create_sheet("Pesquisa Dizimista")
    style_title_row(ws, "Pesquisa Rápida de Dizimista", 7, "🔎")

    ws["A3"] = "Nome do Dizimista:"
    ws["A3"].font = _font(bold=True, color=NAVY)
    apply_input_style(ws, 3, 2)
    add_list_validation(ws, "B", 3, 3, f"'Cadastro Dizimistas'!$B${DATA_START}:$B${DATA_END}")

    ws["A4"] = "Telefone:"
    ws["A4"].font = _font(bold=True, color=NAVY)
    ws["B4"] = f"=SEERRO(PROCV(B3;'Cadastro Dizimistas'!$B${DATA_START}:$C${DATA_END};2;FALSO);\"\")"
    apply_formula_style(ws, 4, 2)

    ws["A5"] = "Endereço:"
    ws["A5"].font = _font(bold=True, color=NAVY)
    ws["B5"] = f"=SEERRO(PROCV(B3;'Cadastro Dizimistas'!$B${DATA_START}:$D${DATA_END};3;FALSO);\"\")"
    apply_formula_style(ws, 5, 2)

    ws["A7"] = "Historico de dizimos:"
    ws["A7"].font = _font(bold=True, color=NAVY)
    ws["A8"] = "Use o filtro na aba Dízimos Mensais para ver os lancamentos do dizimista selecionado."
    ws["A8"].font = _font(italic=True, color="666666")
    ws.merge_cells("A8:G8")
    set_column_widths(ws, {1: 14, 2: 28, 3: 14, 4: 18, 5: 12, 6: 30, 7: 14})
    ws.sheet_properties.tabColor = GOLD
    return ws


def create_dizimos_sheet(wb: Workbook):
    ws = wb.create_sheet("Dízimos Mensais")
    headers = [
        "Data", "Nome", "Telefone", "Mês", "Valor",
        "Forma de Pagamento", "Conta", "Observação",
    ]
    style_title_row(ws, "Dízimos Mensais", len(headers), "💰")
    style_header_row(ws, headers)

    for row in range(DATA_START, DATA_END + 1):
        apply_input_style(ws, row, 1)
        ws.cell(row=row, column=1).number_format = DATE_FMT
        for col in (2, 6, 7, 8):
            apply_input_style(ws, row, col)
        ws.cell(row=row, column=3).value = (
            f'=SE(B{row}="";"";SEERRO(PROCV(B{row};\'Cadastro Dizimistas\'!$B${DATA_START}:$C${DATA_END};2;FALSO);""))'
        )
        apply_formula_style(ws, row, 3)
        ws.cell(row=row, column=4).value = (
            f'=SE(A{row}="";"";ÍNDICE(Configurações!{CFG_MESES};MÊS(A{row})))'
        )
        apply_formula_style(ws, row, 4)
        apply_input_style(ws, row, 5)
        ws.cell(row=row, column=5).number_format = CURRENCY_FMT

    add_date_validation(ws, "A", DATA_START, DATA_END)
    add_list_validation(ws, "B", DATA_START, DATA_END, f"'Cadastro Dizimistas'!$B${DATA_START}:$B${DATA_END}")
    add_list_validation(ws, "F", DATA_START, DATA_END, f"Configurações!{CFG_FORMAS}")
    add_list_validation(ws, "G", DATA_START, DATA_END, f"Configurações!{CFG_CONTAS}")
    add_autofilter(ws, len(headers))
    set_column_widths(ws, {1: 14, 2: 28, 3: 16, 4: 14, 5: 14, 6: 20, 7: 12, 8: 28})
    ws.sheet_properties.tabColor = GOLD
    return ws


def create_doacoes_sheet(wb: Workbook):
    ws = wb.create_sheet("Doações")
    headers = ["Data", "Doador", "Tipo da Doação", "Valor", "Conta", "Destinação", "Observação"]
    style_title_row(ws, "Doações", len(headers), "🎁")
    style_header_row(ws, headers)

    for row in range(DATA_START, DATA_END + 1):
        for col in (1, 2, 3, 4, 5, 6, 7):
            apply_input_style(ws, row, col)
        ws.cell(row=row, column=1).number_format = DATE_FMT
        ws.cell(row=row, column=4).number_format = CURRENCY_FMT

    add_date_validation(ws, "A", DATA_START, DATA_END)
    add_list_validation(ws, "C", DATA_START, DATA_END, f"Configurações!{CFG_DOACAO_TIPOS}")
    add_list_validation(ws, "E", DATA_START, DATA_END, f"Configurações!{CFG_CONTAS}")
    add_autofilter(ws, len(headers))
    set_column_widths(ws, {1: 14, 2: 28, 3: 18, 4: 14, 5: 12, 6: 22, 7: 28})
    ws.sheet_properties.tabColor = GOLD
    return ws


def create_missas_sheet(wb: Workbook, year: int = YEAR):
    ws = wb.create_sheet("Arrecadação das Missas")
    headers = ["Data", "Domingo", "Mês", "Valor", "Conta", "Responsável", "Observação"]
    style_title_row(ws, "Arrecadação das Missas", len(headers), "⛪")
    style_header_row(ws, headers)

    row = DATA_START
    for sunday_date, num in sundays_of_year(year):
        ws.cell(row=row, column=1, value=sunday_date)
        apply_input_style(ws, row, 1)
        ws.cell(row=row, column=1).number_format = DATE_FMT
        ws.cell(row=row, column=2, value=f"Domingo {num}")
        apply_formula_style(ws, row, 2)
        ws.cell(row=row, column=3).value = (
            f'=SE(A{row}="";"";ÍNDICE(Configurações!{CFG_MESES};MÊS(A{row})))'
        )
        apply_formula_style(ws, row, 3)
        apply_input_style(ws, row, 4)
        ws.cell(row=row, column=4).number_format = CURRENCY_FMT
        apply_input_style(ws, row, 5)
        apply_input_style(ws, row, 6)
        apply_input_style(ws, row, 7)
        row += 1

    for extra_row in range(row, DATA_END + 1):
        apply_input_style(ws, extra_row, 1)
        ws.cell(row=extra_row, column=1).number_format = DATE_FMT
        apply_input_style(ws, extra_row, 2)
        ws.cell(row=extra_row, column=3).value = (
            f'=SE(A{extra_row}="";"";ÍNDICE(Configurações!{CFG_MESES};MÊS(A{extra_row})))'
        )
        apply_formula_style(ws, extra_row, 3)
        apply_input_style(ws, extra_row, 4)
        ws.cell(row=extra_row, column=4).number_format = CURRENCY_FMT
        apply_input_style(ws, extra_row, 5)
        apply_input_style(ws, extra_row, 6)
        apply_input_style(ws, extra_row, 7)

    add_list_validation(ws, "E", DATA_START, DATA_END, f"Configurações!{CFG_CONTAS}")
    add_list_validation(ws, "F", DATA_START, DATA_END, f"Configurações!{CFG_RESP}")
    add_autofilter(ws, len(headers))
    set_column_widths(ws, {1: 14, 2: 16, 3: 14, 4: 14, 5: 12, 6: 20, 7: 28})
    ws.sheet_properties.tabColor = GOLD
    return ws


def create_gastos_sheet(wb: Workbook):
    ws = wb.create_sheet("Gastos")
    headers = [
        "Data", "Categoria", "Descrição", "Fornecedor",
        "Forma de Pagamento", "Valor", "Conta", "Responsável", "Observação",
    ]
    style_title_row(ws, "Gastos", len(headers), "💸")
    style_header_row(ws, headers)

    for row in range(DATA_START, DATA_END + 1):
        for col in range(1, 10):
            apply_input_style(ws, row, col)
        ws.cell(row=row, column=1).number_format = DATE_FMT
        ws.cell(row=row, column=6).number_format = CURRENCY_FMT

    add_date_validation(ws, "A", DATA_START, DATA_END)
    add_list_validation(ws, "B", DATA_START, DATA_END, f"Configurações!{CFG_CATEGORIAS}")
    add_list_validation(ws, "E", DATA_START, DATA_END, f"Configurações!{CFG_FORMAS}")
    add_list_validation(ws, "G", DATA_START, DATA_END, f"Configurações!{CFG_CONTAS}")
    add_list_validation(ws, "H", DATA_START, DATA_END, f"Configurações!{CFG_RESP}")
    add_autofilter(ws, len(headers))
    set_column_widths(ws, {1: 14, 2: 18, 3: 26, 4: 20, 5: 18, 6: 14, 7: 12, 8: 18, 9: 26})
    ws.sheet_properties.tabColor = GOLD
    return ws


def create_contas_sheet(wb: Workbook):
    ws = wb.create_sheet("Contas Caixa e Banco")
    style_title_row(ws, "Controle de Contas — Caixa e Banco", 6, "💳")
    headers = ["Conta", "Saldo Inicial", "Entradas", "Saídas", "Saldo Atual"]
    style_header_row(ws, headers)

    contas_cfg = [("Caixa", CFG_CAIXA), ("Banco", CFG_BANCO)]
    for i, (conta, saldo_ref) in enumerate(contas_cfg):
        row = DATA_START + i
        ws.cell(row=row, column=1, value=conta).font = _font(bold=True, color=NAVY)
        ws.cell(row=row, column=2, value=f"=Configurações!{saldo_ref}")
        apply_formula_style(ws, row, 2)
        ws.cell(row=row, column=2).number_format = CURRENCY_FMT

        ws.cell(row=row, column=3).value = (
            f'=SOMASE(\'Dízimos Mensais\'!$G${DATA_START}:$G${DATA_END};A{row};'
            f'\'Dízimos Mensais\'!$E${DATA_START}:$E${DATA_END})'
            f'+SOMASE(Doações!$E${DATA_START}:$E${DATA_END};A{row};Doações!$D${DATA_START}:$D${DATA_END})'
            f"+SOMASE('Arrecadação das Missas'!$E${DATA_START}:$E${DATA_END};A{row};"
            f"'Arrecadação das Missas'!$D${DATA_START}:$D${DATA_END})"
        )
        ws.cell(row=row, column=4).value = (
            f'=SOMASE(Gastos!$G${DATA_START}:$G${DATA_END};A{row};Gastos!$F${DATA_START}:$F${DATA_END})'
        )
        ws.cell(row=row, column=5).value = f"=B{row}+C{row}-D{row}"
        for col in range(3, 6):
            apply_formula_style(ws, row, col)
            ws.cell(row=row, column=col).number_format = CURRENCY_FMT

    total_row = DATA_START + 2
    ws.cell(row=total_row, column=1, value="TOTAL").font = _font(bold=True, color=WHITE)
    ws.cell(row=total_row, column=1).fill = TITLE_FILL
    for col in range(2, 6):
        letter = get_column_letter(col)
        ws.cell(row=total_row, column=col).value = f"=SOMA({letter}{DATA_START}:{letter}{DATA_START + 1})"
        ws.cell(row=total_row, column=col).fill = TITLE_FILL
        ws.cell(row=total_row, column=col).font = _font(bold=True, color=WHITE)
        ws.cell(row=total_row, column=col).number_format = CURRENCY_FMT

    set_column_widths(ws, {1: 14, 2: 16, 3: 16, 4: 16, 5: 16})
    ws.sheet_properties.tabColor = NAVY
    return ws


def create_resumo_sheet(wb: Workbook):
    ws = wb.create_sheet("Resumo Financeiro")
    headers = [
        "Mês", "Dízimos", "Doações", "Missas", "Total Entradas",
        "Gastos", "Saldo do Mês", "Meta", "% Meta",
    ]
    style_title_row(ws, "Resumo Financeiro Anual", len(headers), "📊")
    style_header_row(ws, headers)

    for i, month in enumerate(MONTHS_PT):
        row = DATA_START + i
        month_num = i + 1
        ws.cell(row=row, column=1, value=month)
        ws.cell(row=row, column=1).font = _font(bold=True, color=NAVY)
        ws.cell(row=row, column=1).border = _border()

        ws.cell(row=row, column=2).value = (
            f"=SOMASE('Dízimos Mensais'!$D${DATA_START}:$D${DATA_END};A{row};"
            f"'Dízimos Mensais'!$E${DATA_START}:$E${DATA_END})"
        )
        ws.cell(row=row, column=3).value = (
            f"=SOMASES(Doações!$D${DATA_START}:$D${DATA_END};"
            f"Doações!$A${DATA_START}:$A${DATA_END};\">=\"&DATA(Configurações!{CFG_ANO};{month_num};1);"
            f"Doações!$A${DATA_START}:$A${DATA_END};\"<=\"&FIM.MÊS(DATA(Configurações!{CFG_ANO};{month_num};1)))"
        )
        ws.cell(row=row, column=4).value = (
            f"=SOMASE('Arrecadação das Missas'!$C${DATA_START}:$C${DATA_END};A{row};"
            f"'Arrecadação das Missas'!$D${DATA_START}:$D${DATA_END})"
        )
        ws.cell(row=row, column=5).value = f"=SOMA(B{row}:D{row})"
        ws.cell(row=row, column=6).value = (
            f"=SOMASES(Gastos!$F${DATA_START}:$F${DATA_END};"
            f"Gastos!$A${DATA_START}:$A${DATA_END};\">=\"&DATA(Configurações!{CFG_ANO};{month_num};1);"
            f"Gastos!$A${DATA_START}:$A${DATA_END};\"<=\"&FIM.MÊS(DATA(Configurações!{CFG_ANO};{month_num};1)))"
        )
        ws.cell(row=row, column=7).value = f"=E{row}-F{row}"
        ws.cell(row=row, column=8).value = f"=ÍNDICE(Configurações!{CFG_METAS};{month_num})"
        ws.cell(row=row, column=9).value = f'=SE(H{row}=0;"";E{row}/H{row})'

        for col in range(2, 10):
            apply_formula_style(ws, row, col)
            ws.cell(row=row, column=col).number_format = CURRENCY_FMT if col != 9 else PCT_FMT
            ws.cell(row=row, column=col).alignment = RIGHT

    total_row = RESUMO_TOTAL_ROW
    ws.cell(row=total_row, column=1, value="TOTAL ANUAL")
    ws.cell(row=total_row, column=1).font = _font(bold=True, color=WHITE)
    ws.cell(row=total_row, column=1).fill = TITLE_FILL
    for col in range(2, 9):
        letter = get_column_letter(col)
        ws.cell(row=total_row, column=col).value = f"=SOMA({letter}{DATA_START}:{letter}{DATA_START + 11})"
        ws.cell(row=total_row, column=col).font = _font(bold=True, color=WHITE)
        ws.cell(row=total_row, column=col).fill = TITLE_FILL
        ws.cell(row=total_row, column=col).number_format = CURRENCY_FMT

    ws.conditional_formatting.add(
        f"F{DATA_START}:F{DATA_START + 11}",
        FormulaRule(formula=[f"F{DATA_START}>E{DATA_START}"], fill=_fill(RED_LIGHT)),
    )
    ws.conditional_formatting.add(
        f"G{DATA_START}:G{DATA_START + 11}",
        FormulaRule(formula=[f"G{DATA_START}<0"], fill=_fill(RED_LIGHT)),
    )
    ws.conditional_formatting.add(
        f"I{DATA_START}:I{DATA_START + 11}",
        FormulaRule(formula=[f"I{DATA_START}>=1"], fill=_fill(GREEN_LIGHT)),
    )

    add_autofilter(ws, len(headers))
    set_column_widths(ws, {1: 14, 2: 13, 3: 13, 4: 13, 5: 15, 6: 13, 7: 14, 8: 13, 9: 10})
    ws.sheet_properties.tabColor = NAVY
    return ws


def create_conciliacao_sheet(wb: Workbook):
    ws = wb.create_sheet("Conciliação Bancária")
    headers = ["Data", "Descrição", "Valor Planilha", "Valor Extrato", "Diferença", "Conferido"]
    style_title_row(ws, "Conciliação Bancária", len(headers), "🏦")
    style_header_row(ws, headers)

    for row in range(DATA_START, DATA_END + 1):
        apply_input_style(ws, row, 1)
        ws.cell(row=row, column=1).number_format = DATE_FMT
        apply_input_style(ws, row, 2)
        apply_input_style(ws, row, 3)
        ws.cell(row=row, column=3).number_format = CURRENCY_FMT
        apply_input_style(ws, row, 4)
        ws.cell(row=row, column=4).number_format = CURRENCY_FMT
        ws.cell(row=row, column=5).value = f"=SE(OU(A{row}=\"\";C{row}=\"\";D{row}=\"\");\"\";C{row}-D{row})"
        apply_formula_style(ws, row, 5)
        ws.cell(row=row, column=5).number_format = CURRENCY_FMT
        ws.cell(row=row, column=6).value = (
            f"=SE(OU(A{row}=\"\";C{row}=\"\";D{row}=\"\");\"\";"
            f"SE(C{row}=D{row};\"Conferido\";\"Divergente\"))"
        )
        apply_formula_style(ws, row, 6)

    add_date_validation(ws, "A", DATA_START, DATA_END)
    add_autofilter(ws, len(headers))
    ws.conditional_formatting.add(
        f"F{DATA_START}:F{DATA_END}",
        FormulaRule(formula=[f'F{DATA_START}="Divergente"'], fill=_fill(RED_LIGHT)),
    )
    set_column_widths(ws, {1: 14, 2: 35, 3: 16, 4: 16, 5: 14, 6: 12})
    ws.sheet_properties.tabColor = NAVY
    return ws


def create_relatorio_anual_sheet(wb: Workbook):
    ws = wb.create_sheet("Relatório Anual")
    style_title_row(ws, "Relatório Anual — Prestação de Contas", 4, "📑")
    tr = RESUMO_TOTAL_ROW

    ws["A3"] = "RECEITAS"
    ws["A3"].font = _font(bold=True, size=14, color=WHITE)
    ws["A3"].fill = _fill("2E7D32")
    ws.merge_cells("A3:B3")

    receitas = [
        ("Total de Dízimos", f"='Resumo Financeiro'!B{tr}"),
        ("Total de Doações", f"='Resumo Financeiro'!C{tr}"),
        ("Total Coletas de Domingo (Missas)", f"='Resumo Financeiro'!D{tr}"),
        ("Outras Receitas", "0"),
        ("TOTAL DE RECEITAS", f"='Resumo Financeiro'!E{tr}"),
    ]
    for i, (label, formula) in enumerate(receitas, 4):
        ws.cell(row=i, column=1, value=label).font = _font(bold=("TOTAL" in label), color=NAVY)
        ws.cell(row=i, column=2, value=formula)
        apply_formula_style(ws, i, 2)
        ws.cell(row=i, column=2).number_format = CURRENCY_FMT
        if label == "Outras Receitas":
            apply_input_style(ws, i, 2)

    ws["A10"] = "DESPESAS"
    ws["A10"].font = _font(bold=True, size=14, color=WHITE)
    ws["A10"].fill = _fill("C62828")
    ws.merge_cells("A10:B10")

    for i, cat in enumerate(GASTOS_CATEGORIAS):
        row = 11 + i
        ws.cell(row=row, column=1, value=cat)
        ws.cell(row=row, column=2).value = (
            f'=SOMASE(Gastos!$B${DATA_START}:$B${DATA_END};A{row};Gastos!$F${DATA_START}:$F${DATA_END})'
        )
        apply_formula_style(ws, row, 2)
        ws.cell(row=row, column=2).number_format = CURRENCY_FMT

    despesas_total_row = 11 + len(GASTOS_CATEGORIAS)
    ws.cell(row=despesas_total_row, column=1, value="TOTAL DE DESPESAS").font = _font(bold=True, color=NAVY)
    ws.cell(row=despesas_total_row, column=2).value = f"='Resumo Financeiro'!F{tr}"
    apply_formula_style(ws, despesas_total_row, 2)
    ws.cell(row=despesas_total_row, column=2).number_format = CURRENCY_FMT

    result_row = despesas_total_row + 2
    ws.merge_cells(start_row=result_row, start_column=1, end_row=result_row, end_column=2)
    ws.cell(row=result_row, column=1, value="RESULTADO DO EXERCÍCIO").font = _font(bold=True, size=14, color=WHITE)
    ws.cell(row=result_row, column=1).fill = TITLE_FILL

    results = [
        ("Saldo Inicial", f"=Configurações!{CFG_SALDO_TOTAL}"),
        ("(+) Total Receitas", f"='Resumo Financeiro'!E{tr}"),
        ("(-) Total Despesas", f"='Resumo Financeiro'!F{tr}"),
        ("(=) Saldo Final", f"=Configurações!{CFG_SALDO_TOTAL}+'Resumo Financeiro'!G{tr}"),
    ]
    for i, (label, formula) in enumerate(results, result_row + 1):
        ws.cell(row=i, column=1, value=label).font = _font(bold="Saldo Final" in label)
        ws.cell(row=i, column=2, value=formula)
        apply_formula_style(ws, i, 2)
        ws.cell(row=i, column=2).number_format = CURRENCY_FMT

    ws.cell(row=result_row + 6, column=1, value="Ano de referência:")
    ws.cell(row=result_row + 6, column=2, value=f"=Configurações!{CFG_ANO}")
    setup_print_a4(ws, "Relatório Anual")
    set_column_widths(ws, {1: 32, 2: 18})
    ws.sheet_properties.tabColor = NAVY
    return ws


def create_relatorio_despesas_sheet(wb: Workbook):
    ws = wb.create_sheet("Relatório Despesas")
    style_title_row(ws, "Relatório por Categoria de Despesas", 5, "📑")
    headers = ["Categoria", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
               "Jul", "Ago", "Set", "Out", "Nov", "Dez", "Total Anual", "% do Total"]
    style_header_row(ws, headers)

    total_desp_row = DATA_START + len(GASTOS_CATEGORIAS)
    for i, cat in enumerate(GASTOS_CATEGORIAS):
        row = DATA_START + i
        ws.cell(row=row, column=1, value=cat).font = _font(bold=True, color=NAVY)
        for m in range(1, 13):
            col = 1 + m
            ws.cell(row=row, column=col).value = (
                f"=SOMASES(Gastos!$F${DATA_START}:$F${DATA_END};Gastos!$B${DATA_START}:$B${DATA_END};$A{row};"
                f"Gastos!$A${DATA_START}:$A${DATA_END};\">=\"&DATA(Configurações!{CFG_ANO};{m};1);"
                f"Gastos!$A${DATA_START}:$A${DATA_END};\"<=\"&FIM.MÊS(DATA(Configurações!{CFG_ANO};{m};1)))"
            )
            apply_formula_style(ws, row, col)
            ws.cell(row=row, column=col).number_format = CURRENCY_FMT
        ws.cell(row=row, column=14).value = f"=SOMA(B{row}:M{row})"
        apply_formula_style(ws, row, 14)
        ws.cell(row=row, column=14).number_format = CURRENCY_FMT
        ws.cell(row=row, column=15).value = f"=SE(N${total_desp_row}=0;\"\";N{row}/N${total_desp_row})"
        apply_formula_style(ws, row, 15)
        ws.cell(row=row, column=15).number_format = PCT_FMT

    ws.cell(row=total_desp_row, column=1, value="TOTAL").font = _font(bold=True, color=WHITE)
    ws.cell(row=total_desp_row, column=1).fill = TITLE_FILL
    for col in range(2, 15):
        letter = get_column_letter(col)
        ws.cell(row=total_desp_row, column=col).value = f"=SOMA({letter}{DATA_START}:{letter}{total_desp_row - 1})"
        ws.cell(row=total_desp_row, column=col).fill = TITLE_FILL
        ws.cell(row=total_desp_row, column=col).font = _font(bold=True, color=WHITE)
        ws.cell(row=total_desp_row, column=col).number_format = CURRENCY_FMT if col < 15 else PCT_FMT

    setup_print_a4(ws)
    set_column_widths(ws, {1: 18, 14: 14, 15: 10})
    ws.sheet_properties.tabColor = NAVY
    return ws


def create_comparativo_sheet(wb: Workbook):
    ws = wb.create_sheet("Comparativo Anual")
    style_title_row(ws, "Comparativo entre Anos", 5, "📈")
    years = [YEAR - 2, YEAR - 1, YEAR]
    headers = ["Indicador"] + [str(y) for y in years] + ["Variação %"]
    style_header_row(ws, headers)

    indicators = [
        ("Dízimos", "B"),
        ("Doações", "C"),
        ("Missas", "D"),
        ("Total Entradas", "E"),
        ("Total Gastos", "F"),
        ("Resultado", "G"),
    ]
    tr = RESUMO_TOTAL_ROW
    for i, (label, col_letter) in enumerate(indicators):
        row = DATA_START + i
        ws.cell(row=row, column=1, value=label).font = _font(bold=True, color=NAVY)
        for j, year in enumerate(years):
            col = 2 + j
            if year == YEAR:
                ws.cell(row=row, column=col).value = f"='Resumo Financeiro'!{col_letter}{tr}"
                apply_formula_style(ws, row, col)
            else:
                ws.cell(row=row, column=col, value=0)
                apply_input_style(ws, row, col)
            ws.cell(row=row, column=col).number_format = CURRENCY_FMT
        ws.cell(row=row, column=5).value = (
            f'=SE(C{row}=0;"";(D{row}-C{row})/ABS(C{row}))'
        )
        apply_formula_style(ws, row, 5)
        ws.cell(row=row, column=5).number_format = PCT_FMT

    ws["A11"] = "ℹ️ Preencha manualmente os anos anteriores. O ano atual é calculado automaticamente."
    ws["A11"].font = _font(italic=True, color="666666")
    ws.merge_cells("A11:E11")
    set_column_widths(ws, {1: 20, 2: 14, 3: 14, 4: 14, 5: 14})
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
    for i, receita in enumerate(RECEITAS, 1):
        ws.cell(row=4 + i, column=1, value=f"R{i:02d}")
        ws.cell(row=4 + i, column=2, value=receita)
    for i, despesa in enumerate(DESPESAS, 1):
        ws.cell(row=4 + i, column=4, value=f"D{i:02d}")
        ws.cell(row=4 + i, column=5, value=despesa)
    set_column_widths(ws, {1: 10, 2: 24, 4: 10, 5: 24})
    ws.sheet_properties.tabColor = NAVY
    return ws


def create_aux_fluxo_sheet(wb: Workbook):
    ws = wb.create_sheet("_Consolidado")
    row_out = 2
    sources = [
        ("Dízimos Mensais", "A", "B", "E", None, "Dízimo - ", 5),
        ("Doações", "A", "B", "D", None, "Doação - ", 4),
        ("Arrecadação das Missas", "A", "B", "D", None, "Missas - ", 4),
        ("Gastos", "A", "B", "F", "C", "", 6),
    ]
    for sheet, col_d, col_desc, col_val, col_desc2, prefix, _ in sources:
        for r in range(DATA_START, DATA_END + 1):
            s = f"'{sheet}'" if " " in sheet else sheet
            ws.cell(row=row_out, column=1).value = f"=SE({s}!{col_d}{r}=\"\";\"\";{s}!{col_d}{r})"
            if col_desc2:
                ws.cell(row=row_out, column=2).value = (
                    f"=SE({s}!{col_d}{r}=\"\";\"\";{s}!B{r}&\" - \"&{s}!{col_desc2}{r})"
                )
            else:
                ws.cell(row=row_out, column=2).value = (
                    f"=SE({s}!{col_d}{r}=\"\";\"\";\"{prefix}\"&{s}!{col_desc}{r})"
                )
            if sheet == "Gastos":
                ws.cell(row=row_out, column=3).value = f"=SE({s}!{col_d}{r}=\"\";\"\";0)"
                ws.cell(row=row_out, column=4).value = f"=SE({s}!{col_d}{r}=\"\";\"\";{s}!{col_val}{r})"
            else:
                ws.cell(row=row_out, column=3).value = f"=SE({s}!{col_d}{r}=\"\";\"\";{s}!{col_val}{r})"
                ws.cell(row=row_out, column=4).value = f"=SE({s}!{col_d}{r}=\"\";\"\";0)"
            row_out += 1
    ws.sheet_state = "hidden"
    return ws


def create_fluxo_sheet(wb: Workbook):
    ws = wb.create_sheet("Fluxo de Caixa")
    headers = ["Data", "Descrição", "Entrada", "Saída", "Saldo"]
    style_title_row(ws, "Fluxo de Caixa", len(headers), "💵")
    style_header_row(ws, headers)

    ws.cell(row=DATA_START, column=2, value="Saldo Inicial (Caixa + Banco)")
    ws.cell(row=DATA_START, column=2).font = _font(bold=True, color=NAVY)
    ws.cell(row=DATA_START, column=5, value=f"=Configurações!{CFG_SALDO_TOTAL}")
    apply_formula_style(ws, DATA_START, 5)
    ws.cell(row=DATA_START, column=5).number_format = CURRENCY_FMT

    fluxo_start = DATA_START + 1
    for i in range(500):
        fluxo_row = fluxo_start + i
        aux_row = 2 + i
        for col, aux_col in enumerate(range(1, 5), 1):
            ws.cell(row=fluxo_row, column=col).value = f"=_Consolidado!{get_column_letter(aux_col)}{aux_row}"
            apply_formula_style(ws, fluxo_row, col)
        prev = fluxo_row - 1
        ws.cell(row=fluxo_row, column=5).value = (
            f'=SE(A{fluxo_row}="";E{prev};E{prev}+SEERRO(C{fluxo_row};0)-SEERRO(D{fluxo_row};0))'
        )
        apply_formula_style(ws, fluxo_row, 5)
        ws.cell(row=fluxo_row, column=1).number_format = DATE_FMT
        for col in (3, 4, 5):
            ws.cell(row=fluxo_row, column=col).number_format = CURRENCY_FMT

    setup_print_a4(ws, "Fluxo de Caixa")
    add_autofilter(ws, len(headers))
    set_column_widths(ws, {1: 14, 2: 40, 3: 14, 4: 14, 5: 16})
    ws.sheet_properties.tabColor = NAVY
    return ws


def create_impressao_sheet(wb: Workbook):
    ws = wb.create_sheet("Impressão A4")
    style_title_row(ws, "Relatórios para Impressão (A4)", 6, "🖨️")
    ws["A3"] = "Selecione o mês para relatórios mensais:"
    ws["A3"].font = _font(bold=True, color=NAVY)
    ws["B3"] = "Janeiro"
    apply_input_style(ws, 3, 2)
    add_list_validation(ws, "B", 3, 3, f"Configurações!{CFG_MESES}")

    blocks = [
        (5, "PRESTAÇÃO DE CONTAS MENSAL", "=B3&\" / \"&Configurações!$B$3"),
        (45, "PRESTAÇÃO DE CONTAS ANUAL", "=Configurações!$B$3"),
        (85, "RESUMO FINANCEIRO", "=B3"),
        (120, "DEMONSTRATIVO DE RECEITAS E DESPESAS", "=Configurações!$B$3"),
    ]
    for start_row, titulo, periodo in blocks:
        print_report_header(ws, start_row, titulo, periodo, 6)
        hdr = start_row + 4
        if "MENSAL" in titulo:
            ws.cell(row=hdr, column=1, value="Item")
            ws.cell(row=hdr, column=2, value="Valor")
            items = [
                ("Dízimos", f"=ÍNDICE('Resumo Financeiro'!B{DATA_START}:B{DATA_START+11};CORRESP(B3;'Resumo Financeiro'!A{DATA_START}:A{DATA_START+11};0))"),
                ("Doações", f"=ÍNDICE('Resumo Financeiro'!C{DATA_START}:C{DATA_START+11};CORRESP(B3;'Resumo Financeiro'!A{DATA_START}:A{DATA_START+11};0))"),
                ("Missas", f"=ÍNDICE('Resumo Financeiro'!D{DATA_START}:D{DATA_START+11};CORRESP(B3;'Resumo Financeiro'!A{DATA_START}:A{DATA_START+11};0))"),
                ("Total Entradas", f"=ÍNDICE('Resumo Financeiro'!E{DATA_START}:E{DATA_START+11};CORRESP(B3;'Resumo Financeiro'!A{DATA_START}:A{DATA_START+11};0))"),
                ("Gastos", f"=ÍNDICE('Resumo Financeiro'!F{DATA_START}:F{DATA_START+11};CORRESP(B3;'Resumo Financeiro'!A{DATA_START}:A{DATA_START+11};0))"),
                ("Saldo do Mês", f"=ÍNDICE('Resumo Financeiro'!G{DATA_START}:G{DATA_START+11};CORRESP(B3;'Resumo Financeiro'!A{DATA_START}:A{DATA_START+11};0))"),
            ]
            for j, (item, formula) in enumerate(items, hdr + 1):
                ws.cell(row=j, column=1, value=item)
                ws.cell(row=j, column=2, value=formula)
                ws.cell(row=j, column=2).number_format = CURRENCY_FMT
            add_signature_block(ws, hdr + 9)
        elif "ANUAL" in titulo and "MENSAL" not in titulo:
            tr = RESUMO_TOTAL_ROW
            items = [
                ("Total Receitas", f"='Resumo Financeiro'!E{tr}"),
                ("Total Despesas", f"='Resumo Financeiro'!F{tr}"),
                ("Resultado", f"='Resumo Financeiro'!G{tr}"),
                ("Saldo Final", f"=Configurações!{CFG_SALDO_TOTAL}+'Resumo Financeiro'!G{tr}"),
            ]
            for j, (item, formula) in enumerate(items, hdr + 1):
                ws.cell(row=j, column=1, value=item)
                ws.cell(row=j, column=2, value=formula)
                ws.cell(row=j, column=2).number_format = CURRENCY_FMT
            add_signature_block(ws, hdr + 7)
        else:
            ws.cell(row=hdr, column=1, value="Consulte as abas Relatório Anual e Relatório Despesas para detalhes.")
            ws.merge_cells(start_row=hdr, start_column=1, end_row=hdr, end_column=6)
            add_signature_block(ws, hdr + 3)

    setup_print_a4(ws, "Prestação de Contas")
    set_column_widths(ws, {1: 28, 2: 18, 3: 14, 4: 14, 5: 14, 6: 14})
    ws.sheet_properties.tabColor = GOLD
    return ws


def create_novo_exercicio_sheet(wb: Workbook):
    ws = wb.create_sheet("Novo Exercício")
    style_title_row(ws, "Iniciar Novo Exercício Financeiro", 4, "🔄")

    ws.merge_cells("A3:D3")
    ws["A3"] = "▶  NOVO EXERCÍCIO"
    ws["A3"].font = _font(bold=True, size=16, color=WHITE)
    ws["A3"].fill = _fill("2E7D32")
    ws["A3"].alignment = CENTER
    ws.row_dimensions[3].height = 36

    instructions = [
        "Este botão/macro executa automaticamente:",
        "  • Limpa lançamentos do ano anterior (Dízimos, Doações, Missas, Gastos, Conciliação)",
        "  • Mantém fórmulas, gráficos, Dashboard e configurações de listas",
        "  • Incrementa o ano (ex: 2026 → 2027)",
        "  • Preenche domingos do novo ano na aba Missas",
        "  • Arquiva totais do ano no Comparativo Anual",
        "",
        "COMO ATIVAR O BOTÃO:",
        "1. Salve a planilha como .xlsm (Excel com macros)",
        "2. Pressione Alt+F11 → Arquivo → Importar → vba/NovoExercicio.bas",
        "3. Desenvolvedor → Inserir → Botão → associe à macro NovoExercicio",
        "",
        "Alternativa sem macro: execute no terminal:",
        "  python gerar_planilha.py --novo-exercicio",
    ]
    for i, line in enumerate(instructions, 5):
        ws.cell(row=i, column=1, value=line)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=4)

    ws.sheet_properties.tabColor = "2E7D32"
    return ws


def create_dados_dashboard_sheet(wb: Workbook):
    """Aba auxiliar com dados dos graficos (evita colunas ocultas no Dashboard)."""
    ws = wb.create_sheet("_DadosDashboard")
    headers = ["Mes", "Entradas", "Gastos", "Saldo Acum."]
    for col, h in enumerate(headers, 1):
        ws.cell(row=2, column=col, value=h)

    for i, month in enumerate(MONTHS_PT):
        r = DATA_START + i
        ws.cell(row=r, column=1, value=month[:3])
        ws.cell(row=r, column=2).value = f"='Resumo Financeiro'!E{r}"
        ws.cell(row=r, column=3).value = f"='Resumo Financeiro'!F{r}"
        ws.cell(row=r, column=4).value = (
            f"=Configurações!{CFG_SALDO_TOTAL}+SOMA($B${DATA_START}:B{r})-SOMA($C${DATA_START}:C{r})"
        )

    pie_row = DATA_START + 14
    ws.cell(row=pie_row, column=1, value="Receita")
    ws.cell(row=pie_row, column=2, value="Total")
    for i, label in enumerate(["Dizimos", "Doacoes", "Missas"]):
        r = pie_row + 1 + i
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2).value = f"='Resumo Financeiro'!{chr(ord('B') + i)}{RESUMO_TOTAL_ROW}"

    ws.cell(row=2, column=6, value="Categoria")
    ws.cell(row=2, column=7, value="Total")
    for i, cat in enumerate(GASTOS_CATEGORIAS):
        r = DATA_START + i
        ws.cell(row=r, column=6, value=cat)
        ws.cell(row=r, column=7).value = (
            f'=SOMASE(Gastos!$B${DATA_START}:$B${DATA_END};F{r};Gastos!$F${DATA_START}:$F${DATA_END})'
        )

    ws.sheet_state = "hidden"
    return ws


def create_dashboard_sheet(wb: Workbook):
    ws = wb.create_sheet("Dashboard")
    style_title_row(ws, "Tesouraria — Dashboard Financeiro", 10, "📑")

    ws["A2"] = "Filtrar Mês:"
    ws["A2"].font = _font(bold=True, color=NAVY)
    ws["B2"] = "Ano Inteiro"
    apply_input_style(ws, 2, 2)
    add_list_validation(ws, "B", 2, 2, f"Configurações!{CFG_FILTRO_MESES}")

    ws["D2"] = "Mês Atual:"
    ws["E2"] = f"=ÍNDICE(Configurações!{CFG_MESES};MÊS(HOJE()))"
    apply_formula_style(ws, 2, 5)

    kpi_row = 4
    ws.merge_cells(start_row=kpi_row, start_column=4, end_row=kpi_row + 1, end_column=7)
    ws.cell(row=kpi_row, column=4, value="🟢 SALDO ATUAL").font = _font(bold=True, size=12, color=NAVY)
    ws.cell(row=kpi_row, column=4).alignment = CENTER

    ws.merge_cells(start_row=kpi_row + 2, start_column=4, end_row=kpi_row + 4, end_column=7)
    saldo = ws.cell(row=kpi_row + 2, column=4)
    saldo.value = (
        f"=Configurações!{CFG_SALDO_TOTAL}+SOMA('Resumo Financeiro'!E{DATA_START}:E{DATA_START + 11})"
        f"-SOMA('Resumo Financeiro'!F{DATA_START}:F{DATA_START + 11})"
    )
    saldo.font = KPI_VALUE_FONT
    saldo.alignment = CENTER
    saldo.number_format = CURRENCY_FMT
    saldo.fill = GOLD_FILL

    tr = RESUMO_TOTAL_ROW
    mes_idx = f'CORRESP(E2;\'Resumo Financeiro\'!A{DATA_START}:A{DATA_START + 11};0)'

    kpis = [
        (1, kpi_row, "💰 Total Entradas", f"='Resumo Financeiro'!E{tr}"),
        (9, kpi_row, "💸 Total Gastos", f"='Resumo Financeiro'!F{tr}"),
        (1, kpi_row + 5, "📈 Resultado Ano", f"='Resumo Financeiro'!G{tr}"),
        (9, kpi_row + 5, "📅 Entradas Mês", f"=ÍNDICE('Resumo Financeiro'!E{DATA_START}:E{DATA_START + 11};MÊS(HOJE()))"),
        (1, kpi_row + 10, "📅 Gastos Mês", f"=ÍNDICE('Resumo Financeiro'!F{DATA_START}:F{DATA_START + 11};MÊS(HOJE()))"),
        (9, kpi_row + 10, "🎯 Meta Mês (%)", f"=ÍNDICE('Resumo Financeiro'!I{DATA_START}:I{DATA_START + 11};MÊS(HOJE()))"),
    ]
    for col, row, label, formula in kpis:
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
        ws.cell(row=row, column=col, value=label).font = KPI_FONT
        ws.cell(row=row, column=col).fill = _fill(LIGHT_BLUE)
        ws.cell(row=row, column=col).alignment = CENTER
        ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + 1)
        val = ws.cell(row=row + 1, column=col, value=formula)
        val.font = _font(bold=True, size=14, color=NAVY)
        val.number_format = CURRENCY_FMT if "%" not in label else PCT_FMT
        val.alignment = CENTER

    ws.merge_cells("A16:J16")
    ws["A16"] = "⚠️ ALERTAS INTELIGENTES"
    ws["A16"].font = SECTION_FONT
    ws["A16"].fill = GOLD_FILL

    alerts = [
        (
            17,
            f'=SE(ÍNDICE(\'Resumo Financeiro\'!F{DATA_START}:F{DATA_START + 11};MÊS(HOJE()))>'
            f'ÍNDICE(\'Resumo Financeiro\'!E{DATA_START}:E{DATA_START + 11};MÊS(HOJE()));'
            f'"ALERTA: Gastos maiores que entradas do mes";"OK: Entradas cobrem os gastos do mes")',
        ),
        (
            18,
            f'=SE({saldo.coordinate}<0;"ALERTA: Saldo negativo";"OK: Saldo positivo")',
        ),
        (
            19,
            f'=SE(ÍNDICE(\'Resumo Financeiro\'!I{DATA_START}:I{DATA_START + 11};MÊS(HOJE()))>=1;'
            f'"OK: Meta mensal atingida";'
            f'"Meta mensal ainda nao atingida")',
        ),
        (
            20,
            '="Verifique limites de gastos na aba Configuracoes"',
        ),
    ]
    for row, formula in alerts:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
        ws.cell(row=row, column=1, value=formula)
        ws.cell(row=row, column=1).font = _font(size=11)

    ws.merge_cells("A22:J22")
    ws["A22"] = "🏦 CONCILIAÇÃO BANCÁRIA"
    ws["A22"].font = SECTION_FONT
    ws["A22"].fill = _fill(LIGHT_BLUE)

    conc_kpis = [
        (23, "Total Conciliado", f'=SOMASE(\'Conciliação Bancária\'!F{DATA_START}:F{DATA_END};"Conferido";\'Conciliação Bancária\'!C{DATA_START}:C{DATA_END})'),
        (24, "Total Pendente", f"=SOMA('Conciliação Bancária'!C{DATA_START}:C{DATA_END})-C23"),
        (25, "Diferenças Encontradas", f"=SOMA('Conciliação Bancária'!E{DATA_START}:E{DATA_END})"),
    ]
    for row, label, formula in conc_kpis:
        ws.cell(row=row, column=1, value=label).font = _font(bold=True, color=NAVY)
        ws.cell(row=row, column=3, value=formula)
        apply_formula_style(ws, row, 3)
        ws.cell(row=row, column=3).number_format = CURRENCY_FMT

    data_ws = wb["_DadosDashboard"]
    cats1 = Reference(data_ws, min_col=1, min_row=DATA_START, max_row=DATA_START + 11)

    chart1 = BarChart()
    chart1.type = "col"
    chart1.grouping = "clustered"
    chart1.title = "Entradas x Gastos por Mes"
    chart1.style = 10
    chart1.width = 16
    chart1.height = 9
    data1 = Reference(data_ws, min_col=2, min_row=2, max_col=3, max_row=DATA_START + 11)
    chart1.add_data(data1, titles_from_data=True)
    chart1.set_categories(cats1)
    ws.add_chart(chart1, "A27")

    chart2 = LineChart()
    chart2.title = "Evolucao do Saldo"
    chart2.style = 10
    chart2.width = 16
    chart2.height = 9
    data2 = Reference(data_ws, min_col=4, min_row=2, max_row=DATA_START + 11)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats1)
    ws.add_chart(chart2, "J27")

    pie_row = DATA_START + 14
    chart3 = PieChart()
    chart3.title = "Origem das Receitas"
    chart3.width = 14
    chart3.height = 9
    chart3.add_data(Reference(data_ws, min_col=2, min_row=pie_row + 1, max_row=pie_row + 3))
    chart3.set_categories(Reference(data_ws, min_col=1, min_row=pie_row + 1, max_row=pie_row + 3))
    chart3.dataLabels = DataLabelList()
    chart3.dataLabels.showPercent = True
    ws.add_chart(chart3, "A42")

    chart4 = PieChart()
    chart4.title = "Gastos por Categoria"
    chart4.width = 14
    chart4.height = 9
    chart4.add_data(Reference(data_ws, min_col=7, min_row=DATA_START, max_row=DATA_START + len(GASTOS_CATEGORIAS) - 1))
    chart4.set_categories(Reference(data_ws, min_col=6, min_row=DATA_START, max_row=DATA_START + len(GASTOS_CATEGORIAS) - 1))
    chart4.dataLabels = DataLabelList()
    chart4.dataLabels.showPercent = True
    ws.add_chart(chart4, "J42")

    ws.conditional_formatting.add(
        saldo.coordinate,
        FormulaRule(formula=[f"{saldo.coordinate}<0"], fill=_fill(RED_LIGHT)),
    )
    ws.sheet_properties.tabColor = GOLD
    return ws


def protect_sheets(wb: Workbook):
    for ws in wb.worksheets:
        if ws.title.startswith("_"):
            continue
        ws.protection.sheet = True
        ws.protection.sort = True
        ws.protection.autoFilter = True
        ws.protection.selectLockedCells = True
        ws.protection.selectUnlockedCells = True


SHEET_ORDER = [
    "Dashboard",
    "Cadastro Dizimistas",
    "Pesquisa Dizimista",
    "Dízimos Mensais",
    "Doações",
    "Arrecadação das Missas",
    "Gastos",
    "Contas Caixa e Banco",
    "Resumo Financeiro",
    "Conciliação Bancária",
    "Relatório Anual",
    "Relatório Despesas",
    "Comparativo Anual",
    "Plano de Contas",
    "Fluxo de Caixa",
    "Impressão A4",
    "Configurações",
    "Novo Exercício",
    "_DadosDashboard",
    "_Consolidado",
]


def generate(output_path: Path, year: int = YEAR) -> Path:
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    create_config_sheet(wb, year)
    create_cadastro_sheet(wb)
    create_dizimos_sheet(wb)
    create_doacoes_sheet(wb)
    create_missas_sheet(wb, year)
    create_gastos_sheet(wb)
    create_contas_sheet(wb)
    create_resumo_sheet(wb)
    create_conciliacao_sheet(wb)
    create_relatorio_anual_sheet(wb)
    create_relatorio_despesas_sheet(wb)
    create_comparativo_sheet(wb)
    create_plano_contas_sheet(wb)
    create_aux_fluxo_sheet(wb)
    create_fluxo_sheet(wb)
    create_impressao_sheet(wb)
    create_novo_exercicio_sheet(wb)
    create_pesquisa_sheet(wb)
    create_dados_dashboard_sheet(wb)
    create_dashboard_sheet(wb)

    for target_idx, name in enumerate(SHEET_ORDER):
        current_idx = wb.sheetnames.index(name)
        if current_idx != target_idx:
            wb.move_sheet(wb[name], target_idx - current_idx)

    protect_sheets(wb)
    wb.active = wb["Dashboard"]
    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Gerar planilha de tesouraria")
    parser.add_argument("--novo-exercicio", action="store_true", help="Incrementa ano e regenera")
    parser.add_argument("--output", type=Path, default=Path(__file__).parent / "Tesouraria_Igreja.xlsx")
    args = parser.parse_args()

    year = YEAR + 1 if args.novo_exercicio else YEAR

    path = generate(args.output, year=year)
    print(f"Planilha gerada: {path}" + (f" (exercício {year})" if args.novo_exercicio else ""))
