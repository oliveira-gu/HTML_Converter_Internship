#!/usr/bin/env python3
"""
xlsx_reader.py -- Extrai runs de texto rico de uma coluna Excel (.xlsx).

O openpyxl representa o conteúdo de uma célula como:
  - Uma string Python simples / número (sem formatação por caractere).
  - Um objeto CellRichText: lista de TextBlock / InlineFont, onde cada
    TextBlock carrega texto + um InlineFont com flags bold/italic/underline.

Este módulo normaliza ambos os casos em uma lista de parágrafos, onde cada
parágrafo é uma lista de Run namedtuples (definidos em converter.py).
Essa estrutura é o que converter.convert_runs() espera.

Um "parágrafo" aqui = célula de uma linha. Quebras de linha Alt+Enter
dentro de uma célula (\\n na string) são mantidas; convert_runs() as divide
e classifica como quebras <br> dentro do parágrafo.

Requer: openpyxl >= 3.1  (pip install openpyxl)
"""

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock

from converter import Run


# ---------------------------------------------------------------------------
# API pública
# ---------------------------------------------------------------------------

def read_column(path, column_name, sheet_name=None, header_row=1):
    """Lê uma coluna de um arquivo .xlsx e retorna dados estruturados de runs.

    Parâmetros
    ----------
    path : str ou Path
        Caminho para o arquivo .xlsx.
    column_name : str
        Nome do cabeçalho da coluna a extrair (correspondência sem distinção de maiúsculas).
    sheet_name : str ou None
        Planilha a ler. None = planilha ativa (primeira).
    header_row : int
        Número da linha (base 1) onde ficam os cabeçalhos. Padrão: 1.

    Retorna
    -------
    lista de dicts com chaves:
        'row'        : número da linha na planilha (base 1)
        'paragraphs' : lista de parágrafo (lista de Run)
        'raw'        : valor original da célula conforme retornado pelo openpyxl

    Exceções
    --------
    ValueError      se o cabeçalho de coluna não for encontrado.
    FileNotFoundError  se o caminho não existir.
    """
    wb = openpyxl.load_workbook(path, rich_text=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    col_index = _find_column(ws, column_name, header_row)

    results = []
    for row_idx in range(header_row + 1, ws.max_row + 1):
        cell = ws.cell(row=row_idx, column=col_index)
        if cell.value is None:
            continue
        paragraphs = _cell_to_paragraphs(cell)
        if not any(_runs_have_text(p) for p in paragraphs):
            continue
        results.append({
            "row": row_idx,
            "paragraphs": paragraphs,
            "raw": cell.value,
        })

    wb.close()
    return results


def read_all_columns(path, sheet_name=None, header_row=1):
    """Retorna os nomes dos cabeçalhos encontrados na planilha (útil para descoberta)."""
    wb = openpyxl.load_workbook(path, rich_text=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    headers = [
        ws.cell(row=header_row, column=c).value
        for c in range(1, ws.max_column + 1)
    ]
    wb.close()
    return [h for h in headers if h is not None]


# ---------------------------------------------------------------------------
# Funções auxiliares internas
# ---------------------------------------------------------------------------

def _find_column(ws, column_name, header_row):
    """Retorna o índice de coluna (base 1) para column_name (sem distinção de maiúsculas)."""
    target = column_name.strip().lower()
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col).value
        if val is not None and str(val).strip().lower() == target:
            return col
    available = [
        ws.cell(row=header_row, column=c).value
        for c in range(1, ws.max_column + 1)
    ]
    raise ValueError(
        "Coluna %r não encontrada na linha de cabeçalho %d. "
        "Colunas disponíveis: %s"
        % (column_name, header_row, [h for h in available if h is not None])
    )


def _cell_to_paragraphs(cell):
    """Converte valor de célula em lista de parágrafos (cada um com lista de Run)."""
    value = cell.value

    if isinstance(value, CellRichText):
        return [_rich_text_to_runs(value)]

    text = str(value)
    font = cell.font or _NullFont()
    return [[Run(text, bool(font.bold), bool(font.italic), bool(font.underline))]]


def _rich_text_to_runs(rich_text):
    """Converte objeto CellRichText em lista de Run namedtuples."""
    runs = []
    for element in rich_text:
        if isinstance(element, TextBlock):
            font = element.font or _NullFont()
            runs.append(Run(
                text=element.text,
                bold=bool(font.bold),
                italic=bool(font.italic),
                underline=bool(font.underline),
            ))
        elif isinstance(element, str):
            runs.append(Run(text=element))
    return runs


def _runs_have_text(paragraph):
    return any(r.text.strip() for r in paragraph)


class _NullFont:
    """Sentinela que torna o acesso a atributos seguro quando cell.font é None."""
    bold = False
    italic = False
    underline = False
