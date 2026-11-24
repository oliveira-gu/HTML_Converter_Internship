#!/usr/bin/env python3
"""
batch.py -- Conversão em lote de uma coluna de descrições para HTML em linha única.

Aceita arquivos .csv ou .xlsx e detecta o formato automaticamente pela extensão.
Converte a coluna-alvo e grava um novo CSV.

  Caminho .xlsx  — lê runs de texto rico do Excel (negrito/itálico/sublinhado
                   preservados) via xlsx_reader + converter.convert_runs().
                   Requer openpyxl.
  Caminho .csv   — lê texto simples (CSV não carrega metadados de formatação);
                   converte com converter.convert() (listas, quebras, URLs).

Uso
---
    # A partir de um arquivo CSV:
    py -3 batch.py --input dados.csv --output resultado.csv --column "Descrição"

    # A partir de um arquivo Excel (negrito/itálico/sublinhado preservados):
    py -3 batch.py --input dados.xlsx --output resultado.csv --column "Descrição"

    # Especificar nome da planilha se o .xlsx tiver múltiplas abas:
    py -3 batch.py --input dados.xlsx --output resultado.csv --column "Descrição" --sheet "Planilha1"

    # Substituir a linha de cabeçalho se os cabeçalhos não estiverem na linha 1:
    py -3 batch.py --input dados.xlsx --output resultado.csv --column "Descrição" --header-row 2

    # Descobrir nomes de colunas sem converter:
    py -3 batch.py --input dados.xlsx --list-columns
    py -3 batch.py --input dados.csv --list-columns

Requisitos
----------
    openpyxl é necessário apenas para entrada .xlsx:  pip install openpyxl

O CSV de saída é UTF-8 com BOM (utf-8-sig) para abrir corretamente no Excel
no Windows sem caracteres corrompidos.
"""

import argparse
import csv
import pathlib
import sys

from converter import convert, convert_runs


# ---------------------------------------------------------------------------
# Caminho CSV
# ---------------------------------------------------------------------------

def _process_csv(input_path, output_path, column_name, header_row):
    """Lê CSV, converte coluna-alvo com convert(), grava CSV de saída."""
    with open(input_path, newline="", encoding="utf-8-sig") as fh:
        reader = csv.reader(fh)
        all_rows = list(reader)

    if len(all_rows) < header_row:
        sys.stderr.write("Erro: arquivo tem menos linhas do que --header-row %d.\n" % header_row)
        return 1

    headers = all_rows[header_row - 1]
    target_lower = column_name.strip().lower()
    try:
        target_idx = next(
            i for i, h in enumerate(headers)
            if h.strip().lower() == target_lower
        )
    except StopIteration:
        sys.stderr.write(
            "Erro: coluna %r não encontrada. Colunas disponíveis: %s\n"
            % (column_name, [h for h in headers if h])
        )
        return 1

    converted = 0
    data_rows = all_rows[header_row:]
    for row in data_rows:
        if target_idx < len(row) and row[target_idx].strip():
            row[target_idx] = convert(row[target_idx])
            converted += 1

    sys.stderr.write("Convertidas %d descrições.\n" % converted)

    with open(output_path, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerows(all_rows[:header_row])
        writer.writerows(data_rows)

    sys.stderr.write("Saída gravada em %s\n" % output_path)
    return 0


def _list_csv_columns(input_path, header_row):
    with open(input_path, newline="", encoding="utf-8-sig") as fh:
        reader = csv.reader(fh)
        for _ in range(header_row - 1):
            next(reader, None)
        headers = next(reader, [])
    sys.stdout.write("Colunas encontradas:\n")
    for name in headers:
        if name.strip():
            sys.stdout.write("  %s\n" % name)
    return 0


# ---------------------------------------------------------------------------
# Caminho XLSX
# ---------------------------------------------------------------------------

def _process_xlsx(input_path, output_path, column_name, sheet_name, header_row):
    import openpyxl
    from xlsx_reader import read_column

    try:
        rows = read_column(
            input_path,
            column_name=column_name,
            sheet_name=sheet_name,
            header_row=header_row,
        )
    except ValueError as exc:
        sys.stderr.write("Erro: %s\n" % exc)
        return 1
    except FileNotFoundError:
        sys.stderr.write("Erro: arquivo não encontrado: %s\n" % input_path)
        return 1

    if not rows:
        sys.stderr.write("Nenhuma linha de dados encontrada na coluna %r.\n" % column_name)
        return 1

    wb = openpyxl.load_workbook(input_path, rich_text=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    headers = [
        ws.cell(row=header_row, column=c).value
        for c in range(1, ws.max_column + 1)
    ]

    target_col_lower = column_name.strip().lower()
    try:
        target_col_idx = next(
            i for i, h in enumerate(headers)
            if h is not None and str(h).strip().lower() == target_col_lower
        )
    except StopIteration:
        sys.stderr.write("Erro: coluna %r não encontrada nos cabeçalhos.\n" % column_name)
        wb.close()
        return 1

    converted_by_row = {r["row"]: convert_runs(r["paragraphs"]) for r in rows}
    sys.stderr.write("Convertidas %d descrições.\n" % len(converted_by_row))

    with open(output_path, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow(headers)
        for row_idx in range(header_row + 1, ws.max_row + 1):
            row_values = [
                ws.cell(row=row_idx, column=c).value
                for c in range(1, ws.max_column + 1)
            ]
            if all(v is None for v in row_values):
                continue
            if row_idx in converted_by_row:
                row_values[target_col_idx] = converted_by_row[row_idx]
            writer.writerow(row_values)

    wb.close()
    sys.stderr.write("Saída gravada em %s\n" % output_path)
    return 0


def _list_xlsx_columns(input_path, sheet_name, header_row):
    from xlsx_reader import read_all_columns
    cols = read_all_columns(input_path, sheet_name=sheet_name, header_row=header_row)
    sys.stdout.write("Colunas encontradas:\n")
    for name in cols:
        sys.stdout.write("  %s\n" % name)
    return 0


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Converte em lote uma coluna de descrições para HTML em linha única. "
                    "Aceita entrada .csv ou .xlsx.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--input", "-i", required=True,
                        help="Caminho para o arquivo de entrada (.csv ou .xlsx).")
    parser.add_argument("--output", "-o", default=None,
                        help="Caminho para o CSV de saída. "
                             "Padrão: <nome_entrada>_converted.csv.")
    parser.add_argument("--column", "-c", default=None,
                        help="Nome do cabeçalho da coluna de descrição a converter.")
    parser.add_argument("--sheet", "-s", default=None,
                        help="Apenas .xlsx: nome da planilha a ler (padrão: primeira planilha).")
    parser.add_argument("--header-row", type=int, default=1,
                        dest="header_row",
                        help="Número da linha (base 1) onde ficam os cabeçalhos (padrão: 1).")
    parser.add_argument("--list-columns", action="store_true",
                        dest="list_columns",
                        help="Imprime os nomes de colunas encontrados no arquivo e encerra.")
    args = parser.parse_args(argv)

    input_path = pathlib.Path(args.input)
    ext = input_path.suffix.lower()

    if ext not in (".csv", ".xlsx"):
        sys.stderr.write("Erro: tipo de arquivo não suportado %r. Use .csv ou .xlsx.\n" % ext)
        return 1

    if not input_path.exists():
        sys.stderr.write("Erro: arquivo não encontrado: %s\n" % input_path)
        return 1

    if args.list_columns:
        if ext == ".csv":
            return _list_csv_columns(input_path, args.header_row)
        return _list_xlsx_columns(input_path, args.sheet, args.header_row)

    if not args.column:
        parser.error("--column é obrigatório a menos que --list-columns seja usado.")

    output_path = pathlib.Path(args.output) if args.output else \
        input_path.with_name(input_path.stem + "_converted.csv")

    sys.stderr.write("Lendo %s ...\n" % input_path)

    if ext == ".csv":
        return _process_csv(input_path, output_path, args.column, args.header_row)
    return _process_xlsx(input_path, output_path, args.column, args.sheet, args.header_row)


if __name__ == "__main__":
    raise SystemExit(main())
