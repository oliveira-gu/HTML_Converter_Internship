#!/usr/bin/env python3
"""Testes demonstrativos reproduziveis do HTML Converter.

Execute na pasta que contem converter.py, batch.py e xlsx_reader.py:
    python -m unittest testar_conversor.py -v

Os dados sao SINTETICOS, gerados em uma pasta temporaria durante a execucao.
Este arquivo NAO representa testes historicos do estagio, e nao testa
carregamentos ou renderizacao dentro do SAP Fieldglass.

Dependencia para os testes XLSX: openpyxl >= 3.1.
"""

import csv
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from converter import Run, convert, convert_runs

BASE = Path(__file__).resolve().parent

DADOS = [
    ["EX-001", "Primeira linha\nSegunda linha", "Quebras de linha"],
    ["EX-002", "- Primeiro requisito\n- Segundo requisito", "Lista não ordenada"],
    ["EX-003", "1. Preparar dados\n2. Validar resultado", "Lista ordenada"],
    ["EX-004", "Texto com & e <sinal> e aspas “especiais”.", "Caracteres especiais"],
    ["EX-005", "Documentação: https://example.org/guia", "URL"],
    ["EX-006", 'Uma linha, com vírgula e "aspas".', "Delimitadores CSV"],
]


def executar_batch(*args):
    return subprocess.run(
        [sys.executable, str(BASE / "batch.py"), *[str(a) for a in args]],
        cwd=str(BASE), text=True, capture_output=True, encoding="utf-8", errors="replace"
    )


def preparar_csv(path):
    with path.open("w", encoding="utf-8-sig", newline="") as arquivo:
        escritor = csv.writer(arquivo)
        escritor.writerow(["ID", "Descrição", "Categoria"])
        escritor.writerows(DADOS)


def preparar_xlsx(path):
    from openpyxl import Workbook
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    from openpyxl.styles import Font

    livro = Workbook()
    planilha = livro.active
    planilha.title = "Descricoes"
    planilha.append(["ID", "Descrição", "Categoria"])
    for linha in DADOS:
        planilha.append(linha)

    # Linhas extras para testar estilos de celula e formatos mistos na mesma celula.
    planilha.append(["EX-007", "Negrito da célula", "Negrito"])
    planilha["B8"].font = Font(bold=True)
    planilha.append(["EX-008", "Itálico da célula", "Itálico"])
    planilha["B9"].font = Font(italic=True)
    planilha.append(["EX-009", "Sublinhado da célula", "Sublinhado"])
    planilha["B10"].font = Font(underline="single")
    planilha.append(["EX-010", "", "Estilos mistos"])
    planilha["B11"] = CellRichText([
        TextBlock(InlineFont(b=True), "negrito"),
        " e ",
        TextBlock(InlineFont(i=True), "itálico"),
    ])
    livro.save(path)
    livro.close()


class TestesConversao(unittest.TestCase):
    def test_CT01_quebra_de_linha(self):
        self.assertEqual(convert(DADOS[0][1]), "Primeira linha<br>Segunda linha")

    def test_CT02_lista_nao_ordenada(self):
        self.assertEqual(convert(DADOS[1][1]),
                         "<ul><li>Primeiro requisito</li><li>Segundo requisito</li></ul>")

    def test_CT03_lista_ordenada(self):
        self.assertEqual(convert(DADOS[2][1]),
                         "<ol><li>Preparar dados</li><li>Validar resultado</li></ol>")

    def test_CT04_caracteres_especiais(self):
        self.assertEqual(convert(DADOS[3][1]),
                         'Texto com &amp; e &lt;sinal&gt; e aspas "especiais".')

    def test_CT05_link_http(self):
        self.assertEqual(convert(DADOS[4][1]),
                         'Documentação: <a href="https://example.org/guia">https://example.org/guia</a>')

    def test_CT11_caracteres_delimitadores_csv(self):
        self.assertEqual(convert(DADOS[5][1]), 'Uma linha, com vírgula e "aspas".')

    def test_CT12_linha_em_branco(self):
        self.assertEqual(convert("Antes\n\nDepois"), "Antes<br><br>Depois")

    def test_CT08_ausencia_de_quebras_fisicas(self):
        for linha in DADOS:
            saida = convert(linha[1])
            with self.subTest(id=linha[0]):
                self.assertNotIn("\n", saida)
                self.assertNotIn("\r", saida)

    def test_CT13_runs_com_negrito_italico_sublinhado(self):
        paragrafos = [[
            Run("negrito", True, False, False), Run(" e "),
            Run("itálico", False, True, False), Run(" e "),
            Run("sublinhado", False, False, True),
        ]]
        self.assertEqual(convert_runs(paragrafos),
                         "<strong>negrito</strong> e <em>itálico</em> e <u>sublinhado</u>")

    def test_CT14_quebras_em_runs_formatados(self):
        saida = convert_runs([[Run("Linha 1\nLinha 2", bold=True)]])
        self.assertEqual(saida, "<strong>Linha 1</strong><br><strong>Linha 2</strong>")
        self.assertNotIn("\n", saida)


class TestesLote(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory(prefix="html_converter_teste_")
        cls.pasta = Path(cls.temp.name)
        cls.csv = cls.pasta / "entrada_exemplo.csv"
        cls.xlsx = cls.pasta / "entrada_exemplo.xlsx"
        preparar_csv(cls.csv)
        preparar_xlsx(cls.xlsx)

    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()

    def test_CT06_csv_lote_preserva_registros_e_colunas(self):
        saida = self.pasta / "resultado_csv.csv"
        proc = executar_batch("--input", self.csv, "--output", saida,
                              "--column", "Descrição")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        with saida.open(encoding="utf-8-sig", newline="") as arq:
            linhas = list(csv.DictReader(arq))
        self.assertEqual(len(linhas), len(DADOS))
        for i, dados in enumerate(DADOS):
            with self.subTest(id=dados[0]):
                self.assertEqual(linhas[i]["ID"], dados[0])
                self.assertEqual(linhas[i]["Categoria"], dados[2])
                self.assertEqual(linhas[i]["Descrição"], convert(dados[1]))
                self.assertNotIn("\n", linhas[i]["Descrição"])

    def test_CT15_csv_saida_tem_bom(self):
        saida = self.pasta / "resultado_bom.csv"
        proc = executar_batch("--input", self.csv, "--output", saida,
                              "--column", "Descrição")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertTrue(saida.read_bytes().startswith(b"\xef\xbb\xbf"))

    def test_CT16_xlsx_lote_seis_exemplos(self):
        saida = self.pasta / "resultado_xlsx.csv"
        proc = executar_batch("--input", self.xlsx, "--output", saida,
                              "--column", "Descrição", "--sheet", "Descricoes")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        with saida.open(encoding="utf-8-sig", newline="") as arq:
            linhas = list(csv.DictReader(arq))
        self.assertEqual(len(linhas), 10)
        for i, dados in enumerate(DADOS):
            with self.subTest(id=dados[0]):
                self.assertEqual(linhas[i]["Descrição"], convert(dados[1]))
                self.assertEqual(linhas[i]["ID"], dados[0])

    def test_CT07_xlsx_estilo_celula_negrito_italico_sublinhado(self):
        saida = self.pasta / "resultado_estilos.csv"
        proc = executar_batch("--input", self.xlsx, "--output", saida,
                              "--column", "Descrição")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        with saida.open(encoding="utf-8-sig", newline="") as arq:
            linhas = {l["ID"]: l["Descrição"] for l in csv.DictReader(arq)}
        self.assertEqual(linhas["EX-007"], "<strong>Negrito da célula</strong>")
        self.assertEqual(linhas["EX-008"], "<em>Itálico da célula</em>")
        self.assertEqual(linhas["EX-009"], "<u>Sublinhado da célula</u>")

    def test_CT17_xlsx_rich_text_em_uma_celula(self):
        saida = self.pasta / "resultado_richtext.csv"
        proc = executar_batch("--input", self.xlsx, "--output", saida,
                              "--column", "Descrição")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        with saida.open(encoding="utf-8-sig", newline="") as arq:
            linhas = {l["ID"]: l["Descrição"] for l in csv.DictReader(arq)}
        self.assertEqual(linhas["EX-010"], "<strong>negrito</strong> e <em>itálico</em>")

    def test_CT18_listagem_de_colunas_csv(self):
        proc = executar_batch("--input", self.csv, "--list-columns")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn("Descrição", proc.stdout)
        self.assertIn("Categoria", proc.stdout)

    def test_CT09_coluna_inexistente_rejeitada(self):
        saida = self.pasta / "nao_deve_existir.csv"
        proc = executar_batch("--input", self.csv, "--output", saida,
                              "--column", "Coluna Ausente")
        self.assertNotEqual(proc.returncode, 0)
        self.assertIn("não encontrada", proc.stderr)
        self.assertFalse(saida.exists())

    def test_CT10_extensao_invalida_rejeitada(self):
        txt = self.pasta / "entrada.txt"
        txt.write_text("qualquer dado", encoding="utf-8")
        proc = executar_batch("--input", txt, "--column", "Descrição")
        self.assertNotEqual(proc.returncode, 0)
        self.assertIn("não suportado", proc.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
