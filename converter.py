#!/usr/bin/env python3
"""
converter.py -- Conversor de texto rico para HTML em linha única.

Converte texto colado (de Word ou Excel) em uma string HTML de linha única,
adequada para sistemas que não aceitam quebras de linha dentro de campos CSV.

Dois caminhos de conversão disponíveis:

  convert(raw)              -- entrada de texto simples (via stdin).
                               Produz <ul>/<ol>/<li>, <br>, <a href>.
                               Não produz <strong>/<em>/<u> pois texto
                               simples não carrega metadados de formatação.

  convert_runs(paragraphs)  -- entrada estruturada de runs (via xlsx_reader.py).
                               Cada parágrafo é uma lista de Run namedtuples
                               com texto + flags bold/italic/underline.
                               Produz todos os tags acima MAIS <strong>, <em>, <u>.

O modo interativo (main) usa o caminho de texto simples. batch.py usa convert_runs.
Apenas biblioteca padrão para convert(); openpyxl é dependência de batch.py / xlsx_reader.
"""

import collections
import html
import re
import sys
import unicodedata

# ---------------------------------------------------------------------------
# Tipo de dado compartilhado: Run
# ---------------------------------------------------------------------------
# Um Run é um trecho contínuo de texto com formatação uniforme. xlsx_reader.py
# produz listas destes; convert_runs() os consome.
# Todos os flags padrão False — quem chama só define o que conhece.

Run = collections.namedtuple("Run", ["text", "bold", "italic", "underline"],
                             defaults=[False, False, False])

# ---------------------------------------------------------------------------
# Classificação de linha  (executada na linha BRUTA, antes da normalização)
# ---------------------------------------------------------------------------
# A ordem importa: decidimos se uma linha é item de lista ANTES de normalizar
# traços, para não confundir uma linha de prosa iniciada com travessão com um
# item de lista.

_BULLET_GLYPHS = "•·●▪■‣⁃∙◦․‧"
_ASCII_BULLET = "-*"

_BULLET_RE = re.compile(
    r"^[ \t]*[" + re.escape(_BULLET_GLYPHS + _ASCII_BULLET) + r"][ \t]+(?P<text>.*\S.*)$"
)

_NUMBERED_RE = re.compile(
    r"^[ \t]*\d+[ \t]*[.)][ \t]+(?P<text>.*\S.*)$"
)


def classify_line(line):
    """Retorna (tipo, conteúdo) para uma linha bruta.

    tipo é 'blank', 'ul', 'ol' ou 'text'. Para itens de lista, conteúdo tem o
    marcador removido. Um marcador sem texto é tratado como linha em branco.
    """
    stripped = line.strip()
    if not stripped:
        return ("blank", "")
    m = _NUMBERED_RE.match(stripped)
    if m:
        return ("ol", m.group("text").strip())
    m = _BULLET_RE.match(stripped)
    if m:
        return ("ul", m.group("text").strip())
    return ("text", stripped)


# ---------------------------------------------------------------------------
# Normalização de caracteres (tipografia "inteligente" do Word/Excel -> ASCII)
# ---------------------------------------------------------------------------

_SMART_PUNCT = {
    "‘": "'", "’": "'", "‚": "'", "‛": "'",
    "′": "'",
    "“": '"', "”": '"', "„": '"', "‟": '"',
    "″": '"',
    "–": "-", "—": "-", "―": "-", "−": "-",
    "…": "...",
    " ": " ", " ": " ", " ": " ", "\t": " ",
    "​": "", "﻿": "",
}
_SMART_TABLE = {ord(k): v for k, v in _SMART_PUNCT.items()}


def normalize_content(text):
    """Normaliza NFC e remove tipografia inteligente do Word/Excel."""
    text = unicodedata.normalize("NFC", text)
    return text.translate(_SMART_TABLE)


# ---------------------------------------------------------------------------
# Escape HTML + auto-link de URLs
# ---------------------------------------------------------------------------

_URL_RE = re.compile(r"(https?://[^\s<>\"]+)")


def escape(text):
    """Escapa caracteres especiais HTML em conteúdo de texto de elemento."""
    return html.escape(text, quote=False)


def escape_and_link(text):
    """Escapa conteúdo e transforma URLs http(s) nuas em <a href="...">."""
    out = []
    last = 0
    for m in _URL_RE.finditer(text):
        out.append(escape(text[last:m.start()]))
        url = m.group(1)
        href = html.escape(url, quote=True)
        out.append('<a href="%s">%s</a>' % (href, escape(url)))
        last = m.end()
    out.append(escape(text[last:]))
    return "".join(out)


# ---------------------------------------------------------------------------
# Montagem de blocos
# ---------------------------------------------------------------------------

def convert(raw):
    """Converte texto rico colado em uma string HTML de linha única."""
    text = raw.replace("\r\n", "\n").replace("\r", "\n")
    classified = [classify_line(ln) for ln in text.split("\n")]

    parts = []
    i = 0
    n = len(classified)
    emitted = False

    while i < n:
        kind = classified[i][0]

        if kind == "blank":
            if emitted and _has_more_content(classified, i):
                parts.append("<br><br>")
            while i < n and classified[i][0] == "blank":
                i += 1
            continue

        if kind in ("ul", "ol"):
            tag = kind
            items = []
            while i < n and classified[i][0] == tag:
                content = normalize_content(classified[i][1])
                items.append("<li>%s</li>" % escape_and_link(content))
                i += 1
            parts.append("<%s>%s</%s>" % (tag, "".join(items), tag))
            emitted = True
            continue

        text_lines = []
        while i < n and classified[i][0] == "text":
            content = normalize_content(classified[i][1])
            text_lines.append(escape_and_link(content))
            i += 1
        parts.append("<br>".join(text_lines))
        emitted = True

    html_out = "".join(parts)
    html_out = html_out.replace("\r", "").replace("\n", "")
    return html_out.strip()


def _has_more_content(classified, idx):
    """Verdadeiro se existe alguma linha não-vazia após a posição idx."""
    return any(classified[j][0] != "blank" for j in range(idx + 1, len(classified)))


# ---------------------------------------------------------------------------
# Caminho de conversão por runs (entrada Excel — reconhece negrito/itálico/sublinhado)
# ---------------------------------------------------------------------------

def _wrap_run(text, bold, italic, underline):
    """Envolve texto escapado em tags de formatação inline conforme necessário."""
    out = escape_and_link(text)
    if underline:
        out = "<u>%s</u>" % out
    if italic:
        out = "<em>%s</em>" % out
    if bold:
        out = "<strong>%s</strong>" % out
    return out


def _runs_to_plain(runs):
    """Extrai texto simples de uma lista de runs (para classificação de linha)."""
    return "".join(r.text for r in runs)


def _render_runs(runs):
    """Renderiza uma lista de runs como HTML."""
    if not runs:
        return ""
    out = []
    for run in runs:
        text = normalize_content(run.text)
        if not text:
            continue
        out.append(_wrap_run(text, run.bold, run.italic, run.underline))
    return "".join(out)


def convert_runs(paragraphs):
    """Converte lista de parágrafos (cada um com lista de Run) para HTML de linha única.

    Quebras de linha Alt+Enter dentro de runs são tratadas como <br> dentro do parágrafo.
    Parágrafos em branco entre conteúdo produzem <br><br>.

    Retorna string HTML sem \\n ou \\r.
    """
    lines = []
    for para in paragraphs:
        expanded = [[]]
        for run in para:
            segments = run.text.split("\n")
            expanded[-1].append(Run(segments[0], run.bold, run.italic, run.underline))
            for seg in segments[1:]:
                expanded.append([Run(seg, run.bold, run.italic, run.underline)])
        lines.extend(expanded)
        lines.append(None)

    parts = []
    emitted = False
    i = 0
    n = len(lines)

    while i < n:
        line_runs = lines[i]

        if line_runs is None:
            has_more = any(
                lines[j] is not None and _runs_to_plain(lines[j]).strip()
                for j in range(i + 1, n)
            )
            if emitted and has_more:
                parts.append("<br><br>")
            i += 1
            continue

        plain = _runs_to_plain(line_runs)
        kind, content_plain = classify_line(plain)

        if kind == "blank":
            i += 1
            continue

        if kind in ("ul", "ol"):
            tag = kind
            items = []
            while i < n and lines[i] is not None:
                pk, _ = classify_line(_runs_to_plain(lines[i]))
                if pk != tag:
                    break
                adjusted = _strip_list_marker(lines[i], tag)
                items.append("<li>%s</li>" % _render_runs(adjusted))
                i += 1
            parts.append("<%s>%s</%s>" % (tag, "".join(items), tag))
            emitted = True
            continue

        text_parts = []
        while i < n and lines[i] is not None:
            pk, _ = classify_line(_runs_to_plain(lines[i]))
            if pk != "text":
                break
            text_parts.append(_render_runs(lines[i]))
            i += 1
        parts.append("<br>".join(text_parts))
        emitted = True

    html_out = "".join(parts)
    html_out = html_out.replace("\r", "").replace("\n", "")
    return html_out.strip()


def _strip_list_marker(runs, kind):
    """Retorna nova lista de runs sem o marcador de lista/número à esquerda."""
    if not runs:
        return runs

    plain = _runs_to_plain(runs)
    if kind == "ul":
        m = _BULLET_RE.match(plain.strip())
    else:
        m = _NUMBERED_RE.match(plain.strip())

    if not m:
        return runs

    content_start = plain.index(m.group("text"))

    result = []
    consumed = 0
    for run in runs:
        run_end = consumed + len(run.text)
        if run_end <= content_start:
            consumed = run_end
            continue
        if consumed < content_start:
            trim = content_start - consumed
            result.append(Run(run.text[trim:], run.bold, run.italic, run.underline))
        else:
            result.append(run)
        consumed = run_end
    return result


# ---------------------------------------------------------------------------
# Entrada interativa (robusto no Windows)
# ---------------------------------------------------------------------------

_PROMPT = (
    "Cole seu texto abaixo (copiado do Word ou Excel).\n"
    "Quando terminar, pressione Ctrl-Z e depois Enter (Windows), ou Ctrl-D (Mac/Linux).\n"
    "-------------------------------------------------------------------\n"
)


def read_input(stream=None):
    """Lê o buffer colado do stdin até EOF."""
    stream = stream or sys.stdin
    reconfigure = getattr(stream, "reconfigure", None)
    if reconfigure is not None:
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except (ValueError, OSError):
            pass
    return stream.read()


def main(argv=None):
    if sys.stdin.isatty():
        sys.stderr.write(_PROMPT)
        sys.stderr.flush()

    try:
        raw = read_input()
    except KeyboardInterrupt:
        sys.stderr.write("\nCancelado.\n")
        return 130

    if not raw.strip():
        sys.stderr.write("Nenhuma entrada recebida. Nada a converter.\n")
        return 1

    if "\t" in raw:
        sys.stderr.write(
            "Atenção: tabulações detectadas (parece uma linha do Excel). As colunas "
            "serão unidas como texto simples, não mantidas como campos separados.\n"
        )

    result = convert(raw)

    if sys.stdin.isatty():
        sys.stderr.write("\n-------------------- HTML em linha única --------------------\n")
        sys.stderr.flush()

    print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
