# Conversor de Texto Rico para HTML

Ferramenta de linha de comando que converte texto rico (copiado do Word ou Excel) em uma **string HTML de linha única**, adequada para sistemas de importação via CSV que não aceitam quebras de linha dentro de campos.

## Motivação

Algumas plataformas de gestão de força de trabalho utilizam conectores de upload via CSV para importar descrições de vagas ou modelos de texto. Um problema comum nesses conectores é que qualquer quebra de linha dentro de um campo CSV é interpretada como início de um novo registro, corrompendo a importação.

A solução é representar a formatação do texto em HTML — listas, parágrafos, negrito, itálico — em uma única linha contínua. Este projeto automatiza essa conversão.

## Estrutura do Projeto

| Arquivo | Função |
|---|---|
| `converter.py` | Lógica principal de conversão. `convert(raw)` para entrada de texto simples; `convert_runs(paragraphs)` para dados de runs do Excel. Também é o ponto de entrada do modo interativo. |
| `xlsx_reader.py` | Lê um arquivo `.xlsx` com openpyxl e extrai os runs de texto rico de cada célula como namedtuples `Run`. |
| `batch.py` | CLI de lote: detecta `.csv` ou `.xlsx` pela extensão e encaminha para a conversão adequada. Grava o CSV de saída. |

## Requisitos

- Python 3.8 ou superior
- Para modo de lote com arquivos Excel: `openpyxl`

```bash
pip install openpyxl
```

No Windows, se o comando `python` não funcionar, use `py -3` ou o caminho completo do interpretador.

## Como Usar

### Modo interativo — colar e converter (texto simples)

```bash
py -3 converter.py
```

Cole o texto quando solicitado. No Windows, finalize com **Ctrl-Z e depois Enter**. No Mac/Linux, use **Ctrl-D**. O HTML de linha única é impresso no terminal, pronto para copiar.

**Via pipe (para testes rápidos):**

```bash
printf -- '- Primeiro item\n- Segundo item' | py -3 converter.py
```

### Modo de lote — converter arquivo CSV (texto simples)

```bash
# Descobrir nomes de colunas:
py -3 batch.py --input dados.csv --list-columns

# Converter a coluna "Descrição":
py -3 batch.py --input dados.csv --output resultado.csv --column "Descrição"
```

### Modo de lote — converter arquivo Excel (preserva negrito, itálico, sublinhado)

```bash
# Descobrir nomes de colunas:
py -3 batch.py --input dados.xlsx --list-columns

# Converter e gravar resultado.csv:
py -3 batch.py --input dados.xlsx --output resultado.csv --column "Descrição"

# Especificar planilha (se houver múltiplas abas):
py -3 batch.py --input dados.xlsx --output resultado.csv --column "Descrição" --sheet "Planilha1"

# Especificar linha de cabeçalho (se os cabeçalhos não estiverem na linha 1):
py -3 batch.py --input dados.xlsx --output resultado.csv --column "Descrição" --header-row 3
```

O CSV de saída é gravado em UTF-8 com BOM para abrir corretamente no Excel no Windows.

## O que é convertido

| Entrada | Saída HTML |
|---|---|
| `- item` / `* item` / `• item` | `<ul><li>item</li></ul>` |
| `1. item` / `1) item` | `<ol><li>item</li></ol>` |
| Quebra de linha simples dentro de um bloco | `<br>` |
| Linha em branco entre blocos | `<br><br>` |
| URL `http(s)://` nua | `<a href="...">...</a>` |
| `&`, `<`, `>` no texto | Escapados (`&amp;`, `&lt;`, `&gt;`) |
| **Run em negrito (modo Excel)** | `<strong>...</strong>` |
| **Run em itálico (modo Excel)** | `<em>...</em>` |
| **Run sublinhado (modo Excel)** | `<u>...</u>` |

Aspas inteligentes, travessões e espaços especiais do Word são normalizados para ASCII equivalente. A saída tem garantia de não conter `\n` ou `\r`.

## O que não é suportado (por design)

- **Sem `<strong>` / `<em>` / `<u>` a partir de texto simples ou CSV.** Esses formatos não carregam metadados de formatação. Use o modo de lote com `.xlsx` para obter essas tags.
- **Sem entrada `.docx`.** Apenas `.xlsx` é suportado para processamento em lote com texto rico.
- **Sem conversão direta na planilha.** A ferramenta opera externamente sobre o arquivo; não modifica o arquivo original.
