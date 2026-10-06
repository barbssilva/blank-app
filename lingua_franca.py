import sys
from copy import copy
from fractions import Fraction
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ---------------- CONFIGURAÇÃO ----------------
TEXTO_INICIO_TABELA = "CODE"   # texto na coluna A que marca a linha de títulos da tabela
PRIMEIRA_COLUNA_INCHES = 3     # 3 = coluna C (TOL). Colunas A e B (código e descrição) não mudam
TITULO_COLUNA_CM = "cm"        # texto que aparece no título de cada nova coluna
FORMATO_CM = "0.00"             # 1 casa decimal (use "0.00" para 2)
CASAS_DECIMAIS = 2             # precisão guardada na célula
AJUSTAR_IMPRESSAO = True       # imprime a folha com 1 página de largura
LARGURA_COLUNA_CM = None       # largura das colunas novas; None = igual à coluna em polegadas
# ----------------------------------------------


def polegadas_para_numero(valor):
    """Converte '24 1/2', '-1/2', '3/8', 25, '6\xa03/4' em número (float)."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).replace("\xa0", " ").replace('"', "").strip()
    negativo = texto.startswith("-")
    texto = texto.lstrip("-").strip()
    try:
        total = sum(Fraction(parte) for parte in texto.split())
    except (ValueError, ZeroDivisionError):
        return None  # texto que não é medida -> fica em branco
    return float(-total if negativo else total)


def encontrar_tabela(ws):
    """Devolve (linha_titulos, ultima_linha, ultima_coluna) da tabela de medidas."""
    linha_titulos = None
    for linha in range(1, ws.max_row + 1):
        v = ws.cell(linha, 1).value
        if isinstance(v, str) and v.strip().upper() == TEXTO_INICIO_TABELA:
            linha_titulos = linha
            break
    if linha_titulos is None:
        raise ValueError(f"Não encontrei '{TEXTO_INICIO_TABELA}' na coluna A.")

    ultima_linha = linha_titulos
    for linha in range(linha_titulos + 1, ws.max_row + 1):
        if any(ws.cell(linha, c).value not in (None, "") for c in (1, 2)):
            ultima_linha = linha

    ultima_coluna = 1
    for linha in range(linha_titulos, ultima_linha + 1):
        for col in range(1, ws.max_column + 1):
            if ws.cell(linha, col).value not in (None, ""):
                ultima_coluna = max(ultima_coluna, col)
    return linha_titulos, ultima_linha, ultima_coluna


def processar(caminho_entrada, caminho_saida):
    wb = load_workbook(caminho_entrada,rich_text=True)
    ws = wb.active
    linha_titulos, ultima_linha, ultima_coluna = encontrar_tabela(ws)

    # Proteção: a tabela não pode ter células unidas (as do cabeçalho não são tocadas)
    for m in ws.merged_cells.ranges:
        if m.max_row >= linha_titulos and m.min_row <= ultima_linha:
            raise ValueError(f"A tabela tem células unidas ({m}); separe-as primeiro.")

    colunas_inches = list(range(PRIMEIRA_COLUNA_INCHES, ultima_coluna + 1))
    larguras = {
        col: ws.column_dimensions[get_column_letter(col)].width
        for col in colunas_inches
    }

    # Guardar o conteúdo e o formato originais da tabela
    original = {}
    for linha in range(linha_titulos, ultima_linha + 1):
        for col in colunas_inches:
            c = ws.cell(linha, col)
            original[(linha, col)] = (c.value, c)

    # Copiar o formato antes de reescrever (porque as células vão ser sobrepostas)
    estilos = {}
    for chave, (valor, cel) in original.items():
        estilos[chave] = {
            "font": copy(cel.font), "fill": copy(cel.fill), "border": copy(cel.border),
            "alignment": copy(cel.alignment), "protection": copy(cel.protection),
            "number_format": cel.number_format,
        }

    def aplicar(cel, est, formato=None):
        cel.font, cel.fill, cel.border = est["font"], est["fill"], est["border"]
        cel.alignment, cel.protection = est["alignment"], est["protection"]
        cel.number_format = formato or est["number_format"]

    # Reescrever da direita para a esquerda: coluna k -> nova posição + coluna cm ao lado
    for i, col in reversed(list(enumerate(colunas_inches))):
        nova_col = PRIMEIRA_COLUNA_INCHES + 2 * i
        col_cm = nova_col + 1
        for linha in range(linha_titulos, ultima_linha + 1):
            valor, _ = original[(linha, col)]
            est = estilos[(linha, col)]

            destino = ws.cell(linha, nova_col)
            destino.value = valor
            aplicar(destino, est)

            cel_cm = ws.cell(linha, col_cm)
            aplicar(cel_cm, est, FORMATO_CM)
            if linha == linha_titulos:
                cel_cm.value = TITULO_COLUNA_CM
                cel_cm.number_format = "General"
                # usa a letra do título "TOL" (as células vazias do título não têm letra definida)
                cel_cm.font = copy(estilos[(linha_titulos, PRIMEIRA_COLUNA_INCHES)]["font"])
                cel_cm.alignment = copy(estilos[(linha_titulos, PRIMEIRA_COLUNA_INCHES)]["alignment"])
            else:
                num = polegadas_para_numero(valor)
                cel_cm.value = None if num is None else round(num * 2.54, CASAS_DECIMAIS)

        # Larguras: só define as colunas novas (à direita da tabela original),
        # para que as colunas do cabeçalho mantenham exatamente a largura que tinham
        w = larguras.get(col)
        for c in (nova_col, col_cm):
            letra = get_column_letter(c)
            if c > ultima_coluna and w:
                ws.column_dimensions[letra].width = (LARGURA_COLUNA_CM or w) if c == col_cm else w

    if AJUSTAR_IMPRESSAO:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_setup.orientation = "landscape"

    wb.save(caminho_saida)
