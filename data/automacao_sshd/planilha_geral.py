from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string

from .config import COLUNAS_PLANILHA_GERAL, VALORES_PADRAO
from .exceptions import ErroLeituraPlanilha
from .normalizacao import limpar_cabecalho, valor_excel, valor_vazio


CABECALHOS_MINIMOS = {"nome_colaborador", "cpf", "cargo"}


def _encontrar_linha_cabecalho(worksheet) -> int:
    for numero_linha in range(1, min(worksheet.max_row, 10) + 1):
        cabecalhos = {
            limpar_cabecalho(worksheet.cell(numero_linha, coluna).value)
            for coluna in range(1, worksheet.max_column + 1)
        }
        if CABECALHOS_MINIMOS.issubset(cabecalhos):
            return numero_linha

    raise ErroLeituraPlanilha(
        "Nao foi possivel localizar o cabecalho da Planilha Geral. "
        "Verifique se existem as colunas Nome Colaborador, CPF e Cargo."
    )


def _ler_por_letra(row: int, worksheet, letra_coluna: str) -> Any:
    indice_coluna = column_index_from_string(letra_coluna)
    return valor_excel(worksheet.cell(row, indice_coluna).value)


def ler_planilha_geral(caminho_planilha: str | Path) -> list[dict[str, Any]]:
    caminho = Path(caminho_planilha)

    if not caminho.exists():
        raise ErroLeituraPlanilha(f"Planilha nao encontrada: {caminho}")
    if caminho.suffix.lower() != ".xlsx":
        raise ErroLeituraPlanilha("Nesta etapa inicial, selecione uma planilha no formato .xlsx.")

    try:
        workbook = load_workbook(caminho, data_only=True, read_only=False)
    except Exception as erro:
        raise ErroLeituraPlanilha(f"Nao foi possivel abrir a planilha de origem: {erro}") from erro

    worksheet = workbook.active
    linha_cabecalho = _encontrar_linha_cabecalho(worksheet)
    primeira_linha_dados = linha_cabecalho + 1

    colaboradores: list[dict[str, Any]] = []

    for numero_linha in range(primeira_linha_dados, worksheet.max_row + 1):
        registro = {
            campo: _ler_por_letra(numero_linha, worksheet, letra_coluna)
            for campo, letra_coluna in COLUNAS_PLANILHA_GERAL.items()
        }

        if all(valor_vazio(valor) for valor in registro.values()):
            continue

        # A Prefeitura exige estes campos, mas eles nao existem explicitamente
        # na Planilha Geral enviada; por isso entram como regra inicial.
        for campo, valor_padrao in VALORES_PADRAO.items():
            if valor_vazio(registro.get(campo)):
                registro[campo] = valor_padrao

        registro["linha_origem"] = numero_linha
        colaboradores.append(registro)

    if not colaboradores:
        raise ErroLeituraPlanilha("Nenhum colaborador foi encontrado na planilha selecionada.")

    return colaboradores
