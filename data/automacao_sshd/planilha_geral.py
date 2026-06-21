from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string

from .config import COLUNAS_PLANILHA_GERAL, VALORES_AUTOMATICOS
from .exceptions import ErroLeituraPlanilha
from .normalizacao import limpar_cabecalho, normalizar_texto, valor_excel, valor_vazio


LIMITE_BUSCA_CABECALHO = 50
CABECALHOS_OBRIGATORIOS = {
    "Nome Colaborador": {
        "colaborador",
        "funcionario",
        "nome",
        "nome_colaborador",
        "nome_completo",
        "nome_do_colaborador",
        "nome_do_funcionario",
        "nome_funcionario",
    },
    "CPF": {
        "cpf",
        "cpf_colaborador",
        "cpf_do_colaborador",
        "cpf_funcionario",
    },
    "Cargo": {
        "cargo",
        "cargo_funcao",
        "cargo_ou_funcao",
        "funcao",
    },
}


def _cabecalhos_da_linha(worksheet, numero_linha: int) -> set[str]:
    return {
        limpar_cabecalho(worksheet.cell(numero_linha, coluna).value)
        for coluna in range(1, worksheet.max_column + 1)
        if not valor_vazio(worksheet.cell(numero_linha, coluna).value)
    }


def _linha_tem_cabecalhos_obrigatorios(cabecalhos: set[str]) -> bool:
    return all(
        bool(cabecalhos.intersection(alias))
        for alias in CABECALHOS_OBRIGATORIOS.values()
    )


def _descrever_linhas_lidas(workbook) -> str:
    amostras: list[str] = []

    for worksheet in workbook.worksheets:
        for numero_linha in range(1, min(worksheet.max_row, LIMITE_BUSCA_CABECALHO) + 1):
            valores = [
                normalizar_texto(worksheet.cell(numero_linha, coluna).value)
                for coluna in range(1, min(worksheet.max_column, 12) + 1)
                if not valor_vazio(worksheet.cell(numero_linha, coluna).value)
            ]
            if valores:
                amostras.append(
                    f"Aba '{worksheet.title}', linha {numero_linha}: " + " | ".join(valores[:8])
                )
            if len(amostras) >= 8:
                return "\n".join(amostras)

    return "Nenhum texto foi encontrado nas primeiras linhas das abas."


def _encontrar_planilha_e_linha_cabecalho(workbook) -> tuple[Any, int]:
    for worksheet in workbook.worksheets:
        for numero_linha in range(1, min(worksheet.max_row, LIMITE_BUSCA_CABECALHO) + 1):
            cabecalhos = _cabecalhos_da_linha(worksheet, numero_linha)
            if _linha_tem_cabecalhos_obrigatorios(cabecalhos):
                return worksheet, numero_linha

    amostras = _descrever_linhas_lidas(workbook)
    raise ErroLeituraPlanilha(
        "Nao foi possivel localizar o cabecalho da Planilha Geral. "
        "Procurei por colunas equivalentes a Nome Colaborador, CPF e Cargo "
        f"nas primeiras {LIMITE_BUSCA_CABECALHO} linhas de todas as abas.\n\n"
        "Amostra do que foi encontrado:\n"
        f"{amostras}"
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

    worksheet, linha_cabecalho = _encontrar_planilha_e_linha_cabecalho(workbook)
    primeira_linha_dados = linha_cabecalho + 1

    colaboradores: list[dict[str, Any]] = []

    for numero_linha in range(primeira_linha_dados, worksheet.max_row + 1):
        registro = {
            campo: _ler_por_letra(numero_linha, worksheet, letra_coluna)
            for campo, letra_coluna in COLUNAS_PLANILHA_GERAL.items()
        }

        if all(valor_vazio(valor) for valor in registro.values()):
            continue

        # Campos definidos por regra de negocio. Genero e Orientacao Sexual
        # nao devem vir da planilha de origem.
        registro.update(VALORES_AUTOMATICOS)

        registro["linha_origem"] = numero_linha
        colaboradores.append(registro)

    if not colaboradores:
        raise ErroLeituraPlanilha("Nenhum colaborador foi encontrado na planilha selecionada.")

    return colaboradores
