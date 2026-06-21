from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .normalizacao import proximo_caminho_disponivel
from .validacao import PendenciaValidacao


VERDE = "22C55E"
VERMELHO = "EF4444"
FUNDO_CABECALHO = "0F1F17"
TEXTO_CLARO = "F4F7F5"


def _formatar_cabecalho(worksheet) -> None:
    for cell in worksheet[1]:
        cell.font = Font(bold=True, color=TEXTO_CLARO)
        cell.fill = PatternFill("solid", fgColor=FUNDO_CABECALHO)
        cell.alignment = Alignment(horizontal="center")


def _ajustar_larguras(worksheet) -> None:
    for coluna in worksheet.columns:
        largura = 0
        letra = get_column_letter(coluna[0].column)
        for cell in coluna:
            valor = "" if cell.value is None else str(cell.value)
            largura = max(largura, len(valor))
        worksheet.column_dimensions[letra].width = min(max(largura + 2, 12), 60)


def gerar_relatorio_validacao(
    *,
    caminho_relatorio: str | Path,
    colaboradores: list[dict[str, Any]],
    pendencias: list[PendenciaValidacao],
    caminho_fonte: str | Path,
) -> Path:
    saida = proximo_caminho_disponivel(Path(caminho_relatorio))
    saida.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    resumo = workbook.active
    resumo.title = "Resumo"

    status = "APROVADO" if not pendencias else "COM PENDÊNCIAS"
    resumo.append(["Item", "Valor"])
    resumo.append(["Status", status])
    resumo.append(["Arquivo fonte", str(caminho_fonte)])
    resumo.append(["Total de profissionais", len(colaboradores)])
    resumo.append(["Total de pendências", len(pendencias)])
    resumo.append(["Gerado em", datetime.now().strftime("%d/%m/%Y %H:%M:%S")])
    _formatar_cabecalho(resumo)
    _ajustar_larguras(resumo)

    detalhes = workbook.create_sheet("Detalhes")
    detalhes.append(["Status", "Linha origem", "Nome", "CPF", "Campo", "Mensagem"])

    if pendencias:
        for pendencia in pendencias:
            detalhes.append(
                [
                    "PENDÊNCIA",
                    pendencia.linha_origem,
                    pendencia.nome,
                    pendencia.cpf,
                    pendencia.campo,
                    pendencia.mensagem,
                ]
            )
    else:
        for colaborador in colaboradores:
            detalhes.append(
                [
                    "OK",
                    colaborador.get("linha_origem", ""),
                    colaborador.get("nome_completo", ""),
                    colaborador.get("cpf", ""),
                    "-",
                    "Registro validado com sucesso.",
                ]
            )

    _formatar_cabecalho(detalhes)
    for row in detalhes.iter_rows(min_row=2):
        row[0].fill = PatternFill("solid", fgColor=VERDE if row[0].value == "OK" else VERMELHO)
        row[0].font = Font(bold=True, color=TEXTO_CLARO)
    _ajustar_larguras(detalhes)

    workbook.save(saida)
    return saida
