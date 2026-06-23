from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .config import TEMPLATE_PADRAO
from .normalizacao import nome_arquivo_seguro
from .planilha_geral import ler_planilha_geral
from .prefeitura_writer import preencher_template_prefeitura
from .relatorio_validacao import gerar_relatorio_validacao
from .validacao import (
    coletar_pendencias_colaboradores,
    validar_campos_solicitante,
    validar_colaboradores,
)


@dataclass(frozen=True)
class DadosSolicitante:
    nome: str
    sshd: str
    cargo: str


@dataclass(frozen=True)
class ResultadoGeracao:
    caminho_saida: Path
    caminho_relatorio: Path
    total_colaboradores: int


def gerar_fichas_sshd(
    *,
    caminho_planilha_geral: str | Path,
    solicitante: DadosSolicitante,
    caminho_template: str | Path = TEMPLATE_PADRAO,
    caminho_saida: str | Path | None = None,
) -> ResultadoGeracao:
    validar_campos_solicitante(solicitante.nome, solicitante.sshd, solicitante.cargo)

    colaboradores = ler_planilha_geral(caminho_planilha_geral)

    caminho_origem = Path(caminho_planilha_geral)
    if caminho_saida is None:
        nome_saida = f"{nome_arquivo_seguro(caminho_origem.stem)}_SSHD.xlsx"
        caminho_saida = caminho_origem.with_name(nome_saida)
    caminho_saida = Path(caminho_saida)

    caminho_relatorio = caminho_saida.with_name("RELATORIO_VALIDACAO_SSHD.xlsx")
    pendencias = coletar_pendencias_colaboradores(colaboradores)
    caminho_relatorio_gerado = gerar_relatorio_validacao(
        caminho_relatorio=caminho_relatorio,
        colaboradores=colaboradores,
        pendencias=pendencias,
        caminho_fonte=caminho_origem,
    )

    try:
        validar_colaboradores(colaboradores, pendencias)
    except Exception as erro:
        raise type(erro)(f"{erro}\n\nRelatório de validação salvo em:\n{caminho_relatorio_gerado}") from erro

    caminho_gerado = preencher_template_prefeitura(
        caminho_template=caminho_template,
        colaboradores=colaboradores,
        nome_solicitante=solicitante.nome.strip(),
        sshd_solicitante=solicitante.sshd.strip(),
        cargo_solicitante=solicitante.cargo.strip(),
        caminho_saida=caminho_saida,
    )

    return ResultadoGeracao(
        caminho_saida=caminho_gerado,
        caminho_relatorio=caminho_relatorio_gerado,
        total_colaboradores=len(colaboradores),
    )
