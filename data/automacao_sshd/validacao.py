from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .config import CAMPOS_OBRIGATORIOS_PREFEITURA
from .exceptions import ErroValidacao
from .normalizacao import valor_vazio


@dataclass(frozen=True)
class PendenciaValidacao:
    linha_origem: Any
    nome: str
    cpf: Any
    campo: str
    mensagem: str


def validar_campos_solicitante(nome: str, sshd: str, cargo: str) -> None:
    campos = {
        "Nome do Solicitante": nome,
        "SSHD": sshd,
        "Cargo": cargo,
    }
    ausentes = [campo for campo, valor in campos.items() if valor_vazio(valor)]

    if ausentes:
        raise ErroValidacao(
            "Preencha os campos obrigatorios do solicitante: " + ", ".join(ausentes)
        )


def coletar_pendencias_colaboradores(
    colaboradores: list[dict[str, Any]],
) -> list[PendenciaValidacao]:
    pendencias: list[PendenciaValidacao] = []

    for colaborador in colaboradores:
        linha = colaborador.get("linha_origem", "?")
        nome = colaborador.get("nome_completo") or "sem nome"
        cpf = colaborador.get("cpf") or ""

        for campo, rotulo in CAMPOS_OBRIGATORIOS_PREFEITURA.items():
            if valor_vazio(colaborador.get(campo)):
                pendencias.append(
                    PendenciaValidacao(
                        linha_origem=linha,
                        nome=str(nome),
                        cpf=cpf,
                        campo=rotulo,
                        mensagem=f"Campo obrigatório ausente: {rotulo}",
                    )
                )

    return pendencias


def validar_colaboradores(
    colaboradores: list[dict[str, Any]],
    pendencias: list[PendenciaValidacao] | None = None,
) -> None:
    pendencias = pendencias if pendencias is not None else coletar_pendencias_colaboradores(colaboradores)

    if pendencias:
        exibidas = pendencias[:20]
        restante = len(pendencias) - len(exibidas)
        mensagem = "\n".join(
            f"Linha {pendencia.linha_origem} ({pendencia.nome}): {pendencia.mensagem}"
            for pendencia in exibidas
        )
        if restante > 0:
            mensagem += f"\n... e mais {restante} pendencia(s)."
        raise ErroValidacao(mensagem)
