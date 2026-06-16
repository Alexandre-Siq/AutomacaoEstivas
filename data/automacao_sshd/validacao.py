from __future__ import annotations

from typing import Any

from .config import CAMPOS_OBRIGATORIOS_PREFEITURA
from .exceptions import ErroValidacao
from .normalizacao import valor_vazio


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


def validar_colaboradores(colaboradores: list[dict[str, Any]]) -> None:
    erros: list[str] = []

    for colaborador in colaboradores:
        linha = colaborador.get("linha_origem", "?")
        nome = colaborador.get("nome_completo") or "sem nome"

        for campo, rotulo in CAMPOS_OBRIGATORIOS_PREFEITURA.items():
            if valor_vazio(colaborador.get(campo)):
                erros.append(f"Linha {linha} ({nome}): campo obrigatorio ausente - {rotulo}")

    if erros:
        exibidos = erros[:20]
        restante = len(erros) - len(exibidos)
        mensagem = "\n".join(exibidos)
        if restante > 0:
            mensagem += f"\n... e mais {restante} pendencia(s)."
        raise ErroValidacao(mensagem)
