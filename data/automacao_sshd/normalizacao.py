from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any


CARACTERES_INVALIDOS_ABA = r'[\[\]\*\?/:\\]'


def normalizar_texto(valor: Any) -> str:
    if valor is None:
        return ""
    texto = str(valor).strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def valor_vazio(valor: Any) -> bool:
    if valor is None:
        return True
    if isinstance(valor, str):
        return not valor.strip()
    return False


def limpar_cabecalho(valor: Any) -> str:
    texto = normalizar_texto(valor).casefold()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(caractere for caractere in texto if not unicodedata.combining(caractere))
    texto = re.sub(r"[^a-z0-9]+", "_", texto)
    return texto.strip("_")


def valor_excel(valor: Any) -> Any:
    if valor_vazio(valor):
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, float) and valor.is_integer():
        return int(valor)
    if isinstance(valor, Decimal) and valor == valor.to_integral_value():
        return int(valor)
    if isinstance(valor, str):
        return normalizar_texto(valor)
    return valor


def nome_arquivo_seguro(nome: str, padrao: str = "SSHD_PREENCHIDO_FINAL") -> str:
    texto = normalizar_texto(nome) or padrao
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(caractere for caractere in texto if not unicodedata.combining(caractere))
    texto = re.sub(r"[^A-Za-z0-9_-]+", "_", texto)
    texto = re.sub(r"_+", "_", texto).strip("_")
    return texto or padrao


def nome_aba_seguro(nome: str, indice: int, nomes_usados: set[str]) -> str:
    base = re.sub(CARACTERES_INVALIDOS_ABA, "", normalizar_texto(nome)) or f"Cadastro_{indice}"
    base = base[:31]

    candidato = base
    contador = 2
    while candidato in nomes_usados:
        sufixo = f"_{contador}"
        candidato = f"{base[:31 - len(sufixo)]}{sufixo}"
        contador += 1

    nomes_usados.add(candidato)
    return candidato


def proximo_caminho_disponivel(caminho: Path) -> Path:
    if not caminho.exists():
        return caminho

    for indice in range(2, 1000):
        candidato = caminho.with_name(f"{caminho.stem}_{indice}{caminho.suffix}")
        if not candidato.exists():
            return candidato

    raise FileExistsError(f"Nao foi possivel encontrar um nome livre para: {caminho}")
