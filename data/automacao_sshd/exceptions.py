class ErroAutomacao(Exception):
    """Erro base tratado pela interface da automacao."""


class ErroLeituraPlanilha(ErroAutomacao):
    """Erro ao abrir ou interpretar a planilha de origem."""


class ErroValidacao(ErroAutomacao):
    """Erro causado por campos obrigatorios ausentes ou inconsistentes."""
