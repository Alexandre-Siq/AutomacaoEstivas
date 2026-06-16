from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_PADRAO = BASE_DIR / "TEMPLATE_NOVO.xlsx"
ABA_TEMPLATE = "SSHD"

COLUNAS_PLANILHA_GERAL = {
    "registro": "B",
    "nome_completo": "C",
    "data_nascimento": "D",
    "genero": "E",
    "estado_civil": "F",
    "naturalidade": "G",
    "nacionalidade": "H",
    "cpf": "I",
    "cargo_funcao": "J",
    "nome_mae": "O",
    "logradouro": "P",
    "numero": "Q",
    "complemento": "R",
    "bairro": "S",
    "cidade": "T",
    "estado": "U",
    "cep": "V",
    "email_profissional": "AC",
}

CELULAS_PREFEITURA = {
    "nome_completo": "B11",
    "data_nascimento": "B12",
    "cpf": "B13",
    "genero": "B14",
    "orientacao_sexual": "B15",
    "estado_civil": "B16",
    "nome_mae": "B17",
    "nacionalidade": "B18",
    "naturalidade": "B19",
    "email_profissional": "B20",
    "logradouro": "B22",
    "numero": "B23",
    "complemento": "B24",
    "bairro": "B25",
    "cidade": "B26",
    "estado": "B27",
    "cep": "B28",
    "regime": "B30",
    "cargo_funcao": "B31",
    "registro": "B32",
}

CELULAS_SOLICITANTE = {
    "nome": "B5",
    "sshd": "B6",
    "cargo": "B7",
}

VALORES_FIXOS = {
    "B34": "COMPLEXO HOSPITALAR DOS ESTIVADORES",
    "B35": "SMS",
    "B37": "COMPLEXO HOSPITALAR DOS ESTIVADORES",
    "B38": "SMS",
}

CAMPOS_OBRIGATORIOS_PREFEITURA = {
    "nome_completo": "Nome completo",
    "data_nascimento": "Data de nascimento",
    "cpf": "CPF",
    "genero": "Genero",
    "orientacao_sexual": "Orientacao Sexual",
    "estado_civil": "Estado Civil",
    "nome_mae": "Nome da mae",
    "nacionalidade": "Nacionalidade",
    "naturalidade": "Naturalidade",
    "email_profissional": "Email profissional",
    "logradouro": "Logradouro",
    "numero": "Numero",
    "bairro": "Bairro",
    "cidade": "Cidade",
    "estado": "Estado",
    "cep": "CEP",
    "regime": "Regime",
    "cargo_funcao": "Cargo/funcao",
    "registro": "Registro",
}

VALORES_PADRAO = {
    "genero": "Nao informado",
    "orientacao_sexual": "Nao Informado",
    "regime": "CLT",
}
