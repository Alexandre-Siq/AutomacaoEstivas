from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_PADRAO = BASE_DIR / "TEMPLATE_NOVO.xlsx"
ABA_TEMPLATE = "SSHD"

COLUNAS_FONTE_DADOS = {
    "registro": {
        "cadastro_mv",
        "crm",
        "crm_",
        "registro",
        "registro_do_funcionario",
        "registro_funcionario",
    },
    "nome_completo": {
        "colaborador",
        "funcionario",
        "nome",
        "nome_colaborador",
        "nome_completo",
        "nome_do_colaborador",
        "nome_do_funcionario",
        "nome_funcionario",
    },
    "data_nascimento": {
        "data_de_nascimento",
        "data_nascimento",
        "dt_nascimento",
        "nascimento",
    },
    "estado_civil": {
        "estado_civil",
    },
    "naturalidade": {
        "naturalidade",
    },
    "nacionalidade": {
        "nacionalidade",
    },
    "cpf": {
        "cpf",
        "cpf_colaborador",
        "cpf_do_colaborador",
        "cpf_funcionario",
    },
    "cargo_funcao": {
        "cargo",
        "cargo_funcao",
        "cargo_ou_funcao",
        "especialidade",
        "funcao",
    },
    "nome_mae": {
        "mae",
        "nome_completo_da_mae",
        "nome_da_mae",
        "nome_de_mae",
    },
    "logradouro": {
        "endereco",
        "logradouro",
    },
    "numero": {
        "n",
        "no",
        "numero",
    },
    "complemento": {
        "complemento",
    },
    "bairro": {
        "bairro",
    },
    "cidade": {
        "cidade",
    },
    "estado": {
        "estado",
        "uf",
    },
    "cep": {
        "cep",
    },
    "email_profissional": {
        "e_mail",
        "email",
        "mail",
    },
    "regime": {
        "regime",
        "tipo_de_prestador",
        "tipo_prestador",
    },
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

VALORES_FIXOS_CAMPOS = {
    "genero": "não informado",
    "orientacao_sexual": "não informado",
}

VALORES_PADRAO = {
    "regime": "CLT",
}
