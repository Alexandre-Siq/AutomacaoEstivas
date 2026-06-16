# Automacao de Fichas SSHD

Projeto para gerar fichas de cadastramento SSHD no modelo padrao da Prefeitura a
partir da Planilha Geral de colaboradores.

## Situacao atual

Esta primeira versao implementa o fluxo da **Planilha Geral**:

1. O usuario seleciona a planilha de colaboradores em `.xlsx`.
2. O usuario preenche:
   - Nome do Solicitante
   - SSHD
   - Cargo
3. A aplicacao gera um arquivo Excel com uma aba por colaborador.
4. Cada aba preserva o layout do template `data/TEMPLATE_NOVO.xlsx`.

A Planilha de Medicos sera adicionada como uma segunda origem assim que o
de/para especifico for confirmado.

## Campos fixos no modelo da Prefeitura

Os campos abaixo sao preenchidos automaticamente em todas as fichas:

- `COMPLEXO HOSPITALAR DOS ESTIVADORES`
- `SMS`

## Mapeamento da Planilha Geral

| Campo Prefeitura | Origem Planilha Geral |
| --- | --- |
| Nome completo | Nome Colaborador |
| Data de nascimento | Data Nascimento |
| CPF | CPF |
| Genero | Sexo |
| Orientacao Sexual | Valor padrao: Nao Informado |
| Estado Civil | Estado Civil |
| Nome da mae | Nome Completo da Mae |
| Nacionalidade | Nacionalidade |
| Naturalidade | Naturalidade |
| Email profissional | E-MAIL |
| Logradouro | Endereco |
| Numero | Nº da coluna Q |
| Complemento | Complemento |
| Bairro | Bairro |
| Cidade | Cidade |
| Estado | UF |
| CEP | CEP |
| Regime | Valor padrao: CLT |
| Cargo/funcao | Cargo |
| Registro | Registro do Funcionario |

## Como executar em modo desenvolvimento

```bash
python3 -m pip install -r requirements.txt
python3 data/gerador_planilhas.py
```

## Arquivos principais

- `data/gerador_planilhas.py`: interface grafica.
- `data/TEMPLATE_NOVO.xlsx`: template oficial usado como base.
- `data/criar_template_prefeitura.py`: recria o template a partir do modelo
  original.
- `data/automacao_sshd/`: nucleo de leitura, validacao, mapeamento e escrita.

## Geracao de executavel

No Windows, dentro da pasta `data`, execute:

```bash
pyinstaller gerador_planilhas.spec
```

O arquivo gerado deve incluir o template e o icone usados pela interface.
