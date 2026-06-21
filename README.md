# Automacao de Fichas SSHD

Projeto para gerar fichas de cadastramento SSHD no modelo padrão da Prefeitura a
partir da planilha fonte de colaboradores.

## Situacao atual

Esta primeira versão implementa o fluxo da **planilha fonte**:

1. O usuario seleciona a planilha de colaboradores em `.xlsx`.
2. O usuario preenche:
   - Nome do Solicitante
   - SSHD
   - Cargo
3. A aplicação gera um arquivo Excel com uma aba por colaborador.
4. Cada aba preserva o layout do template `data/TEMPLATE_NOVO.xlsx`.

A interface utiliza tema escuro com destaque verde, fonte Oswald e janela fixa
compacta de 680x580, com status de processamento, ação para limpar os campos
antes de uma nova geração e o crédito `Desenvolvido por Alexandre Siqueira -
Analista de Suporte`.

A Planilha de Médicos será adicionada como uma segunda origem assim que o
de/para específico for confirmado.

## Formatos de fonte suportados

A automação aceita os dois formatos atualmente usados no projeto:

- `data/exemplo_fonte_dados.xlsx`: formato novo, com colunas como `Nome completo`,
  `CRM`, `Tipo de Prestador`, `Especialidade` e `Cadastro MV`.
- `data/1304_e_1604.xlsx`: formato do aplicativo raiz, com colunas como
  `Nome Colaborador`, `Registro do Funcionário`, `Cargo`, `Endereço`, `Nº` e
  `E-MAIL`.

Quando a fonte contém mais de uma pessoa, o arquivo final é gerado com uma aba
por profissional, sempre copiando o template oficial antes de preencher os
dados daquela pessoa.

## Campos fixos no modelo da Prefeitura

Os campos abaixo são preenchidos automaticamente em todas as fichas:

- `COMPLEXO HOSPITALAR DOS ESTIVADORES`
- `SMS`

## Mapeamento da planilha fonte

| Campo Prefeitura | Origem da planilha fonte |
| --- | --- |
| Nome completo | Nome completo / Nome Colaborador |
| Data de nascimento | Data de nascimento / Data Nascimento |
| CPF | CPF |
| Gênero | Valor fixo: não informado |
| Orientação Sexual | Valor fixo: não informado |
| Estado Civil | Estado Civil |
| Nome da mãe | Nome da mãe / Nome Completo da Mãe |
| Nacionalidade | Nacionalidade |
| Naturalidade | Naturalidade |
| Email profissional | E-mail / E-MAIL |
| Logradouro | Logradouro / Endereco |
| Número | Nº ou extraído do Logradouro |
| Complemento | Complemento |
| Bairro | Bairro |
| Cidade | Cidade |
| Estado | Estado / UF |
| CEP | CEP |
| Regime | Tipo de Prestador; se ausente, CLT |
| Cargo/função | Especialidade / Cargo |
| Registro | CRM / Registro do Funcionário / Cadastro MV |

Os campos `Gênero` e `Orientação Sexual` são sempre preenchidos como
`não informado`, independentemente do conteúdo da planilha de origem.

## Como executar em modo desenvolvimento

```bash
python3 -m pip install -r requirements.txt
python3 data/gerador_planilhas.py
```

## Solução de problemas

### Erro: não foi possível localizar o cabeçalho da planilha fonte

O leitor procura, nas primeiras 50 linhas de todas as abas, colunas equivalentes
a:

- Nome Colaborador
- CPF
- Cargo

Também são aceitas variações como `Nome do Colaborador`, `Nome Completo`,
`Cargo/Função` e `Função`.

Se o erro aparecer, confira se a planilha selecionada é realmente a fonte
de colaboradores e se esses campos aparecem em uma linha de cabeçalho
antes das linhas de dados.

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
