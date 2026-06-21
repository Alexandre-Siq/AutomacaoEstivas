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

A interface utiliza tema escuro com destaque verde, fonte Oswald e janela fixa
compacta de 680x580, com status de processamento, ação para limpar os campos
antes de uma nova geração e créditos para Alexandre Siqueira, Analista de
Suporte.

A Planilha de Medicos sera adicionada como uma segunda origem assim que o
de/para especifico for confirmado.

## Campos fixos no modelo da Prefeitura

Os campos abaixo sao preenchidos automaticamente em todas as fichas:

- `COMPLEXO HOSPITALAR DOS ESTIVADORES`
- `SMS`

## Mapeamento da Planilha Geral

| Campo Prefeitura | Origem da planilha fonte |
| --- | --- |
| Nome completo | Nome completo / Nome Colaborador |
| Data de nascimento | Data de nascimento / Data Nascimento |
| CPF | CPF |
| Genero | Valor fixo: não informado |
| Orientacao Sexual | Valor fixo: não informado |
| Estado Civil | Estado Civil |
| Nome da mae | Nome da mãe / Nome Completo da Mae |
| Nacionalidade | Nacionalidade |
| Naturalidade | Naturalidade |
| Email profissional | E-mail / E-MAIL |
| Logradouro | Logradouro / Endereco |
| Numero | Nº ou extraido do Logradouro |
| Complemento | Complemento |
| Bairro | Bairro |
| Cidade | Cidade |
| Estado | Estado / UF |
| CEP | CEP |
| Regime | Tipo de Prestador; se ausente, CLT |
| Cargo/funcao | Especialidade / Cargo |
| Registro | CRM / Registro do Funcionario / Cadastro MV |

Os campos `Genero` e `Orientacao Sexual` sao sempre preenchidos como
`não informado`, independentemente do conteudo da planilha de origem.

O arquivo `data/exemplo_fonte_dados.xlsx` representa o formato atualmente
esperado para a fonte de dados.

## Como executar em modo desenvolvimento

```bash
python3 -m pip install -r requirements.txt
python3 data/gerador_planilhas.py
```

## Solucao de problemas

### Erro: nao foi possivel localizar o cabecalho da Planilha Geral

O leitor procura, nas primeiras 50 linhas de todas as abas, colunas equivalentes
a:

- Nome Colaborador
- CPF
- Cargo

Tambem sao aceitas variacoes como `Nome do Colaborador`, `Nome Completo`,
`Cargo/Funcao` e `Funcao`.

Se o erro aparecer, confira se a planilha selecionada e realmente a Planilha
Geral de colaboradores e se esses campos aparecem em uma linha de cabecalho
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
