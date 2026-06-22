# Automacao de Fichas SSHD

Projeto para gerar fichas de cadastramento SSHD no modelo padrão da Prefeitura a
partir da planilha fonte de colaboradores.

## Situação atual

Esta é a **Versão 2.0** do aplicativo e implementa o fluxo da **planilha fonte**:

1. O usuário seleciona a planilha de colaboradores em `.xlsx`.
2. O usuário escolhe a pasta de saída, ou mantém a mesma pasta da fonte.
3. O usuário preenche:
   - Nome do Solicitante
   - SSHD
   - Cargo
4. A aplicação gera um arquivo Excel com uma aba por colaborador.
5. A aplicação gera um relatório de validação em Excel.
6. Cada aba preserva o layout do template `data/TEMPLATE_NOVO.xlsx`.

A interface utiliza tema escuro com destaque verde, fonte Oswald e janela fixa
compacta de 680x700, com status de processamento, ação para limpar os campos
antes de uma nova geração, versão visível e o crédito `Desenvolvido por
Alexandre Siqueira - Analista de Suporte`.

Após a geração, o botão `Abrir arquivo` fica habilitado para abrir diretamente
o Excel final.

O executável gerado com PyInstaller exibe uma splash screen durante o
carregamento inicial e fecha automaticamente quando a janela principal fica
pronta.

Para gerar o executável com splash screen, o Python usado no build precisa ter
`tkinter` disponível. No Windows, isso normalmente já vem no instalador oficial
do Python.

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

## Relatório de validação

A cada execução é criado um relatório ao lado do arquivo final, com o sufixo
`_RELATORIO_VALIDACAO.xlsx`.

O relatório contém:

- status geral da execução;
- caminho da fonte usada;
- total de profissionais encontrados;
- total de pendências;
- detalhes por linha/profissional.

Se houver pendências obrigatórias, a geração das fichas é bloqueada e o
relatório informa quais campos precisam ser corrigidos.

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
- `data/splash_sshd.png`: imagem exibida na abertura do executável.
- `data/automacao_sshd/`: nucleo de leitura, validacao, mapeamento e escrita.

## Geracao de executavel

No Windows, dentro da pasta `data`, execute:

```bash
python -m PyInstaller gerador_planilhas.spec
```

Para abrir mais rapido, o projeto usa o modo `onedir` do PyInstaller. O
resultado fica em:

```text
data/dist/gerador_planilhas/
```

Distribua a pasta `gerador_planilhas` inteira para a equipe e oriente o uso do
arquivo:

```text
data/dist/gerador_planilhas/gerador_planilhas.exe
```

Nao envie somente o `.exe`, pois as dependencias ficam na mesma pasta.

Antes de compactar, copie tambem o arquivo `data/LEIA-ME_USUARIOS.txt` para
dentro da pasta `data/dist/gerador_planilhas/`, para que os usuarios tenham as
instrucoes de uso junto do executavel.
