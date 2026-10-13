<img width="1818" height="903" alt="image" src="https://github.com/user-attachments/assets/5c51348c-1c2c-48ce-883c-081751d19c7d" /># Automação de Cadastro no CADIN

Automação em duas etapas para o cadastramento em massa de responsáveis no
[CADIN/PGFN](https://cadin.pgfn.gov.br), a partir da planilha de
reprocessamento:

1. **Python** lê a planilha de reprocessamento, monta as colunas `CPF`,
   `Referencia`, `Complemento` e `Data` e gera `dados_processados_cadastro.xlsx`.
2. **Power Automate Desktop** lê essa planilha gerada, faz login no CADIN com
   gov.br e preenche o formulário de cadastro, linha por linha.

```
Planilha de reprocessamento ──► processa_cadastro.py ──► dados_processados_cadastro.xlsx ──► Power Automate ──► CADIN
```

## Estrutura do repositório

```
automacao-cadin/
├── python/
│   ├── processa_cadastro.py        # Etapa 1: tratamento da planilha
│   └── requirements.txt
├── power-automate/
│   └── fluxo_cadastro_cadin.txt    # Etapa 2: fluxo exportado do Power Automate Desktop
├── docs/img/                       # Prints do fluxo usados neste README
├── .gitignore
└── README.md
```

> As planilhas (`.xlsx`) **não** fazem parte do repositório, porque contêm
> CPF/CNPJ e nomes de responsáveis. O `.gitignore` já impede que elas sejam
> enviadas.

## Pré-requisitos

- Windows com **Power Automate Desktop** instalado e a extensão do Power
  Automate no **Google Chrome**.
- **Python 3.11 ou superior** (testado com Python 3.13). Baixe em
  https://www.python.org/downloads/. No instalador, marque a opção
  **"Add python.exe to PATH"**.
- Conta gov.br com acesso ao CADIN.

---

## Etapa 1: Processar a planilha (Python)

1. Abra o terminal (PowerShell) na pasta `python/`. No Explorador de
   Arquivos, clique com o botão direito dentro da pasta e escolha
   "Abrir no Terminal".

2. *(Opcional, recomendado)* Crie e ative um ambiente virtual:
   ```powershell
   python -m venv venv
   venv\Scripts\activate
   ```

3. Instale as dependências:
   ```powershell
   pip install -r requirements.txt
   ```

4. Coloque a planilha de entrada **na pasta `python/`**, com o nome exato:
   ```
   C - REPROCESSAMENTO Cadin-PGFN - 2026 - Dados para ROBÔ.xlsx
   ```
   Se o nome do arquivo ou da aba mudar, ajuste as constantes
   `ARQUIVO_ENTRADA` e `ABA` no topo de `processa_cadastro.py`.

5. Rode o script:
   ```powershell
   python processa_cadastro.py
   ```

6. O arquivo `dados_processados_cadastro.xlsx` é gerado na mesma pasta. O
   script mostra quantas linhas foram processadas e avisa se alguma ficou
   com data vazia ou inválida.

### O que o script faz

| Coluna gerada | Como é montada |
|---|---|
| `CPF` | Coluna F, só com dígitos e completada com zeros à esquerda (11 dígitos para CPF, 14 para CNPJ) |
| `Referencia` | `PR` + Nº Processo (coluna I) |
| `Complemento` | Instrumento + Nº Instr. + `SF` + SIAFI + `TC` + Processo TC + Julgamento (`ACORDAO` → `AC`) + Número + Seção |
| `Data` | Data de inscrição (coluna H) no formato `dd/mm/aaaa` |

Linhas sem CPF/CNPJ (por exemplo, linhas em branco no fim da planilha) são
descartadas.

### Estrutura esperada da planilha de entrada

Aba `Resp. ATIVOS e CONTROLE Cadin ` (repare no espaço no final do nome),
colunas **F até P**, com 4 linhas de cabeçalho antes dos títulos:

| Coluna | Campo |
|---|---|
| F | CPF/CNPJ |
| G | Nome Responsável |
| H | Data Inscrição no CADIN |
| I | Nº Processo |
| J | Instrumento |
| K | Nº Instr. |
| L | SIAFI |
| M | Processo TC |
| N | Julgamento |
| O | Número |
| P | Seção |

Se a estrutura da planilha mudar (colunas fora de ordem, campos novos, aba
renomeada), ajuste os índices `iloc[:, N]` ou o `usecols`/`ABA` no script.

---

## Etapa 2: Cadastrar no CADIN (Power Automate Desktop)

### 2.1 Salvar o login do gov.br no Windows (uma vez só)

O fluxo **não guarda usuário nem senha**. Ele lê os dois do Gerenciador de
Credenciais do Windows, numa credencial chamada `PowerAutomate_GovBR`.

Para criá-la, rode no PowerShell (troque pelo seu CPF, só números):

```powershell
cmdkey /generic:PowerAutomate_GovBR /user:SEU_CPF /pass
```

O comando pede a senha do gov.br. Ela não aparece na tela enquanto você
digita. Para conferir ou editar depois, abra **Painel de Controle →
Gerenciador de Credenciais → Credenciais do Windows**.

### 2.2 Importar o fluxo

1. No Power Automate Desktop, clique em **+ Novo fluxo** e dê um nome
   (ex.: "Automação do CADIN").
2. Abra `power-automate/fluxo_cadastro_cadin.txt` no Bloco de Notas, selecione
   tudo (**Ctrl+A**) e copie (**Ctrl+C**).
3. No designer do fluxo, clique na área vazia da aba **Main** e cole
   (**Ctrl+V**). As ações e os elementos de interface são criados
   automaticamente.
4. Abra a primeira ação, **Iniciar o Excel**, e troque o caminho
   `C:\CAMINHO\PARA\dados_processados_cadastro.xlsx` pelo caminho real do
   arquivo gerado na Etapa 1.
5. Na ação **Ler da planilha do Excel**, ajuste a linha final (`EndRow`) para
   a quantidade de linhas da planilha gerada + 1 (cabeçalho).
6. Salve o fluxo.

### 2.3 Executar

Clique em **Executar**. O robô:

1. lê a planilha gerada e fecha o Excel;
2. abre o Chrome no CADIN e entra com gov.br usando a credencial salva;
3. acessa **Cadastro → Incluir Cadastro**;
4. para cada linha da planilha, preenche CPF, Referência, Complemento e Data.

> **Dica:** antes de rodar a planilha inteira, coloque um ponto de
> interrupção (clique no número da linha) no `Para cada` e avance com
> **Executar próxima ação** para conferir a primeira linha.

### Visão geral do fluxo
<img width="1818" height="903" alt="Captura de tela 2026-10-13 110318" src="https://github.com/user-attachments/assets/a565dde0-b22e-450f-987b-742780e53993" />

<img width="1812" height="899" alt="Captura de tela 2026-10-13 110446" src="https://github.com/user-attachments/assets/2a1187e8-1cd7-4bcf-a666-0bec610a8600" />




---

## Problemas comuns

- **`ValueError: Worksheet named '...' not found`**: o nome da aba no Excel
  não bate com o do script. Para listar os nomes reais das abas:
  ```python
  import pandas as pd
  print(pd.ExcelFile('NOME_DO_ARQUIVO.xlsx').sheet_names)
  ```

- **`FileNotFoundError`**: a planilha de entrada não está na pasta `python/`
  ou o nome está diferente. Rode `dir` na pasta para conferir.

- **O robô não preenche o CPF ou a senha no gov.br**: confira se a
  credencial `PowerAutomate_GovBR` existe no Gerenciador de Credenciais com
  usuário **e** senha preenchidos.

- **O robô não encontra um botão ou campo**: o site do CADIN ou do gov.br
  pode ter mudado. Abra o painel **Elementos de IU** no Power Automate,
  localize o elemento com problema e capture-o de novo.
