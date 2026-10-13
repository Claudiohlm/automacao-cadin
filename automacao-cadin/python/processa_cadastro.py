"""
Etapa 1 da automacao CADIN.

Le a planilha de reprocessamento do CADIN/PGFN e gera
'dados_processados_cadastro.xlsx' com as colunas CPF, Referencia,
Complemento e Data, que o fluxo do Power Automate (etapa 2) usa
para cadastrar cada linha no CADIN.
"""
from pathlib import Path

import pandas as pd

# Os arquivos ficam na mesma pasta do script, de onde quer que ele seja
# executado (terminal, duplo clique ou chamado pelo Power Automate).
PASTA = Path(__file__).resolve().parent
ARQUIVO_ENTRADA = PASTA / 'C - REPROCESSAMENTO Cadin-PGFN - 2026 - Dados para ROBÔ.xlsx'
ABA = 'Resp. ATIVOS e CONTROLE Cadin '  # atencao: tem um espaco no final
ARQUIVO_SAIDA = PASTA / 'dados_processados_cadastro.xlsx'


def formatar_cpf_cnpj(valor):
    """
    Garante o tamanho correto do documento (11 digitos p/ CPF, 14 p/ CNPJ),
    mesmo que a celula tenha perdido algum zero a esquerda.
    Diferenciacao feita pela quantidade de digitos, pois nao ha coluna
    indicando o tipo do documento.
    """
    if pd.isna(valor):
        return ''
    digitos = ''.join(c for c in str(valor).strip() if c.isdigit())
    if not digitos:
        return ''
    if len(digitos) <= 11:
        return digitos.zfill(11)
    return digitos.zfill(14)


def main():
    # 1. Leitura (colunas F a P, pulando as 4 linhas de cabecalho)
    planilha = pd.read_excel(
        ARQUIVO_ENTRADA,
        sheet_name=ABA,
        usecols='F:P',
        skiprows=4,
        dtype=str,
    )

    # Celulas vazias viram texto vazio, e nao 'nan', na hora de concatenar
    planilha = planilha.fillna('')

    # 2. Processamento das colunas
    planilha['CPF'] = planilha.iloc[:, 0].apply(formatar_cpf_cnpj)

    # Descarta linhas sem documento (ex.: linhas em branco no fim da planilha),
    # para o robo nao tentar cadastrar registros vazios
    planilha = planilha[planilha['CPF'] != ''].copy()

    planilha['Referencia'] = 'PR' + planilha.iloc[:, 3].str.strip()  # I - No Processo

    julgamento = planilha.iloc[:, 8].str.replace('ACORDAO', 'AC', regex=False)

    planilha['Complemento'] = (
        planilha.iloc[:, 4]      # J - Instrumento
        + planilha.iloc[:, 5]    # K - No Instr.
        + 'SF'
        + planilha.iloc[:, 6]    # L - SIAFI
        + 'TC'
        + planilha.iloc[:, 7]    # M - Processo TC
        + julgamento             # N - Julgamento (abreviado p/ AC)
        + planilha.iloc[:, 9]    # O - Numero
        + planilha.iloc[:, 10]   # P - Secao
    )

    planilha['Data'] = (
        pd.to_datetime(planilha.iloc[:, 2], errors='coerce')  # H - Data Inscricao
        .dt.strftime('%d/%m/%Y')
        .fillna('')
    )

    # 3. Gravacao
    planilha_final = planilha[['CPF', 'Referencia', 'Complemento', 'Data']]
    planilha_final.to_excel(ARQUIVO_SAIDA, index=False)

    print(f'{len(planilha_final)} linhas processadas.')
    sem_data = (planilha_final['Data'] == '').sum()
    if sem_data:
        print(f'Atencao: {sem_data} linha(s) com data vazia ou invalida. Confira antes de rodar o robo.')
    print(f'Arquivo gerado: {ARQUIVO_SAIDA}')


if __name__ == '__main__':
    main()
