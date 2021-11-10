from numpy import string_
import pandas as pd
from collections import OrderedDict

WB_PATH = './data/Planilha para Resolução de Exercícios.xlsx'
WB_SHEET = 'Ecomuseu'


def main():
    # Le a planilha e remove todas as colunas e linhas em branco
    df = pd.read_excel(WB_PATH, sheet_name=WB_SHEET)

    df = df.dropna(axis='columns', how='all')
    df = df.dropna(how='all')

    # Define o cabeçalho do DataFrame
    df.columns = df.iloc[0]

    # Remove a coluna TOTAL
    df = df.drop('TOTAL', axis='columns')

    # Reseta o index para realizar iteraçãa
    df = df.reset_index(drop=True)

    # Lista vazia onde serão adicionados os indexes da linhas
    # a serem removidas
    to_drop = []
    to_keep = ['NACIONAIS', 'ESTRANGEIROS']

    for index, row in df.iterrows():
        if isinstance(row['ANO'], str):
            to_drop.append(index)

        if row['VISITANTES'] not in to_keep:
            to_drop.append(index)

    # Remove indexes duplicados e remove do DataFrame os indexes
    to_drop = list(dict.fromkeys(to_drop))
    df = df.drop(to_drop)

    # Reseta o index para realizar iteraçãa
    df = df.reset_index(drop=True)

    df_nacionais = pd.DataFrame()
    df_estrangeiros = pd.DataFrame()

    for index, row in df.iterrows():
        if pd.isna(row['ANO']):
            df.loc[index, 'ANO'] = df.loc[index + 1, 'ANO']

        if row['VISITANTES'] == 'NACIONAIS':
            df_nacionais = df_nacionais.append(df.loc[[index]])

        if row['VISITANTES'] == 'ESTRANGEIROS':
            df_estrangeiros = df_estrangeiros.append(df.loc[[index]])

    df_nacionais = df_nacionais.drop('VISITANTES', axis='columns')
    df_nacionais = df_nacionais.melt(
        id_vars=['ANO'], var_name='MES', value_name='QUANTIDADE NACIONAL')

    df_estrangeiros = df_estrangeiros.drop('VISITANTES', axis='columns')
    df_estrangeiros = df_estrangeiros.melt(
        id_vars=['ANO'], var_name='MES', value_name='QUANTIDADE ESTRANGEIROS')

    df = pd.merge(df_nacionais, df_estrangeiros, how='left', on=['ANO', 'MES'])

    df = df.replace('...', 0)

    print(df)

    # Salva no excel a planilha
    df.to_excel('./data/treated_03.xlsx', sheet_name='Ecomuseu', index=False)


if __name__ == '__main__':
    main()
