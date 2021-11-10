from numpy import string_
import pandas as pd
from collections import OrderedDict

WB_PATH = './data/Planilha para Resolução de Exercícios.xlsx'
WB_SHEET = 'Marco'


def main():
    # Le a planilha e remove todas as colunas e linhas em branco
    df = pd.read_excel(WB_PATH, sheet_name=WB_SHEET)

    df = df.dropna(axis='columns', how='all')
    df = df.dropna(how='all')

    # Define o cabeçalho do DataFrame
    df.columns = df.iloc[0]

    # Renomeia a primeira coluna do cabeçalho
    df = df.rename(columns={'VISITANTES - 2017': 'TAG'})

    # Remove a coluna TOTAL
    df = df.drop('TOTAL', axis='columns')

    # Reseta o index para realizar iteraçãa
    df = df.reset_index(drop=True)

    # Lista vazia onde serão adicionados os indexes da linhas
    # a serem removidas
    to_drop = []
    to_remove = ['TOTAL']

    for index, row in df.iterrows():
        if row['TAG'] in to_remove:
            to_drop.append(index)

    # Remove indexes duplicados e remove do DataFrame os indexes
    to_drop = list(dict.fromkeys(to_drop))
    df = df.drop(to_drop)

    # Reseta o index para realizar iteraçãa
    df = df.reset_index(drop=True)

    year = 0
    for index, row in df.iterrows():
        if 'VISITANTES' in row['TAG']:
            year = str(row['TAG'][-4:]).strip()
            continue

        df.loc[index, 'ANO'] = year

    to_drop = []
    for index, row in df.iterrows():
        if 'VISITANTES' in row['TAG']:
            to_drop.append(index)

    to_drop = list(dict.fromkeys(to_drop))
    df = df.drop(to_drop)

    df_brasileiros = pd.DataFrame()
    df_mercosul = pd.DataFrame()
    df_estrangeiros = pd.DataFrame()
    df_moradores = pd.DataFrame()
    df_isentos = pd.DataFrame()

    for index, row in df.iterrows():
        if row['TAG'] == 'BRASILEIROS':
            df_brasileiros = df_brasileiros.append(df.loc[[index]])

        if row['TAG'] == 'MERCOSUL':
            df_mercosul = df_mercosul.append(df.loc[[index]])

        if row['TAG'] == 'ESTRANGEIROS':
            df_estrangeiros = df_estrangeiros.append(df.loc[[index]])

        if row['TAG'] == 'MORADORES':
            df_moradores = df_moradores.append(df.loc[[index]])

        if row['TAG'] == 'Tripulantes(Isentos)':
            df_isentos = df_isentos.append(df.loc[[index]])

    df_brasileiros = df_brasileiros.melt(
        id_vars=['ANO', 'TAG'], var_name='MES', value_name='QUANTIDADE BRASILEIROS')

    df_brasileiros = df_brasileiros.drop('TAG', axis='columns')

    df_mercosul = df_mercosul.melt(
        id_vars=['ANO', 'TAG'], var_name='MES', value_name='QUANTIDADE MERCOSUL')

    df_mercosul = df_mercosul.drop('TAG', axis='columns')
    
    df_estrangeiros = df_estrangeiros.melt(
        id_vars=['ANO', 'TAG'], var_name='MES', value_name='QUANTIDADE ESTRANGEIROS')

    df_estrangeiros = df_estrangeiros.drop('TAG', axis='columns')
    
    df_moradores = df_moradores.melt(
        id_vars=['ANO', 'TAG'], var_name='MES', value_name='QUANTIDADE MORADORES')

    df_moradores = df_moradores.drop('TAG', axis='columns')
    
    df_isentos = df_isentos.melt(
        id_vars=['ANO', 'TAG'], var_name='MES', value_name='QUANTIDADE ISENTOS')

    df_isentos = df_isentos.drop('TAG', axis='columns')
    
    # dfs = [df_brasileiros, df_mercosul,
    #        df_estrangeiros, df_moradores, df_isentos]

    # df = pd.DataFrame(columns=['TAG', 'MES', 'ANO'])
    # for _df in dfs:
    #     df = df.merge(_df, how='left', on=['TAG', 'MES', 'ANO'])

    df = pd.merge(df_brasileiros, df_mercosul,
                  how='left', on=['MES', 'ANO'])

    df = df.merge(df_estrangeiros, how='left', on=['MES', 'ANO'])
    df = df.merge(df_moradores, how='left', on=['MES', 'ANO'])
    df = df.merge(df_isentos, how='left', on=['MES', 'ANO'])

    df = df.fillna(0)
    
    print(df)

    # Salva no excel a planilha
    df.to_excel('./data/treated_02.xlsx', sheet_name='Marco', index=False)


if __name__ == '__main__':
    main()
