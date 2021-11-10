import pandas as pd

WB_PATH = './data/Planilha para Resolução de Exercícios.xlsx'
WB_SHEET = 'Itaipu '


def main():
    # Le a planilha e remove todas as colunas e linhas em branco
    df = pd.read_excel(WB_PATH, sheet_name=WB_SHEET)
    df = df.dropna(axis='columns', how='all')
    df = df.dropna(how='all')

    # Define o cabeçalho do DataFrame
    df.columns = df.iloc[0]

    # Remove a colula TOTAL
    df = df.drop('TOTAL', axis='columns')

    # Reseta o index para realizar iteraçãa
    df = df.reset_index(drop=True)

    # Lista vazia onde serão adicionados os indexes da linhas
    # a serem removidas
    to_drop = []

    for index, row in df.iterrows():
        # Adiciona na lista to_drop o index da linha em que
        # 'ANO' for do tipo string
        if isinstance(row['ANO'], str):
            to_drop.append(index)

        # Adiciona na lista to_drop o index da linha em que
        # 'VISITANTES' não for do tipo string
        if not isinstance(row['VISITANTES'], str):
            to_drop.append(index)

        # Adiciona na lista to_drop o index da linha em que
        # 'VISITANTES' for igual a 'TOTAL'
        if row['VISITANTES'] == 'TOTAL':
            to_drop.append(index)

    # Remove indexes duplicados e remove do DataFrame os indexes
    to_drop = list(dict.fromkeys(to_drop))
    df = df.drop(to_drop)

    # Reseta o index para realizar iteraçãa
    df = df.reset_index(drop=True)

    # Adiciona o valor de 'ANO' da linha seguinte caso a linha
    # atual seja NAN
    for index, row in df.iterrows():
        year_col = 'ANO'
        year_cell = row[year_col]

        if pd.isna(year_cell):
            df.loc[index, year_col] = df.loc[index + 1, year_col]

    # Preenche com 0 as celulas 'NAN'
    df = df.fillna(0)

    # Novos DataFrames para separar os dados
    df_nacional = pd.DataFrame()
    df_estrangeiro = pd.DataFrame()

    for index, row in df.iterrows():
        if row['VISITANTES'] == 'NACIONAIS':
            df_nacional = df_nacional.append(df.loc[[index]])

        if row['VISITANTES'] == 'ESTRANGEIROS':
            df_estrangeiro = df_estrangeiro.append(df.loc[[index]])

    df_nacional = df_nacional.melt(
        id_vars=['ANO', 'VISITANTES'], var_name='MES', value_name='QUANTIDADE NACIONAL')

    df_estrangeiro = df_estrangeiro.melt(
        id_vars=['ANO', 'VISITANTES'], var_name='MES', value_name='QUANTIDADE ESTRANGEIRO')

    df_nacional = df_nacional.drop('VISITANTES', axis='columns')

    df_estrangeiro = df_estrangeiro.drop('VISITANTES', axis='columns')

    df = pd.merge(df_nacional, df_estrangeiro, how='left', on=['ANO', 'MES'])

    print(df.head())

    # Salva no excel a planilha
    df.to_excel('./data/treated_01.xlsx', sheet_name='Itaipu', index=False)


if __name__ == '__main__':
    main()
