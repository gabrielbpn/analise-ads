import pandas as pd

def analisar_excel(arquivo):
    # Carrega o arquivo Excel em um DataFrame
    df = pd.read_excel(arquivo)

    # Exemplo de análise de células
    #celula_a1 = df.iloc[0, 0]  # Acessa a célula A1
    #celula_b2 = df.loc[1, 'B']  # Acessa a célula da segunda linha na coluna B
    print("URL:\n\n")

    coluna = df['Link']
    coluna = coluna.astype(str)

    if coluna.nunique() == 1:
        valor_igual = coluna.iloc[0]
        conclusao = f'Todas as URL são iguais. Essa aqui: {valor_igual}'
        print(conclusao)
    else:
        valores_diferentes = coluna.unique()
        conclusao = f'Opa! Encontramos algumas URLs diferentes:\n\n'
        i = 1
        for valor in valores_diferentes:
            conclusao += f'{i}: {valor}\n\n'
            i+=1
        print(conclusao)

    print("\n---------\nEvento de Otimização:\n\n")

    coluna = df['Optimized Event']
    coluna = coluna.astype(str)

    if coluna.nunique() == 1:
        valor_igual = coluna.iloc[0]
        conclusao = f'Todas os eventos são iguais. Essa aqui: {valor_igual}'
        print(conclusao)
    else:
        valores_diferentes = coluna.unique()
        conclusao = f'Opa! Encontramos alguns eventos diferentes:\n\n'
        i = 1
        for valor in valores_diferentes:
            conclusao += f'{i}: {valor}\n\n'
            i+=1
        print(conclusao)

    print("\n---------\nPaíses anunciando:\n\n")

    coluna = df['Countries']
    coluna = coluna.astype(str)

    if coluna.nunique() == 1:
        valor_igual = coluna.iloc[0]
        conclusao = f'Todas os países são iguais. Essa aqui: {valor_igual}'
        print(conclusao)
    else:
        valores_diferentes = coluna.unique()
        conclusao = f'Opa! Encontramos alguns países diferentes:\n\n'
        i = 1
        for valor in valores_diferentes:
            conclusao += f'{i}: {valor}\n\n'
            i+=1
        print(conclusao)
  
    celulas_vazias = df[df['Custom Audiences'].isnull()]

    if len(celulas_vazias) > 0:
        print("\n---------\nTem conjunto sem público!\n\n")

        for index, row in celulas_vazias.iterrows():
                linha = index + 1  # O índice das linhas começa em 0, somamos 1 para ter o número da linha real
                valor_celula_adjacente = df.at[index, 'Ad Set Name']
                print(f"Esse conjunto está sem público: {valor_celula_adjacente}")




nome_arquivo = 'exemplo.xlsx'
analisar_excel(nome_arquivo)
