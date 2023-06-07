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

    coluna = df['Campaign Objective']
    coluna = coluna.astype(str)

    print("\n---------\nObjetivo da campanha:\n\n")

    if coluna.nunique() == 1:
        valor_igual = coluna.iloc[0]
        if valor_igual == 'Conversions':
            valor_igual = 'Conversão'
        conclusao = f'Objetivos das campanhas são esses: {valor_igual}'
        print(conclusao)
    else:
        valores_diferentes = coluna.unique()
        conclusao = f'Opa! Encontramos objetivos diferentes nas campanhas:\n\n'
        i = 1
        for valor in valores_diferentes:
            conclusao += f'{i}: {valor}\n\n'
            i+=1
        print(conclusao)


    coluna = df['Attribution Spec']
    coluna = coluna.astype(str)

    print("\n---------\nJanela de Atribuição:\n\n")

    if coluna.nunique() == 1:
        valor_igual = coluna.iloc[0]
        conclusao = f'Janela de atribuição está igual em todas: {valor_igual}'
        print(conclusao)
    else:
        valores_diferentes = coluna.unique()
        conclusao = f'Opa! Encontramos janelas de atribuição diferentes nas campanhas:\n\n'
        i = 1
        for valor in valores_diferentes:
            if valor == '[{"event_type":"CLICK_THROUGH","window_days":7}]':
                valor = 'Clique de 7 dias'

            elif valor == '[{"event_type":"CLICK_THROUGH","window_days":7},{"event_type":"VIEW_THROUGH","window_days":1}]':
                valor = "Clique de 7 dias e atribuição de 1 dia"

            conclusao += f'{i}: {valor}\n\n'
            i+=1
        print(conclusao)

    coluna = df['Title']
    coluna = coluna.astype(str)

    print("\n---------\nTítulo:\n\n")

    if coluna.nunique() == 1:
        valor_igual = coluna.iloc[0]
        conclusao = f'Títulos iguais: {valor_igual}'
        print(conclusao)
    else:
        valores_diferentes = coluna.unique()
        conclusao = f'Opa! Encontramos títulos diferentes nas campanhas:\n\n'
        i = 1
        for valor in valores_diferentes:
            conclusao += f'{i}: {valor}\n\n'
            i+=1
        print(conclusao)

    
    coluna = df['Body']
    coluna = coluna.astype(str)

    print("\n---------\nTexto (corpo) do anúncio:\n\n")

    if coluna.nunique() == 1:
        valor_igual = coluna.iloc[0]
        conclusao = f'Textos iguais: {valor_igual}'
        print(conclusao)
    else:
        valores_diferentes = coluna.unique()
        conclusao = f'Opa! Encontramos textos diferentes nas campanhas:\n\n'
        i = 1
        for valor in valores_diferentes:
            conclusao += f'{i}: {valor}\n\n'
            i+=1
        print(conclusao)

    coluna = df['Call to Action']
    coluna = coluna.astype(str)

    print("\n---------\nCTA do anúncio:\n\n")

    if coluna.nunique() == 1:
        valor_igual = coluna.iloc[0]
        if valor_igual == 'LEARN_MORE':
            valor_igual = 'Saiba Mais'
        conclusao = f'CTA iguais: {valor_igual}'
        print(conclusao)
    else:
        valores_diferentes = coluna.unique()
        conclusao = f'Opa! Encontramos CTAs diferentes nas campanhas:\n\n'
        i = 1
        for valor in valores_diferentes:
            if valor == 'LEAN_MORE':
                valor == 'Saiba Mais'
            conclusao += f'{i}: {valor}\n\n'
            i+=1
        print(conclusao)

    
    coluna = df['Instagram Account ID']
    coluna = coluna.astype(str)

    print("\n---------\nPerfil do Instagram:\n\n")

    if coluna.nunique() == 1:
        valor_igual = coluna.iloc[0]
        conclusao = f'Perfis do IG iguais: {valor_igual}'
        print(conclusao)
    else:
        valores_diferentes = coluna.unique()
        conclusao = f'Opa! Encontramos Perfis diferentes nos anúncios. Vá procurar porque não tem posso te dar o nome deles.\n\n'
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
