import pandas as pd
import traceback

# ler arquivo Excel
df = pd.read_excel('G:/Drives compartilhados/Qualidade/NOTIFICAÇÕES E ANÁLISE DE EVENTOS/NOTIFICAÇÕES 2026/03 - Março/terminado_Notificações CMA - Março(1).xlsx')

# criar colunas, se quiser
df['Qual incidente ocorreu'] = None
df['Titulo do Incidente'] = None
df['Descrição 2'] = None

for i in range(len(df)):
    if pd.isna(df.loc[i, 'classificacao']):
        continue
    try:
        texto_original = str(df.loc[i, 'classificacao'])

        corte1 = texto_original.find('[')
        corte2 = texto_original.find(']')

        # valida se encontrou [ e ]
        if corte1 == -1 or corte2 == -1 or corte2 <= corte1:
            print(f'Linha {i}: formato inválido -> {texto_original}')
            continue

        texto = texto_original[corte1+1:corte2]
        lista = texto.split(';', 2)

        if len(lista) < 3:
            print(f'Linha {i}: menos de 3 partes -> {lista}')
            continue

        df.loc[i, 'Qual incidente ocorreu'] = lista[0].replace('"', '').strip().replace('Classificação: ', '')
        df.loc[i, 'Titulo do Incidente'] = lista[1].replace('"', '').strip()
        df.loc[i, 'Descrição 2'] = lista[2].replace('"', '').strip()

    except:
        print(i)
        print(traceback.format_exc())
        break

# salvar resultado em outro arquivo Excel
df.to_excel('G:/Drives compartilhados/Qualidade/NOTIFICAÇÕES E ANÁLISE DE EVENTOS/NOTIFICAÇÕES 2026/03 - Março/saida.xlsx', index=False)
print('terminou')
print(df.head())