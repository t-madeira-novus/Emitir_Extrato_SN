import pandas as pd

df = pd.read_excel('prestados.xlsx')

df_new = pd.DataFrame(columns=['Código', 'Nome'])

for i in df.index:
    aux = str(df.at[i, "NOME"]).split("-")
    print(aux)
    df_new.at[i, 'Código'] = aux[0]
    df_new.at[i, 'Nome'] = aux[1]

df_new.to_excel('empresas_emitir_extrato_sn.xlsx', index=False)