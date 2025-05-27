import pandas as pd
from datetime import datetime

# Lê o CSV com separador ';'
df = pd.read_csv('Relatório de Assistência LS Luziânia(Sheet1).csv', sep=';', dtype=str)

# Remove linhas sem data válida
df = df[df['Data da Reunião:'].notna()]

# Converte a data para datetime
df['Data da Reunião:'] = pd.to_datetime(df['Data da Reunião:'], dayfirst=True, errors='coerce')

# Filtra apenas datas a partir de setembro (mês >= 9)
df = df[df['Data da Reunião:'].dt.month >= 9]

# Calcula o total de surdos
df['SURDOS na reunião PRESENCIAL'] = pd.to_numeric(df['SURDOS na reunião PRESENCIAL:'], errors='coerce').fillna(0)
df['SURDOS na reunião pelo ZOOM'] = pd.to_numeric(df['SURDOS na reunião pelo ZOOM:'], errors='coerce').fillna(0)
df['total_surdos'] = df['SURDOS na reunião PRESENCIAL'] + df['SURDOS na reunião pelo ZOOM']

# Assistência (coluna "soma")
df['assistencia'] = pd.to_numeric(df['soma'], errors='coerce').fillna(0)

# Para datas iguais, manter o maior valor de assistência
df = df.sort_values(['Data da Reunião:', 'assistencia'], ascending=[True, False])
df = df.drop_duplicates(subset=['Data da Reunião:'], keep='first')

# Monta o DataFrame final
df_final = pd.DataFrame({
    'A': df['Data da Reunião:'].dt.strftime('%d/%m/%Y'),
    'B': df['total_surdos'].astype(int),
    'C': df['assistencia'].astype(int)
})

# Salva em Excel
df_final.to_excel('assistencia_filtrada.xlsx', index=False, header=False)
