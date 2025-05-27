import pandas as pd
from datetime import datetime

# Lê o CSV com separador ';'
df = pd.read_csv('Relatório de Assistência LS Luziânia(Sheet1).csv', sep=';', dtype=str)

# Remove linhas sem data válida
df = df[df['Data da Reunião:'].notna()]

# Converte a data para datetime
df['Data da Reunião:'] = pd.to_datetime(df['Data da Reunião:'], dayfirst=True, errors='coerce')

# Filtra apenas datas a partir de setembro de 2024 (inclusive)
df = df[df['Data da Reunião:'] >= pd.Timestamp(2024, 9, 1)]

# Filtra apenas reuniões de terça-feira (1), quinta-feira (3) ou domingo (6)
df = df[df['Data da Reunião:'].dt.weekday.isin([1, 3, 6])]

# Converte as colunas de assistência para numérico
df['SURDOS na reunião PRESENCIAL'] = pd.to_numeric(df['SURDOS na reunião PRESENCIAL:'], errors='coerce').fillna(0)
df['SURDOS na reunião pelo ZOOM'] = pd.to_numeric(df['SURDOS na reunião pelo ZOOM:'], errors='coerce').fillna(0)
df['OUVINTES na reunião PRESENCIAL'] = pd.to_numeric(df['OUVINTES na reunião PRESENCIAL:'], errors='coerce').fillna(0)
df['OUVINTES na reunião pelo ZOOM'] = pd.to_numeric(df['OUVINTES na reunião pelo ZOOM:'], errors='coerce').fillna(0)

# Para datas e grupos iguais, manter o maior valor de assistência
df = df.sort_values(['Data da Reunião:', 'Seu nome:', 'SURDOS na reunião PRESENCIAL', 'SURDOS na reunião pelo ZOOM',
                    'OUVINTES na reunião PRESENCIAL', 'OUVINTES na reunião pelo ZOOM'], ascending=[True, True, False, False, False, False])
df = df.drop_duplicates(subset=['Data da Reunião:', 'Seu nome:'], keep='first')

# Monta linhas para cada tipo de assistência
records = []
for _, row in df.iterrows():
    data = row['Data da Reunião:'].strftime('%d/%m/%Y')
    obs = row['Observação:'] if 'Observação:' in row and pd.notna(row['Observação:']) else ''
    records.append([data, 'Surdos na Assistencia Presencial', int(row['SURDOS na reunião PRESENCIAL']), '', obs])
    records.append([data, 'Surdos na Assistencia Zoom', int(row['SURDOS na reunião pelo ZOOM']), '', obs])
    records.append([data, 'Ouvintes na Assistencia Presencial', int(row['OUVINTES na reunião PRESENCIAL']), '', obs])
    records.append([data, 'Ouvintes na Assistencia Zoom', int(row['OUVINTES na reunião pelo ZOOM']), '', obs])

df_final = pd.DataFrame(records, columns=['A', 'B', 'C', 'D', 'E'])

# Salva em Excel
df_final.to_excel('assistencia_filtrada.xlsx', index=False, header=True)
