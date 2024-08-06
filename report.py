import json
import pandas as pd

# Carregar dados do arquivo JSON
with open('dados.json', 'r') as file:
    data = json.load(file)

# Obter publishers e reports
publishers = data.get('publishers', [])
reports = data.get('reports', [])

# Converter reports para um DataFrame
reports_df = pd.json_normalize(reports)

# Adicionar informações dos publishers ao DataFrame de reports
reports_df['firstname'] = reports_df['user.id'].apply(
    lambda user_id: next((pub['firstname'] for pub in publishers if pub['id'] == user_id), '')
)
reports_df['middlename'] = reports_df['user.id'].apply(
    lambda user_id: next((pub['middlename'] for pub in publishers if pub['id'] == user_id), '')
)
reports_df['lastname'] = reports_df['user.id'].apply(
    lambda user_id: next((pub['lastname'] for pub in publishers if pub['id'] == user_id), '')
)

# Criar a nova DataFrame com as colunas especificadas
df = pd.DataFrame({
    'A': reports_df.apply(lambda row: f"{row['firstname']} {row['middlename'] + ' ' if pd.notna(row['middlename']) else ''}{row['lastname']}", axis=1),
    'B': reports_df['year'],
    'C': reports_df['month'],
    'D': reports_df['placements'].fillna(0),
    'E': reports_df['videoshowings'].fillna(0),
    'F': reports_df['minutes'] / 60,
    'G': [''] * len(reports_df),
    'H': reports_df['credithours'].fillna(0),
    'I': reports_df['returnvisits'].fillna(0),
    'J': reports_df['studies'].fillna(0),
    'K': reports_df.apply(
        lambda row: 2 if row['pioneer'] == 'Regular' else (1 if row['pioneer'] == 'Auxiliary' else (0 if row['minutes'] >= 0 else '')),
        axis=1
    ),
    'M': reports_df['remarks'].fillna('')
})

# Escrever para um arquivo Excel
df.to_excel('reports_data.xlsx', index=False, header=False)
