import json
import pandas as pd
from datetime import datetime, timedelta

# Função para encontrar a primeira quinta-feira de um mês
def first_thursday(year_month):
    year, month = map(int, year_month.split('-'))
    first_day = datetime(year, month, 1)
    first_thursday = first_day + timedelta(days=(3 - first_day.weekday() + 7) % 7)
    return first_thursday

# Função para encontrar o próximo domingo dado uma data
def next_sunday(date):
    return date + timedelta(days=(6 - date.weekday()))

# Carregar dados do arquivo JSON
with open('hourglass-export.json', 'r') as file:
    data = json.load(file)

# Função para criar a lista de datas e valores para um mês
def create_attendance_list(month_data):
    year_month = month_data['month']
    first_thursday_date = first_thursday(year_month)
    dates_and_values = []

    for i in range(4):
        thursday_date = first_thursday_date + timedelta(weeks=i)
        sunday_date = next_sunday(thursday_date)
        thursday_value = month_data.get(f'mw{i+1}', 0)
        sunday_value = month_data.get(f'we{i+1}', 0)
        
        dates_and_values.append((thursday_date, thursday_value))
        dates_and_values.append((sunday_date, sunday_value))

    return dates_and_values

# Extração dos dados necessários e criação do DataFrame
attendance_data = []
for entry in data['attendance']['attendance']:
    attendance_data.extend(create_attendance_list(entry))

# Formatar as datas e valores no DataFrame
df = pd.DataFrame(attendance_data, columns=['Date', 'Value'])
df['Date'] = df['Date'].apply(lambda x: x.strftime('%d/%m/%Y'))
df['B'] = [''] * len(df)  # Adiciona uma coluna B vazia

# Reorganizar as colunas para corresponder à especificação
df = df[['Date', 'B', 'Value']]

# Escrever para um arquivo Excel
df.to_excel('attendance_data.xlsx', index=False, header=False)
