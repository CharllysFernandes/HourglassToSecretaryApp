import json
import pandas as pd
from datetime import datetime
import re

# Função para formatar datas no formato "DD/MM/YYYY"
def format_date(date_str):
    try:
        return datetime.strptime(date_str, "%Y-%m-%d").strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return ''

# Função para extrair apenas números de uma string
def extract_numbers(phone_str):
    if phone_str is None:
        return ''
    return re.sub(r'\D', '', phone_str)

# Carregar dados do arquivo JSON
with open('dados.json', 'r') as file:
    data = json.load(file)

# Extração dos dados necessários
firstnames = [publisher['firstname'] for publisher in data['publishers']]
middlenames = [publisher.get('middlename', '') for publisher in data['publishers']]
lastnames = [publisher['lastname'] for publisher in data['publishers']]
statuses = [2 if publisher.get('status') == "Regular Pioneer" else (1 if publisher.get('status') == "Continuous Auxiliary Pioneer" else '') for publisher in data['publishers']]
appt_ms = [1 if publisher.get('appt') == "MS" else 0 for publisher in data['publishers']]
appt_elder = [1 if publisher.get('appt') == "Elder" else 0 for publisher in data['publishers']]
births = [format_date(publisher.get('birth', '')) for publisher in data['publishers']]
baptisms = [publisher.get('baptism', '') for publisher in data['publishers']]
sexes = [1 if publisher.get('sex') == "Male" else 0 for publisher in data['publishers']]
cellphones = [extract_numbers(publisher.get('cellphone', '')) for publisher in data['publishers']]
loginemails = [publisher.get('loginemail', '') for publisher in data['publishers']]

# Criar um DataFrame com os dados extraídos
df = pd.DataFrame({
    'B': firstnames,
    'C': middlenames,
    'D': lastnames,
    'F': statuses,
    'G': appt_ms,
    'H': appt_elder,
    'L': births,
    'M': baptisms,
    'N': sexes,
    'O': cellphones,
    'R': loginemails
})

# Escrever para um arquivo Excel
df.to_excel('publishers_data.xlsx', index=False, header=False, startcol=1)
