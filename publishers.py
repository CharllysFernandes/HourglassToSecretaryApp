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
baptisms = [format_date(publisher.get('baptism', '')) for publisher in data['publishers']]
sexes = [1 if publisher.get('sex') == "Male" else 0 for publisher in data['publishers']]
cellphones = [extract_numbers(publisher.get('cellphone', '')) for publisher in data['publishers']]
loginemails = [publisher.get('loginemail', '') for publisher in data['publishers']]

# Extração dos endereços
addresses = {address['id']: address for address in data['addresses']}
address_lines = []

for publisher in data['publishers']:
    address_id = publisher.get('address_id')
    if address_id and address_id in addresses:
        address = addresses[address_id]
        line1 = address.get('line1', '')
        line2 = address.get('line2', '')
        address_lines.append(f"{line1} - {line2}")
    else:
        address_lines.append('')

# Criar um DataFrame com os dados extraídos
df = pd.DataFrame({
    'A': [''] * len(firstnames),
    'B': firstnames,
    'C': middlenames,
    'D': lastnames,
    'E': [''] * len(firstnames),
    'F': statuses,
    'G': appt_ms,
    'H': appt_elder,
    'I': [''] * len(firstnames),
    'J': [''] * len(firstnames),
    'K': [''] * len(firstnames),
    'L': births,
    'M': baptisms,
    'N': sexes,
    'O': cellphones,
    'P': [''] * len(firstnames),
    'Q': [''] * len(firstnames),
    'R': loginemails,
    'S': [''] * len(firstnames),
    'T': [''] * len(firstnames),
    'U': address_lines,
    'V': ['Luziânia-GO'] * len(firstnames)
})

# Escrever para um arquivo Excel
df.to_excel('publishers_data.xlsx', index=False, header=False, startcol=0)
