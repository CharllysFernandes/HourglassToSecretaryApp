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
with open('hourglass-export.json', 'r', encoding='utf-8') as file:
    data = json.load(file)

# Mapeia id do grupo para nome do grupo
group_id_to_name = {group['id']: group['name'] for group in data.get('fsGroups', [])}

# Extração dos dados necessários
firstnames = [publisher['firstname'] for publisher in data['publishers']]
middlenames = [publisher.get('middlename', '') for publisher in data['publishers']]
lastnames = [publisher['lastname'] for publisher in data['publishers']]
# Extração de tipo de pineiro 
# 0 não é pioneiro
# 1 é auxiliar continuo
# 2 é regular
# 3 é especial
statuses = [2 if publisher.get('status') == "Regular Pioneer" else (1 if publisher.get('status') == "Continuous Auxiliary Pioneer" else '') for publisher in data['publishers']]
# Extração Servo ministerial
# 0 não é servo ministerial
# 1 é servo ministerial
appt_ms = [1 if publisher.get('appt') == "MS" else 0 for publisher in data['publishers']]
# Extração de ancião
# 0 não é ancião
# 1 é ancião
appt_elder = [1 if publisher.get('appt') == "Elder" else 0 for publisher in data['publishers']]

# Adiciona a coluna E: nome do grupo de cada publisher
group_names = [
    group_id_to_name.get(publisher.get('group_id'), '')
    for publisher in data['publishers']
]

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
    'A': [''] * len(firstnames), #opcional, pode ser deixado vazio
    'B': firstnames, #primeiro nome
    'C': middlenames, # nome do meio
    'D': lastnames, # sobrenome
    'E': group_names, #nome do grupo de serviço de campo
    'F': statuses, # Tipo de Pioneiro (auxiliar, auxiliar continuo, Regular, especial)
    'G': appt_ms, # Servo Ministerial
    'H': appt_elder, # Ancião
    'I': [''] * len(firstnames), # opcional, pode ser deixado vazio, status de ungindo
    'J': [''] * len(firstnames), # opcional, pode ser deixado vazio, status de removido
    'K': [''] * len(firstnames), # opcional, pode ser deixado vazio, status de desativado
    'L': births, # datas de nascimento
    'M': baptisms, # datas de batismo
    'N': sexes, # sexo (1 masculino, 0 feminino)
    'O': cellphones, # telefones celulares
    'P': [''] * len(firstnames), # opcional, pode ser deixado vazio, contato casa 
    'Q': [''] * len(firstnames), # opcional, pode ser deixado vazio, contato trabalho
    'R': loginemails, # emails de contato
    'S': [''] * len(firstnames), # opcional, pode ser deixado vazio, contato de emergência nome
    'T': [''] * len(firstnames), # opcional, pode ser deixado vazio, contato de emergência telefone
    'U': address_lines, # endereços rua
    'V': ['Luziânia-GO'] * len(firstnames), # cidade
    'W': [''] * len(firstnames), # opcional, pode ser deixado vazio, observações
})

# Garante que todos os campos de texto sejam string para evitar problemas de pontuação/acentuação
for col in ['B', 'C', 'D', 'E', 'O', 'R', 'U', 'V', 'W']:
    df[col] = df[col].astype(str)

# Escrever para um arquivo Excel (.xlsx) obrigatoriamente, usando openpyxl
df.to_excel('publishers_data.xlsx', index=False, header=False, startcol=0, engine='openpyxl')
