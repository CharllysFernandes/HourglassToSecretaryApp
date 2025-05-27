import json
import pandas as pd

def load_json(filepath):
    with open(filepath, 'r', encoding='utf-8') as file:
        return json.load(file)

def get_publisher_field(publishers, user_id, field):
    return next((pub.get(field, '') for pub in publishers if pub.get('id') == user_id), '')

def enrich_reports_with_names(reports_df, publishers):
    reports_df['firstname'] = reports_df['user.id'].apply(lambda uid: get_publisher_field(publishers, uid, 'firstname'))
    reports_df['middlename'] = reports_df['user.id'].apply(lambda uid: get_publisher_field(publishers, uid, 'middlename'))
    reports_df['lastname'] = reports_df['user.id'].apply(lambda uid: get_publisher_field(publishers, uid, 'lastname'))
    return reports_df

def build_report_dataframe(reports_df):
    # Helper for full name
    def full_name(row):
        mid = f"{row['middlename']} " if pd.notna(row['middlename']) and row['middlename'] else ''
        return f"{row['firstname']} {mid}{row['lastname']}"
    # Helper for pioneer profile
    def pioneer_profile(x):
        if pd.isna(x) or x == "" or str(x).lower() == "nan":
            return 0
        if x == "Regular":
            return 2
        elif x == "Auxiliary":
            return 1
        else:
            return 0

    def pioneer_hours(row):
        x = row.get('pioneer')
        if pd.isna(x) or x == "" or str(x).lower() == "null":
            return 0
        return row.get('minutes_as_hours', 0)

    n = len(reports_df)
    return pd.DataFrame({
        'A': reports_df.apply(full_name, axis=1),
        'B': reports_df['year'],
        'C': reports_df['month'],
        'D': reports_df['minutes_as_hours'].apply(lambda x: 1 if pd.notna(x) and float(x) > 0.01 else 0),
        'E': reports_df.get('studies', pd.Series([None]*n)),
        'F': [''] * n,  # Estudo Biblico Extra (surdos) - vazio
        'G': reports_df.get('pioneer', pd.Series([None]*n)).apply(pioneer_profile),
        'H': reports_df.apply(pioneer_hours, axis=1),
        'I': reports_df['credithours'],
        'J': [''] * n,  # Opcional, hora de escolas aprovadas
        'K': reports_df.get('remarks', pd.Series(['']*n)),
        'L': [0] * n,   # Sempre 0
    })

def main():
    data = load_json('hourglass-export.json')
    publishers = data.get('publishers', [])
    reports = data.get('reports', [])
    reports_df = pd.json_normalize(reports)
    reports_df = enrich_reports_with_names(reports_df, publishers)
    # Filtra apenas anos 2023, 2024 e 2025
    reports_df = reports_df[reports_df['year'].isin([2023, 2024, 2025])]
    df = build_report_dataframe(reports_df)
    df.to_excel('reports_data.xlsx', index=False, header=False)

if __name__ == "__main__":
    main()
