import pandas as pd
import numpy as np

df = pd.read_excel('LOGCOMEX-BIDROUTE.xlsx')
def formatar_numero(numero):
    if not pd.isna(numero):
        numero_inteiro = int(numero)
        if numero_inteiro > 999:
            numero_formatado = f'{numero_inteiro // 1000}.{numero_inteiro % 1000:03d}'
            return numero_formatado
        else:
            return str(numero)
    else:
        return str(numero)

df['kgs_teu'] = df['kgs_teu'].apply(formatar_numero)
df['volumes'] = df['volumes'].apply(formatar_numero)

df['player'] = df['player'].astype(str)
df['kgs_teu'] = df['kgs_teu'].astype(str)
df['volumes'] = df['volumes'].astype(str)
df['opportunity_type'] = df['opportunity_type'].astype(str)
df['logcomex'] = df['logcomex'].astype(str)

grouped = df.groupby(['mode', 'profile_id', 'destination_id', 'origin_id']).agg({
    'id': 'first',
    'deleted_at': 'first',
    'created_at': 'first',
    'updated_at': 'first',
    'player': ', '.join,
    'kgs_teu': ', '.join,
    'volumes': ', '.join,
    'opportunity_type': ', '.join,
    'logcomex': 'first',
    'created_by_id': 'first',
    'updated_by_id': 'first'
}).reset_index()

def adicionar_chaves(texto):
    return f'{{{texto}}}'

grouped['player'] = grouped['player'].apply(adicionar_chaves)
grouped['kgs_teu'] = grouped['kgs_teu'].apply(adicionar_chaves)
grouped['volumes'] = grouped['volumes'].apply(adicionar_chaves)

output_filename = 'output.xlsx'
grouped.to_excel(output_filename, index=False)
print(f'O DataFrame foi salvo em {output_filename}')
