import pandas as pd

# Read Custom Report
df = pd.read_csv('Report_1779546464766.csv', low_memory=False)
print('Total de colunas:', len(df.columns))

# Find columns with Payment or Reference
payment_cols = [col for col in df.columns if 'Payment' in col or 'Reference' in col]
print('\nColunas com Payment ou Reference:')
for col in payment_cols:
    print(f'  - {col}')

# Search for UR19919 in Payment Reference columns
print('\nProcurando UR19919 em colunas de Payment Reference...')
found = False
for col in payment_cols:
    if 'Reference' in col:
        matches = df[df[col].astype(str).str.contains('UR19919', case=False, na=False)]
        if not matches.empty:
            print(f'\nEncontrado em coluna: {col}')
            print(f'Linhas: {[idx + 2 for idx in matches.index.tolist()]}')
            print(f'Valores:')
            print(matches[[col]].head())
            found = True

if not found:
    print('\nUR19919 NAO encontrado no Custom Report (Report_1779546464766.csv)')
    print('\nVerificando se existe em QUALQUER coluna...')
    for col in df.columns:
        matches = df[df[col].astype(str).str.contains('UR19919', case=False, na=False)]
        if not matches.empty:
            print(f'\nEncontrado em coluna: {col}')
            print(f'Linhas: {[idx + 2 for idx in matches.index.tolist()]}')
            break

# Made with Bob
