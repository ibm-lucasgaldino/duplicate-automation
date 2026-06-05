import pandas as pd
import sys

# Set UTF-8 encoding for output
sys.stdout.reconfigure(encoding='utf-8')

print("=" * 60)
print("VERIFICANDO UR19919 EM TODOS OS ARQUIVOS")
print("=" * 60)

# 1. Custom Report CSV
print("\n1. CUSTOM REPORT (Report_1779546464766.csv)")
print("-" * 60)
try:
    df_custom = pd.read_csv('Report_1779546464766.csv', low_memory=False)
    print(f"Total de linhas: {len(df_custom)}")
    print(f"Total de colunas: {len(df_custom.columns)}")
    
    # Procurar coluna Payment - Payment Reference
    payment_ref_cols = [col for col in df_custom.columns if 'Payment' in col and 'Reference' in col]
    print(f"\nColunas com 'Payment' e 'Reference': {payment_ref_cols}")
    
    # Procurar UR19919 em todas as colunas
    found_custom = False
    for col in df_custom.columns:
        matches = df_custom[df_custom[col].astype(str).str.contains('UR19919', case=False, na=False)]
        if not matches.empty:
            print(f"\n[OK] ENCONTRADO na coluna: {col}")
            print(f"  Linhas: {[idx + 2 for idx in matches.index.tolist()]}")
            found_custom = True
    
    if not found_custom:
        print("\n[X] UR19919 NAO encontrado no Custom Report")
except Exception as e:
    print(f"Erro ao ler Custom Report: {e}")

# 2. DR Report Excel
print("\n\n2. DR REPORT (oppData_052626.xlsx)")
print("-" * 60)
try:
    df_dr = pd.read_excel('oppData_052626.xlsx')
    print(f"Total de linhas: {len(df_dr)}")
    print(f"Total de colunas: {len(df_dr.columns)}")
    
    # Verificar coluna AL (indice 37)
    if len(df_dr.columns) > 37:
        col_al_name = df_dr.columns[37]
        print(f"\nColuna AL (indice 37): {col_al_name}")
        
        matches = df_dr[df_dr.iloc[:, 37].astype(str).str.contains('UR19919', case=False, na=False)]
        if not matches.empty:
            print(f"\n[OK] ENCONTRADO na coluna AL")
            print(f"  Linhas: {[idx + 2 for idx in matches.index.tolist()]}")
            print(f"\n  Valores encontrados:")
            for idx in matches.index:
                print(f"    Linha {idx + 2}: {df_dr.iloc[idx, 37]}")
        else:
            print("\n[X] UR19919 NAO encontrado na coluna AL")
    else:
        print(f"\n[X] Coluna AL nao existe (total de colunas: {len(df_dr.columns)})")
except Exception as e:
    print(f"Erro ao ler DR Report: {e}")

# 3. Arquivo Excel alternativo
print("\n\n3. ARQUIVO ALTERNATIVO (Duplicate payment check Saas SL CASH Mar26.xlsx)")
print("-" * 60)
try:
    df_alt = pd.read_excel('Duplicate payment check Saas SL CASH Mar26.xlsx')
    print(f"Total de linhas: {len(df_alt)}")
    print(f"Total de colunas: {len(df_alt.columns)}")
    
    matches = df_alt[df_alt.apply(lambda row: row.astype(str).str.contains('UR19919', case=False, na=False).any(), axis=1)]
    if not matches.empty:
        print(f"\n[OK] ENCONTRADO")
        print(f"  Linhas: {[idx + 2 for idx in matches.index.tolist()]}")
        for col in df_alt.columns:
            col_matches = df_alt[df_alt[col].astype(str).str.contains('UR19919', case=False, na=False)]
            if not col_matches.empty:
                print(f"  Coluna: {col}")
    else:
        print("\n[X] UR19919 NAO encontrado")
except Exception as e:
    print(f"Erro ao ler arquivo alternativo: {e}")

print("\n" + "=" * 60)
print("FIM DA VERIFICACAO")
print("=" * 60)

# Made with Bob
