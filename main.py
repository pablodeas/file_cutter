import shutil
import glob
import pandas as pd

"""
shutil.copy2(f'{ab}', f'{a}_ASE.xls')
shutil.copy2(f'{ab}', f'{a}_IQ.xls')
shutil.copy2(f'{ab}', f'{a}_SQL.xls')
shutil.copy2(f'{ab}', f'{a}_ORACLE.xls')
print("> Arquivos criados")
print('>')
"""

a = input("> Digite o nome da planilha: ")

for f in glob.glob(f'{a}.*'):
    print(f'> Encontrado arquivo: {f}')
    ab = f
    df = pd.read_excel(f)
    
def strip_values(series):
    return series.str.strip()

def create_new_sheet(bd_type, extention):
    try:

        df['tecnologia'] = strip_values(df['tecnologia'])

        if df['tecnologia'].eq(bd_type).any():
            df_filtered = df[df['tecnologia'] == bd_type]    
            df_filtered.to_excel(f'{a}_{bd_type}{extention}', index=False)
            print(f"> Arquivo criado: {a}_{bd_type}{extention}")
        else:
            print(f"> Não encontrado o valor '{bd_type}' na coluna 'tecnologia'")
    except Exception as e:
        print(f'> ERRO - {e}')

def main():
    try:
        match ab:
            case s if s.endswith('.ods'):
                uniq_tech = set(df['tecnologia'])
                for t in uniq_tech:
                    create_new_sheet(t, '.ods')

            case s if s.endswith('.xlsx'):
                uniq_tech = set(df['tecnologia'])
                for t in uniq_tech:
                    create_new_sheet(t, '.xlsx')

            case s if s.endswith('.csv'):
                uniq_tech = set(df['tecnologia'])
                for t in uniq_tech:
                    create_new_sheet(t, '.csv')

            case s if s.endswith('.xls'):
                uniq_tech = set(df['tecnologia'])
                for t in uniq_tech:
                    create_new_sheet(t, '.xls')

            case _:
                print('> Nenhum arquivo criado.')
                print('>')

    except shutil.ExecError as e:
        print(f'> ERRO - {e}')
        print('>')
    except Exception as e:
        print(f'> ERRO - {e}')
        print('>')

main()