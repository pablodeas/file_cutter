import shutil
import glob
import pandas as pd

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
                create_new_sheet('ASE', '.ods')
                create_new_sheet('IQ', '.ods')
                create_new_sheet('SQL', '.ods')
                create_new_sheet('ORACLE', '.ods')
                create_new_sheet('GCP-MYSQL', '.ods')

            case s if s.endswith('.xlsx'):
                create_new_sheet('ASE', '.xlsx')
                create_new_sheet('IQ', '.xlsx')
                create_new_sheet('SQL', '.xlsx')
                create_new_sheet('ORACLE', '.xlsx')
                create_new_sheet('GCP-MYSQL', '.xlsx')

            case s if s.endswith('.csv'):
                create_new_sheet('ASE', '.csv')
                create_new_sheet('IQ', '.csv')
                create_new_sheet('SQL', '.csv')
                create_new_sheet('ORACLE', '.csv')
                create_new_sheet('GCP-MYSQL', '.csv')

            case s if s.endswith('.xls'):
                create_new_sheet('ASE', 'xls')
                create_new_sheet('IQ', 'xls')
                create_new_sheet('SQL', 'xls')
                create_new_sheet('ORACLE', 'xls')
                create_new_sheet('GCP-MYSQL', 'xls')

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