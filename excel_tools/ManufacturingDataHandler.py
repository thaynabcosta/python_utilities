from pathlib import Path
import pandas as pd
import xlwings as xw
import datetime as dt

class ManufacturingDataHandler:
    def __init__(self, path_download, path_raw, path_db):
        self.path_download = path_download
        self.path_raw = path_raw
        self.path_db = path_db

    def excel_download_to_df(self, file):
        print(f'Transformando {file} em dataframe')
        df = pd.read_excel(file)
        return df
    
    def excel_raw_to_df(self, file, sheet_name):
        print(f'Transformando {file} em dataframe')
        df = pd.read_excel(file, sheet_name=sheet_name)
        return df
    
    def load_db(self):
        print('Carregando banco de dados...')
        self.db = pd.read_csv(self.path_db)
    
    def categorize_time(self, row, column):
        if pd.isna(row[column]):
            return 'Indefinido'
        time = row[column].time()
        if (time >= pd.to_datetime('06:00:00').time()) and (time <= pd.to_datetime('15:49:00').time()):
            return 'Dia'
        else:
            return 'Noite'
    
    def filling_out_production_formulas(self, df):
        print("Preenchendo colunas vazias...")
        df['ERP I/F Time'] = pd.to_datetime(df['ERP I/F Time']) 
        df['Year'] = df['ERP I/F Time'].dt.year
        df['Turno'] = df.apply(lambda row: self.categorize_time(row, 'ERP I/F Time'), axis=1)
        df['Month'] = df['ERP I/F Time'].dt.month
        df['Day'] = df['ERP I/F Time'].dt.day
        df['Gambi'] = df['all'] + df['all']
        df['all'] = df['Gambi'].map(self.db.set_index('T')['U'].to_dict())
        df['Conc'] = df['Tool'] + df['Type']
        df['Tool'] = df['R'].map(self.db.set_index('A')['B'].to_dict())
        df['Type'] = df['R'].map(self.db.set_index('A')['C'].to_dict())
        df['Week'] = df['ERP I/F Time'].dt.strftime('%U')
        df['PROD'] = df['Scan QTY']

        return df
    
    def filling_out_defect_formulas(self, df):
        print("Preenchendo colunas vazias...")
        df['Defect'] = df['Initial Defect Symptom LV1'] + " " + df['Initial Defect Symptom LV2']
        df['Defect 2'] = df['Initial Defect Symptom LV1'] + " " + df['Department']
        df['Turno'] = df.apply(lambda row: self.categorize_time(row, column='Repair Date'), axis=1)
        df['Year'] = df['Repair Date'].dt.year
        df['Month'] = df['Repair Date'].dt.month
        df['Day'] = df['Repair Date'].dt.day
        df['Week'] = df['Repair Date'].dt.strftime('%U')
        df['Department'] = df['BG'].map(self.db.set_index('P')['Q'].to_dict())
        df['Department'] = df.apply(lambda row: 'Supplier Leakage' if row['Department'] == 'Supplier' and row['AI'] == 'Leakage' else
                                              'Supplier Assembly' if row['Department'] == 'Supplier' and row['AI'] != 'Leakage' else
                                              'MFG Leakage' if row['Department'] == 'MFG' and row['AI'] == 'Leakage' else
                                              'MFG Assembly' if row['Department'] == 'MFG' and row['AI'] != 'Leakage' else
                                              row['Department'], axis=1)
        df['Colunas3'] = df['Colunas2'] + df['Colunas2']
        df['Colunas2'] = df['L'].map(self.db.set_index('T')['U'].to_dict())
        df['Colunas1'] = df['Tool'] + df['Type']
        df['Tool'] = df['AB'].map(self.db.set_index('A')['B'].to_dict())
        df['Type'] = df['AB'].map(self.db.set_index('A')['C'].to_dict())

        return df

    def remove_duplicates_production(self, df):
        print("Removendo duplicidades da planilha de produção...")
        # Filtra as linhas que não atendem às condições de duplicidade
        df = df[~((df['ERP I/F Time'].dt.year == 2024) & (df['ERP I/F Time'].dt.month == 12))]
        return df

    def remove_duplicates_defect(self, df):
        print("Removendo duplicidades da planilha de defeito...")
        # Filtra as linhas que não atendem às condições de duplicidade
        df = df[~((df['Repair Date'].dt.year == 2024) & (df['Repair Date'].dt.month == 10) & (df['Repair Date'].dt.day == 12))]
        return df

    def df_treatment_defect(self, df_raw, df_production):
        print("Tratando os dados nos df")
        df_update = pd.concat([df_raw, df_production], ignore_index=True) 
        df = self.filling_out_defect_formulas(df_update)
        df = self.remove_duplicates_defect(df)  # Remove duplicidades
        return df
    
    def df_treatment(self, df_raw, df_production):
        print("Tratando os dados nos df")
        df_update = pd.concat([df_raw, df_production], ignore_index=True) 
        df = self.filling_out_production_formulas(df_update)
        df = self.remove_duplicates_production(df)  # Remove duplicidades
        df = df[(df['ERP I/F Time'].dt.year == 2024) & (df['ERP I/F Time'].dt.month == 12)]
        return df
    
    def clear_sheet(self, file, sheet_name):
        print(f"Limpando a planilha {sheet_name} no arquivo {file}...")
        with xw.App(visible=False) as app:
            wb = app.books.open(file)
            ws = wb.sheets[sheet_name]
            ws.range('A2').expand('table').clear_contents()  # Limpa todo o conteúdo da planilha, exceto o cabeçalho
            wb.save(file)
            wb.close()
    
    def insert_production_information_into_raw(self, sheet_name, file, df):
        print(f'Atualizando planilha {sheet_name} no arquivo {file}...')
        with xw.App(visible=False) as app:
            wb = app.books.open(file)
            ws = wb.sheets[sheet_name]
            ws.range('A2').value = df.values  # Insere os dados a partir da segunda linha
            print("Salvando...")
            wb.save(file)
            wb.close()
    
    def update_raw_data_production(self, file_raw, file_production):
        print(f"Atualizando o arquivo {file_raw}...")
        df_raw = self.excel_raw_to_df(file_raw, sheet_name='AZ_Production Raw Data')
        df_production = self.excel_download_to_df(file_production)
        df_raw = self.df_treatment(df_raw, df_production)
        self.clear_sheet(file_raw, sheet_name='AZ_Production Raw Data')  # Limpa a planilha antes de inserir os dados
        self.insert_production_information_into_raw(sheet_name='AZ_Production Raw Data', file=file_raw, df=df_raw)

    def update_raw_data_defect(self, file_raw, file_defect):
        print(f"Atualizando o arquivo {file_raw}...")
        df_raw = self.excel_raw_to_df(file_raw, sheet_name='AZ_Defect Raw Data')
        df_defect = self.excel_download_to_df(file_defect)
        df_raw = self.df_treatment_defect(df_raw, df_defect)
        self.clear_sheet(file_raw, sheet_name='AZ_Defect Raw Data')  # Limpa a planilha antes de inserir os dados
        self.insert_production_information_into_raw(sheet_name='AZ_Defect Raw Data', file=file_raw, df=df_raw)
