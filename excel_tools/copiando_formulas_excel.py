import xlwings as xw
import pandas as pd

class TLDR_SOP:
    def __init__(self, path_download, path_raw):
        self.path_raw = path_raw
        self.path_download = path_download
        self.rac_production = 'production_data.xlsx'
        self.rac_raw = 'production_data_2.xlsx'

    def copy_production_information(self, production, raw):
        # Carregar os arquivos Excel
        df_production = pd.read_excel(self.path_download / production)
        df_raw = pd.read_excel(self.path_raw / raw)
        
        # Converter a coluna "ERP I/F Time" para o formato de data, se necessário
        for df in [df_production, df_raw]:
            if "ERP I/F Time" in df.columns:
                df["ERP I/F Time"] = pd.to_datetime(
                    df["ERP I/F Time"], 
                    origin="1986-12-01", 
                    unit="d", 
                    errors="coerce"
                )
        
        # Concatenar os dados
        df_combined = pd.concat([df_raw, df_production], ignore_index=True)
        
        return df_combined

    def append_data_to_excel(self, file_path, df_combined):
        # Abrir o arquivo Excel com xlwings
        with xw.App(visible=False) as app:
            wb = app.books.open(file_path)
            ws = wb.sheets[0]  # Ou o índice da planilha que você deseja

            # Encontrar a última linha preenchida
            last_row = ws.cells(ws.cells.last_cell.row, 1).end('up').row

            # Preencher a planilha com os dados a partir da próxima linha disponível
            ws.range(f'A{last_row + 1}').value = df_combined.values  # Preenche a partir da próxima linha disponível

            # Atualizar o formato da coluna de hora (por exemplo, coluna "ERP I/F Time")
            if "ERP I/F Time" in df_combined.columns:
                hour_column_index = df_combined.columns.get_loc("ERP I/F Time") + 1  # Ajusta o índice para o Excel
                # Definir o formato da célula para hora
                ws.range((last_row + 2, hour_column_index), (last_row + len(df_combined) + 1, hour_column_index)).number_format = 'h:mm:ss AM/PM'

            # Salvar o arquivo Excel atualizado
            wb.save()
            wb.close()

    def copy_production_orgs(self):
        # Combinar os dados de produção
        df_combined = self.copy_production_information(self.rac_production, self.rac_raw)
        return df_combined

if __name__ == "__main__":
    path_download = 'caminho/para/o/diretorio/baixados'  # Atualize o caminho
    path_raw = 'caminho/para/o/diretorio/raw'  # Atualize o caminho

    tldr = TLDR_SOP(path_download, path_raw)
    df_combined = tldr.copy_production_orgs()
    
    # Atualize o arquivo Excel com os dados adicionados
    file_path = 'caminho/para/o/arquivo/existente.xlsx'  # Caminho para o arquivo que você quer atualizar
    tldr.append_data_to_excel(file_path, df_combined)
