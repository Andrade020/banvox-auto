import os                                                   # Verba ligata
import json                                                 # O flumen numerorum
import pandas as pd                                         # O magna terra
from base_excel_processador import BaseExcelProcessor
###########################################################################
class MapaCotasProcessor(BaseExcelProcessor):
    '''
    clsse respns por procss arquivos excel contnd maps de cots.
    herda funcionalidades da clsse basxcl.
    funcionalidades:
    - ᕦ(ò_óˇ)ᕤ le arqvs excl sem cabclh.
    - ᕦ(ò_óˇ)ᕤ  identifica o cabclh global (primeira linha do arqvo).
    - ᕦ(ò_óˇ)ᕤ  separa difrnt cartrs dentro do arqvo, basnd-se em linhs de separacao.
    - ᕦ(ò_óˇ)ᕤ  convrt os dads procss em json e salva o resultado na pasta de sada.
    '''
    def process_file(self, file):
        filename = os.path.basename(file)
        try:
            df = pd.read_excel(file, header=None)
        except Exception as e:
            raise Exception(f"Erro ao ler {filename}: {e}")
        if len(df) < 3:
            raise Exception(f"{filename} não possui a estrutura esperada.")
        
        # Linha 0: cabeçalho global de todas as tabelinhas
        global_header = df.iloc[0].tolist()

        portfolios = []
        current_portfolio_name = None
        current_data = []
        # Linha 1: nome da primeira carteira
        if pd.notna(df.iloc[1, 0]):
            current_portfolio_name = df.iloc[1, 0]
        else:
            raise Exception(f"{filename}: A linha 2 (índice 1) não contém o nome da carteira.")

        # Proc as linhas dps da lin 2
        for i in range(2, len(df)):
            row = df.iloc[i]
            # Linh de seprç: somente a primeira coluna tm valr e as dms tão vazia
            if pd.notna(row[0]) and row[1:].isnull().all():
                if current_data:
                    portfolios.append({
                        "carteira": current_portfolio_name,
                        "cabecalho": global_header,
                        "linhas": current_data
                    })
                current_portfolio_name = row[0]
                current_data = []
            else:
                current_data.append(row.tolist())
        if current_portfolio_name is not None and current_data:
            portfolios.append({
                "carteira": current_portfolio_name,
                "cabecalho": global_header,
                "linhas": current_data
            })

        data_json = {
            "data": filename.replace(".xlsx", ""),
            "mapa_cotas": portfolios
        }
        output_filename = filename.replace(".xlsx", ".json")
        output_path = os.path.join(self.output_folder, output_filename)
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(data_json, f, ensure_ascii=False, indent=2, default=str)

if __name__ == "__main__":
    processor = MapaCotasProcessor(
        input_folder=r"raw_data\mapa_cotas",
        output_folder=r"downloads\banvox\mapa_cotas"
    )
    processor.process_all_files()
