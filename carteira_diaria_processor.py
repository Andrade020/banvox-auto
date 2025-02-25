import os
import json
import pandas as pd
from base_excel_processador import BaseExcelProcessor

class CarteiraDiariaProcessor(BaseExcelProcessor):
    def process_sheet(self, df, sheet_name):
        # Obtém o nome da carteira (na cél A3, índ 2,0)
        carteira_name = df.iloc[2, 0]
        tabelas = []
        i = 3  # Inicia após a linha 4 (índ 3)
        while i < len(df):
            # Se a linha atual estiver completamente vazia (NaN), pula
            if df.iloc[i].isnull().all():
                i += 1
                continue
            
            # 1) Lê o nome da tabela
            tabela_nome = df.iloc[i, 0]
            i += 1
            if i >= len(df):
                break

            # 2) Lê o cabeçalho (descartando colunas vazias)
            header = df.iloc[i].dropna().tolist()
            header_length = len(header)
            i += 1

            linhas = []
            # 3) Lê as linhas até encontrar uma linha completamente vazia
            while i < len(df) and not df.iloc[i].isnull().all():
                row_values = df.iloc[i].tolist()
                
                # ➡ Garante que só fiquem as colunas que correspondem ao cabeçalho
                row_values = row_values[:header_length]
                
                # Se a linha tiver menos colunas que o cabeçalho, preenche com None
                if len(row_values) < header_length:
                    row_values += [None] * (header_length - len(row_values))
                
                linhas.append(row_values)
                i += 1

            tabelas.append({
                "nome": tabela_nome,
                "cabecalho": header,
                "linhas": linhas
            })
        return {
            "carteira": carteira_name,
            "sheet": sheet_name,
            "tabelas": tabelas
        }

    def process_file(self, file):
        filename = os.path.basename(file)
        try:
            xls = pd.ExcelFile(file)
        except Exception as e:
            raise Exception(f"Erro ao abrir {filename}: {e}")
        carteiras = []
        for sheet_name in xls.sheet_names:
            try:
                df = pd.read_excel(file, sheet_name=sheet_name, header=None)
                carteira_data = self.process_sheet(df, sheet_name)
                carteiras.append(carteira_data)
            except Exception as e:
                self.logger.error(f"Erro ao processar a sheet {sheet_name} em {filename}: {e}")
                continue
        
        # Monta o objeto final em JSON
        data_json = {
            "data": filename.replace(".xlsx", ""),
            "carteiras": carteiras
        }

        # Define o nome de saída
        output_filename = filename.replace(".xlsx", ".json")
        output_path = os.path.join(self.output_folder, output_filename)

        # Salva o JSON sem as colunas vazias extras
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(data_json, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    processor = CarteiraDiariaProcessor(
        input_folder=r"raw_data\carteira_diaria",
        output_folder=r"downloads\banvox\carteira_diaria"
    )
    processor.process_all_files()
