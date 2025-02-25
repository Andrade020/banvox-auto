import os
import json
import pandas as pd
from base_excel_processador import BaseExcelProcessor

class ExtratosProcessor(BaseExcelProcessor):
    def process_file(self, file):
        filename = os.path.basename(file)
        print(f"🔍 Processando arquivo: {filename}")

        try:
            df = pd.read_excel(file, header=None)
        except Exception as e:
            raise Exception(f"❌ Erro ao ler {filename}: {e}")

        if len(df) < 4:
            raise Exception(f"⚠️ {filename} não possui a estrutura esperada. Pulando...")

        print("✅ Arquivo carregado com sucesso! Exibindo primeiras 10 linhas para verificação:")
        print(df.head(10))

        # 📌 Linha 0: Cabeçalho global (usado para todas as tabelas)
        global_header = df.iloc[0].tolist()
        print(f"📌 Cabeçalho global extraído: {global_header}")

        carteiras = []
        current_carteira_name = None
        current_data = []

        # 🔎 Percorrer o arquivo para identificar cada carteira e sua tabela correspondente
        i = 1
        while i < len(df):
            row = df.iloc[i]

            # 🏦 Se a primeira coluna contém um texto que identifica uma carteira
            if isinstance(row[0], str) and row[0].startswith("Nome Carteira:"):
                # Se já temos dados acumulados da carteira anterior, salvamos
                if current_carteira_name and current_data:
                    carteiras.append({
                        "carteira": current_carteira_name,
                        "cabecalho": global_header,
                        "linhas": current_data
                    })

                # 🆕 Atualiza para a nova carteira e zera os dados anteriores
                current_carteira_name = row[0].replace("Nome Carteira:", "").strip()
                current_data = []
                print(f"🏦 Nova carteira detectada: {current_carteira_name}")

                # A próxima linha (linha i+1) deve ser ignorada (contém "Conta: Todas" ou outro dado irrelevante)
                i += 2  # Pula a linha da carteira e a linha seguinte

            elif current_carteira_name:
                # 🔍 Verifica se a primeira coluna **NÃO** contém uma data -> significa que chegou ao fim da carteira atual
                try:
                    pd.to_datetime(row[0], dayfirst=True)
                    current_data.append(row.tolist())  # ✅ Adiciona os dados da tabela
                except (ValueError, TypeError):
                    print(f"⚠️ Quebra detectada! {row[0]} não é uma data. Mudando de carteira.")
                    i -= 1  # Retrocede para que essa linha seja avaliada como uma possível nova carteira na próxima iteração

            i += 1

        # 🔚 Após terminar a leitura, salva os dados da última carteira encontrada
        if current_carteira_name and current_data:
            carteiras.append({
                "carteira": current_carteira_name,
                "cabecalho": global_header,
                "linhas": current_data
            })

        # 📂 Salvar os dados processados no JSON
        output_filename = filename.replace(".xlsx", ".json")
        output_path = os.path.join(self.output_folder, output_filename)

        data_json = {
            "data": filename.replace(".xlsx", ""),
            "carteiras": carteiras
        }

        #print(f"📄 JSON gerado:\n{json.dumps(data_json, indent=2, ensure_ascii=False, default=str)}")

        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(data_json, f, ensure_ascii=False, indent=2, default=str)

        print(f"✅ JSON salvo com sucesso em {output_path}\n")


if __name__ == "__main__":
    processor = ExtratosProcessor(
        input_folder=r"raw_data\extratos",
        output_folder=r"downloads\banvox\extratos"
    )

    print("🔍 Verificando arquivos para processamento...")
    processor.process_all_files()
