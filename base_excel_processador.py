import os
import glob
import json
import logging
import pandas as pd
from abc import ABC, abstractmethod

class BaseExcelProcessor(ABC):
    def __init__(self, input_folder, output_folder, processed_log_filename="processed_files.json"):
        self.input_folder = input_folder
        self.output_folder = output_folder
        os.makedirs(self.output_folder, exist_ok=True)
        self.processed_log_path = os.path.join(self.output_folder, processed_log_filename)
        self.processed_files = self.load_processed_files()
        self.logger = self.setup_logger()

    def setup_logger(self):
        logger = logging.getLogger(self.__class__.__name__)
        logger.setLevel(logging.INFO)
        # Cria o arquivo de log no mesmo diretório de saída
        log_file = os.path.join(self.output_folder, "processing.log")
        if not logger.handlers:
            fh = logging.FileHandler(log_file)
            fh.setLevel(logging.INFO)
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            fh.setFormatter(formatter)
            logger.addHandler(fh)
        return logger

    def load_processed_files(self):
        if os.path.exists(self.processed_log_path):
            with open(self.processed_log_path, "r", encoding="utf-8") as f:
                return json.load(f)
        else:
            return []

    def save_processed_files(self):
        with open(self.processed_log_path, "w", encoding="utf-8") as f:
            json.dump(self.processed_files, f, ensure_ascii=False, indent=2)

    def get_valid_excel_files(self):
        """Retorna a lista de arquivos Excel cujo nome segue o padrão de data (YYYY-MM-DD.xlsx)."""
        excel_files = glob.glob(os.path.join(self.input_folder, "*.xlsx"))
        valid_files = []
        for file in excel_files:
            filename = os.path.basename(file)
            try:
                pd.to_datetime(filename.replace(".xlsx", ""))
                valid_files.append(file)
            except ValueError:
                continue
        return valid_files

    def process_all_files(self):
        files = self.get_valid_excel_files()
        for file in files:
            filename = os.path.basename(file)
            if filename in self.processed_files:
                self.logger.info(f"{filename} já foi processado. Pulando...")
                continue

            self.logger.info(f"Processando {filename}...")
            try:
                self.process_file(file)
                self.processed_files.append(filename)
                self.save_processed_files()
                self.logger.info(f"{filename} processado e salvo.")
            except Exception as e:
                self.logger.error(f"Erro ao processar {filename}: {e}")

    @abstractmethod
    def process_file(self, file):
        """Implementa o processamento específico para o arquivo Excel."""
        pass
