# Controller.py
import os
from pathlib import Path
from src.Model import DataModel
from src.View import ExcelView
import configparser
import logging
import time

class Controller:
    def __init__(self, project_root: Path):
        """Inicialização do Controller com configuração e componentes de modelo e visão."""
        self.project_root = project_root
        self.config = self.load_config()
        self.model = DataModel(self.project_root, self.config)
        self.view = ExcelView(self.config, self.project_root)
        logging.basicConfig(filename='controller.log', level=logging.INFO)

    def load_config(self) -> configparser.ConfigParser:
        """Carrega configurações do arquivo Config.ini."""
        config = configparser.ConfigParser()
        config.read(self.project_root / 'resources' / 'Config.ini')
        return config

    def run(self) -> None:
        """Executa o processo completo de carregamento, processamento e salvamento dos dados."""
        start_time = time.time()
        try:
            self.load_all_data()
            self.process_and_save_data()
            end_time = time.time()
            logging.info(f"Processo completo executado em {end_time - start_time:.2f} segundos")
        except Exception as e:
            logging.error(f"Erro durante a execução: {str(e)}")
            raise

    def load_all_data(self) -> None:
        """Carrega todos os arquivos CSV do diretório de entrada."""
        input_dir = self.project_root / self.config['DEFAULT']['InputDirectory']
        for file in sorted(os.listdir(input_dir)):
            if file.endswith('.csv'):
                self.model.load_data(input_dir / file)

    def process_and_save_data(self) -> None:
        """Processa os dados e atualiza o arquivo Excel."""
        try:
            self.model.remove_duplicates()
            df_unified, df_evol, df_analysis = self.model.process_data()
            output_file = self.project_root / self.config['DEFAULT']['OutputDirectory'] / self.config['DEFAULT']['OutputFileName']
            self.view.update_excel(str(output_file), df_unified, df_evol, df_analysis)
            
            logging.info(f"\nResumo:")
            logging.info(f"Total de registros processados: {len(df_unified)}")
            logging.info(f"Número de municípios: {df_unified['MUNICIPIO'].nunique()}")
            logging.info(f"Anos cobertos: {', '.join(sorted(df_unified['ANO'].unique()))}")
            logging.info(f"Arquivo Excel atualizado: {output_file}")
        except Exception as e:
            logging.error(f"Erro ao processar e salvar dados: {str(e)}")
            raise