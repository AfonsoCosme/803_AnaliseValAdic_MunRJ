# Model.py
import pandas as pd
import numpy as np
import json
from pathlib import Path
from typing import Dict, Tuple
import configparser
import logging
import time

class DataModel:
    def __init__(self, project_root: Path, config: configparser.ConfigParser):
        """Inicializa o modelo de dados com configuração e caminho do projeto."""
        self.project_root = project_root
        self.config = config
        self.data: pd.DataFrame = pd.DataFrame()
        self.sig_mun_map: Dict[str, str] = self.load_sig_mun_map()
        
        logging.basicConfig(filename='data_model.log', level=logging.INFO)

    def load_sig_mun_map(self) -> Dict[str, str]:
        """Carrega o mapa de SigMun a partir de um arquivo JSON."""
        with open(self.project_root / 'resources' / 'TAB_ApoioSigMun.json', 'r') as f:
            return json.load(f)

    def load_data(self, file_path: Path) -> None:
        """Carrega dados de um arquivo CSV e os adiciona ao DataFrame principal."""
        start_time = time.time()
        cols = ['Inscricao', 'CPF_CNPJ', 'Nome', 'Nome_Cidade']
        df = pd.read_csv(file_path, sep=';', usecols=lambda x: x in cols or x.endswith('(R$)'), decimal=',', thousands='.', encoding='iso-8859-1')
        df = df.rename(columns={
            'Inscricao': 'InscEst',
            'Nome': 'RazSoc',
            'Nome_Cidade': 'MUNICIPIO'
        })
        df['SigMun'] = df['MUNICIPIO'].map(lambda x: next((k for k, v in self.sig_mun_map.items() if v == x), ''))
        
        # Transformação dos dados para formato longo
        df_melted = df.melt(id_vars=['MUNICIPIO', 'InscEst', 'CPF_CNPJ', 'RazSoc', 'SigMun'], var_name='ANO', value_name='VALOR')
        #df_melted['ANO'] = df_melted['ANO'].str.extract('(\d{4})')  # Extrai o ano do nome da coluna
        df_melted['ANO'] = df_melted['ANO'].str.extract(r'(\d{4})')
        df_melted['VALOR'] = df_melted['VALOR'].fillna(0).round(2)
        
        self.data = pd.concat([self.data, df_melted], ignore_index=True)
        
        end_time = time.time()
        logging.info(f"Dados do arquivo {file_path.name} carregados com sucesso. Shape: {df_melted.shape}. Tempo de processamento: {end_time - start_time:.2f} segundos")

    def remove_duplicates(self) -> None:
        """Remove duplicatas do DataFrame principal."""
        before_count = len(self.data)
        self.data = self.data.drop_duplicates(subset=['MUNICIPIO', 'InscEst', 'CPF_CNPJ', 'RazSoc', 'ANO', 'SigMun'], keep='first')
        after_count = len(self.data)
        logging.info(f"Duplicatas removidas. Registros antes: {before_count}, depois: {after_count}")

    def process_data(self) -> Tuple[pd.DataFrame, Dict[str, pd.DataFrame], pd.DataFrame]:
        """Processa os dados, calculando variações e preparando DataFrames para análise."""
        start_time = time.time()
        
        # Criação do DataFrame unificado
        df_unified = self.data.sort_values(['SigMun', 'MUNICIPIO', 'InscEst', 'ANO'])
        df_unified = df_unified[['SigMun', 'MUNICIPIO', 'InscEst', 'CPF_CNPJ', 'RazSoc', 'ANO', 'VALOR']]
        
        # Identificação dinâmica dos anos
        anos_disponiveis = sorted(df_unified['ANO'].unique())
        
        # Criação do DataFrame de evolução
        df_evol = self.data.pivot_table(
            index=['SigMun', 'MUNICIPIO', 'InscEst', 'CPF_CNPJ', 'RazSoc'], 
            columns='ANO', 
            values='VALOR', 
            fill_value=0
        ).reset_index()

        # Garantia de que todos os anos estejam presentes no DataFrame de evolução
        for year in anos_disponiveis:
            if year not in df_evol.columns:
                df_evol[year] = 0

        df_analysis = self.calculate_variations(df_evol, anos_disponiveis)
        
        end_time = time.time()
        logging.info(f"Processamento de dados concluído. Tempo total: {end_time - start_time:.2f} segundos")
        
        return df_unified, df_evol, df_analysis

    def calculate_variations(self, df_evol: pd.DataFrame, anos_disponiveis: list) -> Dict[str, pd.DataFrame]:
        """Calcula as variações percentuais entre anos consecutivos para cada município."""
        analysis_dfs = {}
        
        for municipio in df_evol['MUNICIPIO'].unique():
            df_mun = df_evol[df_evol['MUNICIPIO'] == municipio].copy()
            df_result = df_mun.copy()
            
            # Lista para armazenar a sequência de colunas intercaladas
            colunas_intercaladas = ['MUNICIPIO', 'InscEst', 'CPF_CNPJ', 'RazSoc']
            
            # Adiciona o primeiro par de anos e a variação entre eles
            colunas_intercaladas.extend([anos_disponiveis[0], anos_disponiveis[1], f'{anos_disponiveis[0][-2:]}/{anos_disponiveis[1][-2:]} %'])
            
            # Calcula a variação para o primeiro par de anos
            df_result[f'{anos_disponiveis[0][-2:]}/{anos_disponiveis[1][-2:]} %'] = self._calculate_percentage_change(
                df_result[anos_disponiveis[0]], 
                df_result[anos_disponiveis[1]]
            )
            
            # Construção dinâmica das colunas intercaladas para anos subsequentes
            for i in range(1, len(anos_disponiveis) - 1):
                current_year = anos_disponiveis[i]
                next_year = anos_disponiveis[i + 1]
                var_col = f'{current_year[-2:]}/{next_year[-2:]} %'
                
                # Adiciona o próximo ano e a variação com o ano anterior
                colunas_intercaladas.extend([next_year, var_col])
                
                # Calcula a variação percentual
                df_result[var_col] = self._calculate_percentage_change(
                    df_result[current_year], 
                    df_result[next_year]
                )
            
            # Reorganiza o DataFrame com colunas intercaladas
            df_result = df_result[colunas_intercaladas]
            
            sig_mun = self.data[self.data['MUNICIPIO'] == municipio]['SigMun'].iloc[0]
            analysis_dfs[sig_mun] = df_result
            
            logging.info(f"Variações calculadas para {municipio}. Linhas processadas: {len(df_result)}")
        
        return analysis_dfs

    def _calculate_percentage_change(self, series1, series2):
        """Calcula a variação percentual entre duas séries."""
        return np.where(series1 != 0,
                        ((series2 - series1) / series1 * 100).round(2),
                        np.where(series2 != 0, 100, 0))