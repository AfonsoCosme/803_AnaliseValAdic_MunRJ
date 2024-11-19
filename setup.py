# setup.py
import sys
import os
import json
from pathlib import Path

if sys.version_info < (3, 12):
    print("Este script requer Python 3.12 ou superior.")
    sys.exit(1)

def create_directory(path: Path) -> None:
    if not path.exists():
        path.mkdir(parents=True)
        print(f"Criado diretório: {path}")

def create_file(path: Path, content: str = "") -> None:
    if not path.exists():
        with open(path, 'w', encoding='utf-8') as file:
            file.write(content)
        print(f"Criado arquivo: {path}")
    else:
        print(f"Arquivo já existe (não sobrescrito): {path}")

def setup_project() -> None:
    project_root = Path(__file__).parent

    # Estrutura de diretórios
    directories = [
        project_root / "src",
        project_root / "src" / "utils",
        project_root / "resources",
        project_root / "data" / "input",
        project_root / "data" / "output",
        project_root / "tests",
        project_root / "docs",
        project_root / "logs"
    ]

    # Criar diretórios
    for directory in directories:
        create_directory(directory)

    # Arquivos a serem criados
    files = {
        project_root / "src" / "__init__.py": "",
        project_root / "src" / "Constants.py": "# Constantes do projeto",
        project_root / "src" / "utils" / "__init__.py": "",
        project_root / "src" / "utils" / "helpers.py": "# Funções auxiliares",
        project_root / "docs" / "README.md": "# TabulaValAdic Project\n\nDescrição do projeto e instruções de uso.",
        project_root / "-Requirements.txt": "# Dependências do projeto\npandas>=2.0.0\nopenpyxl>=3.1.0",
        project_root / "TabulaValAdic.spec": "# Especificações para geração do executável",
        project_root / "runtime_hook.py": "# Hook de tempo de execução",
        project_root / "upx.conf": "# Configuração UPX",
        project_root / "file_version_info.txt": "# Informações de versão do arquivo",
        project_root / "version.py": "VERSION = '0.1.0'",
    }

    # Criar arquivos
    for file_path, content in files.items():
        create_file(file_path, content)

    # Criar Config.ini
    config_content = """[DEFAULT]
InputDirectory = data/input
OutputDirectory = data/output
OutputFileName = Tabula_POR_2017a2023.xlsx

[LOGGING]
LogFile = logs/app.log
LogLevel = DEBUG

[PROCESSING]
YearColumns = _2017_r_,_2018_r_,_2019_r_,_2020_r_,_2021_r_,_2022_r_,_2023_r_
"""
    create_file(project_root / "resources" / "Config.ini", config_content)

    # Criar TAB_ApoioSigMun.json
    sig_mun_data = {
        "MGT": "Mangaratiba",
        "ITG": "Itaguaí",
        "POR": "Porto Real",
        "ARE": "Areal"
    }
    create_file(project_root / "resources" / "TAB_ApoioSigMun.json", json.dumps(sig_mun_data, indent=2))

    print("Estrutura do projeto criada com sucesso!")

if __name__ == "__main__":
    setup_project()