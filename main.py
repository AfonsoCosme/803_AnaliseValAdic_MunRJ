# Main
import sys
from pathlib import Path
from src.Controller import Controller

if sys.version_info < (3, 12):
    print("Este script requer Python 3.12 ou superior.")
    sys.exit(1)

if __name__ == "__main__":
    project_root = Path(__file__).parent
    controller = Controller(project_root)
    controller.run()