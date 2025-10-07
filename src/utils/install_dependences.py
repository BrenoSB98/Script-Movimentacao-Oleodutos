import subprocess
import sys

def install_openpyxl():
    """
        Verifica se a biblioteca 'openpyxl' está instalada e a instala silenciosamente, se necessário.

        Esse tipo de verificação é útil quando o script precisa rodar em ambientes onde a dependência
        pode não estar previamente instalada.
    """
    try:
        import openpyxl
    except ImportError:
        # Executa um subprocesso para instalar o pacote usando o pip e o Python atual
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "--quiet"])

def install_dotenv():
    """
        Verifica se a biblioteca 'python-dotenv' está instalada e a instala silenciosamente, se necessário.

        Esse tipo de verificação é útil quando o script precisa rodar em ambientes onde a dependência
        pode não estar previamente instalada.
    """
    try:
        import dotenv
    except ImportError:
        # Executa um subprocesso para instalar o pacote usando o pip e o Python atual
        subprocess.check_call([sys.executable, "-m", "pip", "install", "python-dotenv", "--quiet"])

def install_pillow():
    """
        Verifica se a biblioteca 'Pillow' está instalada e a instala silenciosamente, se necessário.

        Esse tipo de verificação é útil quando o script precisa rodar em ambientes onde a dependência pode não estar previamente instalada.
    """
    try:
        import PIL
    except ImportError:
        # Executa um subprocesso para instalar o pacote usando o pip e o Python atual
        subprocess.check_call([sys.executable, "-m", "pip", "install", "Pillow", "--quiet"])