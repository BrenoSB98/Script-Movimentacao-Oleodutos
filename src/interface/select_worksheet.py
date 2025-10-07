import os
import shutil
from datetime import datetime

import tkinter as tk
from tkinter import simpledialog
from tkinter import filedialog

from utils.install_dependences import install_dotenv, install_pillow

install_dotenv()
install_pillow()

from dotenv import load_dotenv
from PIL import Image, ImageTk

from logger import get_logger
from utils.install_dependences import install_dotenv

load_dotenv()
logger = get_logger()

backup_dir = os.getenv('BACKUP_DIR')
backup_dir = os.path.normpath(backup_dir) if backup_dir else None

def select_worksheets():
    """
        Abre uma interface gráfica para o usuário selecionar a planilha de destino e uma ou mais planilhas de origem.

        Returns:
            tuple:
                - str or None: Caminho para a planilha de destino selecionada, ou None se não selecionado.
                - list of str: Lista com os caminhos das planilhas de origem selecionadas.
    """

    root = tk.Tk()
    root.withdraw()

    # Seleção da planilha de destino (MOV_OLEODUTOS_NF)
    destination = filedialog.askopenfilename(
        title="Selecione a planilha de DESTINO (MOV_OLEODUTOS_NF)",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not destination:
        root.destroy()
        return None, []

    # Seleção da planilha de DESTINO (planilhas das Transportadoras)
    source = filedialog.askopenfilenames(
        title="Selecione uma ou mais planilhas de movimentação das TRANSPORTADORAS",
        filetypes=[("Excel files", "*.xlsx")]
    )
    root.destroy()
    return destination, list(source)

def create_backup(path_file):
    """
        Cria uma cópia de backup do arquivo Excel informado no diretório especificado pela variável de ambiente BACKUP_DIR.
        O nome do backup conterá um timestamp para identificação única.

        Args:
            path_file (str): Caminho absoluto do arquivo de destino que será copiado.

        Returns:
            str: Caminho completo do arquivo de backup criado.
    """

    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)

    _, name = os.path.split(path_file)
    base, ext = os.path.splitext(name)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    name_backup = f"{timestamp}_{base}.backup{ext}"
    path_backup = os.path.join(backup_dir, name_backup)

    # Copia o arquivo original para o novo caminho de backup
    shutil.copy(path_file, path_backup)
    logger.info(f"Backup criado: {path_backup}")
    return path_backup

def mark_file_as_registered(path_file):
    """
        Renomeia o arquivo de origem adicionando o prefixo REGISTRADO_{DATA/HORA}_ ao nome do arquivo.

        Args:
            path_file (str): Caminho absoluto do arquivo de origem que será renomeado.

        Returns:
            str: Novo caminho completo do arquivo renomeado.
    """
    dir_name, file_name = os.path.split(path_file)
    base, ext = os.path.splitext(file_name)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    new_file_name = f"REGISTRADO_{timestamp}_{base}{ext}"
    new_path = os.path.join(dir_name, new_file_name)
    os.rename(path_file, new_path)
    logger.info(f"Arquivo registrado: {new_path}")
    return new_path

def select_company_popup():
    """
        Exibe um pop-up com checkboxes para o usuário escolher entre as empresas LOGUM e TRANSPETRO.
        Permite selecionar apenas uma opção.

        Returns:
            str: 'LOGUM', 'TRANSPETRO' ou None se nenhuma opção for selecionada.
    """
    class CompanyDialog(simpledialog.Dialog):
        def __init__(self, parent, title=None):
            self.photo = None
            try:
                image = Image.open('src\assets\icon.jpeg')
                self.photo = ImageTk.PhotoImage(image)
            except Exception as e:
                logger.warning(f"Ícone 'src\assets\icon.jpeg' não encontrado ou erro ao carregar: {e}. A janela será exibida com o ícone padrão.")
            
            super().__init__(parent, title=title)
        def body(self, master):
            self.geometry("300x150")
            if self.photo:
                self.wm_iconphoto(True, self.photo)
            
            self.var = tk.StringVar(value="")
            tk.Label(master, text="Selecione a empresa:").pack(anchor="w", padx=10, pady=5)
            self.rb1 = tk.Radiobutton(master, text="LOGUM", variable=self.var, value="LOGUM")
            self.rb2 = tk.Radiobutton(master, text="TRANSPETRO", variable=self.var, value="TRANSPETRO")
            self.rb1.pack(anchor="w", padx=20)
            self.rb2.pack(anchor="w", padx=20)
            return self.rb1

        def apply(self):
            self.result = self.var.get()

    root = tk.Tk()
    root.withdraw()
    
    dialog = CompanyDialog(root, title="Escolha da Empresa")
         
    root.destroy()
    return dialog.result if dialog.result else None
