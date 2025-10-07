import os
import sys
import locale
from datetime import datetime, timedelta

import tkinter as tk
from tkinter import messagebox
from interface.select_worksheet import select_worksheets, create_backup, select_company_popup
from core.load_movimentation import load_movimentation_logum, load_movimentation_transpetro

from utils import load_workbooks
from exception import error
from logger import get_logger

logger = get_logger()

def generate_name_sheet_logum():
    """
    Gera o nome de uma planilha dinamicamente com base no mês anterior e ano corrente.

    Formato de retorno: 'NomeDoMês_Ano' (ex: 'Agosto_2025').
    Utiliza o locale 'pt_BR.UTF-8' para os nomes dos meses em português.

    Returns:
        str: O nome da planilha formatado.
    """
    try:
        locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
    except locale.Error:
        logger.info("Aviso: Locale 'pt_BR.UTF-8' não suportado. Usando o locale padrão do sistema.")

    today = datetime.now()
    first_day_actual_month = today.replace(day=1)
    date_month_before = first_day_actual_month - timedelta(days=1)

    name_month = date_month_before.strftime('%B').title()
    year_actual = date_month_before.year

    return f"{name_month}-{year_actual}", f"{name_month}_{year_actual}"

def main():
    """
        Função principal responsável por executar o fluxo completo de importação de dados de movimentação dos oleodutos a partir de múltiplas planilhas para uma planilha destino.

        Etapas:
        1. Seleciona qual empresa está sendo processada (Logum ou Transpetro).
        2. Seleciona arquivos de origem e destino via interface gráfica (Tkinter).
        3. Cria um backup automático da planilha destino.
        4. Carrega e processa cada linha das planilhas de origem, populando a planilha destino.
        5. Salva os dados e exibe mensagem de sucesso ou erro.

        Saída: encerra o processo com código de status (0 para sucesso, 1 para erro).
    """
    company = select_company_popup()
    if not company:
        logger.info("Nenhuma empresa foi selecionada.")
        sys.exit(1)
    
    destiny, origin_list = select_worksheets()
    if not destiny:
        logger.info("Nenhuma planilha de destino foi selecionada.")
        sys.exit(1)

    if not origin_list:
        logger.info("Nenhuma planilha de origem foi selecionada.")
        sys.exit(1)

    create_backup(destiny)

    try:             
        for origin_path in origin_list:
            file_name = os.path.basename(origin_path)
            try:
                logger.info(f"Carregando planilha: {file_name}")
                source_wb, dest_wb = load_workbooks(origin_path, destiny)

                if company.upper() == "LOGUM":
                    sheet_name_dash, sheet_name_underscore = generate_name_sheet_logum()
                    source_sheet = None
                    
                    if sheet_name_dash in source_wb.sheetnames:
                        source_sheet = source_wb[sheet_name_dash]
                    elif sheet_name_underscore in source_wb.sheetnames:
                        source_sheet = source_wb[sheet_name_underscore]
                    
                    if source_sheet is None:
                        error(f'Planilha com nome "{sheet_name_dash}" ou "{sheet_name_underscore}" não encontrada na planilha {file_name}.')
                        continue
                    
                    register_sheet = dest_wb["Movimentação"]

                    for row in range(12, source_sheet.max_row + 1):
                        try:
                            load_movimentation_logum(source_sheet, register_sheet, row)
                            logger.info(f"Linha {row} da planilha {file_name} processada com sucesso.")
                        except Exception as e:
                            error(f"Erro na linha {row} da planilha {file_name}: {e}")     
                
                else:  # Transpetro
                    source_sheet = source_wb["Movimentação"]
                    register_sheet = dest_wb["Movimentação"]                    
                    try:
                        locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
                    except locale.Error:
                        logger.info("Aviso: Locale 'pt_BR.UTF-8' não suportado. Usando o locale padrão do sistema.")

                    today = datetime.now()
                    first_day_actual_month = today.replace(day=1)
                    date_month_before = first_day_actual_month - timedelta(days=1)
                    search_date = date_month_before.replace(day=1)
                    
                    start_row_origin = None
                    for row in range(1, source_sheet.max_row + 1):
                        cell = source_sheet.cell(row=row, column=1)
                        if isinstance(cell.value, datetime) and cell.value.date() == search_date.date():
                            start_row_origin = row
                            break

                    if start_row_origin is None:
                        search_value_str = search_date.strftime('%d/%m/%Y')
                        error(f'Data "{search_value_str}" não encontrada na coluna 1 da planilha {file_name}.')
                        continue
                    
                    for row in range(start_row_origin, source_sheet.max_row + 1):
                        try:
                            load_movimentation_transpetro(source_sheet, register_sheet, row)
                            logger.info(f"Linha {row} da planilha {file_name} processada com sucesso.")
                        except Exception as e:
                            error(f"Erro na linha {row} da planilha {file_name}: {e}")
                               
                # Salva a planilha destino após processar cada origem
                logger.info(f"Salvando todas as alterações no destino: {destiny}")
                dest_wb.save(destiny)
                logger.info(f"Arquivo de destino salvo com sucesso.")
            
            except Exception as e:
                error(f"Erro ao processar a planilha {file_name}: {e}")                

        # Inicializa a raiz da interface Tkinter sem exibir a janela principal
        root = tk.Tk()
        root.withdraw()

        # Agenda ações para garantir que o messagebox apareça em primeiro plano
        root.after_idle(root.deiconify)
        root.after_idle(root.lift)
        root.after_idle(root.focus_force)

        # Exibe mensagem de sucesso após a carga
        messagebox.showinfo("Finalizado", "Processamento concluído com sucesso!", parent=root)
        root.destroy()

        logger.info("Processamento de todas as planilhas concluído.")
        sys.exit(0)

    except Exception as e:
        error(f"Erro fatal na execução: {e}")
        root = tk.Tk()
        root.withdraw()
        root.after_idle(root.deiconify)
        root.after_idle(root.lift)
        root.after_idle(root.focus_force)
        messagebox.showerror("Erro", f"Erro fatal: {e}", parent=root)
        root.destroy()
        sys.exit(1)

if __name__ == "__main__":
    main()
