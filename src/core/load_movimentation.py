from utils.copy_format_cells import copy_format
from datetime import datetime

def load_movimentation_logum(source_sheet, destination_sheet, source_row=12):
    """
    Carrega os dados de movimentação da planilha de origem para a planilha de destino.

    Args:
        source_sheet (Worksheet): A planilha de origem.
        destination_sheet (Worksheet): A planilha de destino.
        source_row (int): O número da linha na planilha de origem a ser copiada.
    """
    try:
        # Mapeamento das colunas da planilha de origem para a planilha de destino
        column_mapping = {
            3: 1,   # mes_de_referencia -> mes_de_referencia
            4: 2,   # codigo_simp_do_duto -> codigo_simp_do_duto
            6: 4,   # codigo_simp_do_trecho -> codigo_simp_do_trecho
            15: 13,   # codigo_simp_do_produto -> codigo_simp_do_produto
            16: 14,   # nome_produto -> nome_produto
            17: 15,   # volume_em_m3 -> volume_em_m3
        }

        # Determina a próxima linha vazia na planilha de destino
        destination_row = destination_sheet.max_row + 1

        for src_col, dest_col in column_mapping.items():
            destination_sheet.cell(row=destination_row, column=dest_col).value = source_sheet.cell(row=source_row, column=src_col).value
            copy_format(source_sheet.cell(row=source_row, column=src_col), destination_sheet.cell(row=destination_row, column=dest_col))

        # Adiciona o nome da transportadora na coluna "empresa" (coluna 16)
        dest_cell = destination_sheet.cell(row=destination_row, column=16)
        dest_cell.value = "LOGUM"
           
        # Adiciona a data e hora atual na coluna "data_atualizacao" (coluna 17)
        current_datetime = datetime.now().strftime("%d/%m/%Y")
        dest_cell = destination_sheet.cell(row=destination_row, column=17)
        dest_cell.value = current_datetime

        # Formata a célula da data e hora
        dest_cell.number_format = "DD/MM/YYYY"
    
    except Exception as e:
        raise ValueError(f"Erro ao preencher células no destino na linha {source_row}: {e}")

def load_movimentation_transpetro(source_sheet, destination_sheet, source_row):
    """
    Carrega os dados de movimentação da planilha de origem para a planilha de destino.

    Args:
        source_sheet (Worksheet): A planilha de origem.
        destination_sheet (Worksheet): A planilha de destino.
        source_row (int): O número da linha na planilha de origem a ser copiada.
    """
    try:
        # Mapeamento das colunas da planilha de origem para a planilha de destino
        column_mapping = {
            1: 1,   # mes_de_referencia -> mes_de_referencia
            2: 2,   # codigo_simp_do_duto -> codigo_simp_do_duto
            4: 4,   # codigo_simp_do_trecho -> codigo_simp_do_trecho
            13: 13,   # codigo_simp_do_produto -> codigo_simp_do_produto
            14: 14,   # nome_produto -> nome_produto
            15: 15,   # volume_em_m3 -> volume_em_m3
        }

        # Determina a próxima linha vazia na planilha de destino
        destination_row = destination_sheet.max_row + 1

        for src_col, dest_col in column_mapping.items():
            destination_sheet.cell(row=destination_row, column=dest_col).value = source_sheet.cell(row=source_row, column=src_col).value
            copy_format(source_sheet.cell(row=source_row, column=src_col), destination_sheet.cell(row=destination_row, column=dest_col))

        # Adiciona o nome da transportadora na coluna "empresa" (coluna 16)
        dest_cell = destination_sheet.cell(row=destination_row, column=16)
        dest_cell.value = "TRANSPETRO"
           
        # Adiciona a data e hora atual na coluna "data_atualizacao" (coluna 17)
        current_datetime = datetime.now().strftime("%d/%m/%Y")
        dest_cell = destination_sheet.cell(row=destination_row, column=17)
        dest_cell.value = current_datetime

        dest_cell.number_format = "DD/MM/YYYY"
    
    except Exception as e:
        raise ValueError(f"Erro ao preencher células no destino na linha {source_row}: {e}")