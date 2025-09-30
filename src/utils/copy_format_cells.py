from copy import copy

def copy_format(source_cell, dest_cell):
    """
    Copia o estilo de formatação da célula de origem para a célula de destino.

    Args:
        source_cell (Cell): Célula de origem.
        dest_cell (Cell): Célula de destino.
    """
    try:
        dest_cell.font = copy(source_cell.font)
        dest_cell.fill = copy(source_cell.fill)
        dest_cell.number_format = source_cell.number_format
        dest_cell.protection = copy(source_cell.protection)
    except AttributeError as e:
        raise ValueError(f"Erro ao copiar a formatação da célula: {e}")
    except Exception as e:
        raise ValueError(f"Erro ao copiar a formatação: {e}")