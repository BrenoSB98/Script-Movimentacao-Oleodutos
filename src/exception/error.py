import logging

def setup():
    """Configura o logging para registrar erros em um arquivo de log."""
    logging.basicConfig(
        filename="error_log.txt",
        level=logging.ERROR,
        format="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

def error(message):
    """Registra o erro no log e exibe a mensagem para o usuário."""
    setup()
    logging.error(message)
    print(f"Erro: {message}")
