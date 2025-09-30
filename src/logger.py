import os
import logging

def get_logger():
    os.makedirs("logs", exist_ok=True)
    logger = logging.getLogger("Carga MOT")
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        handler = logging.FileHandler("logs/execucao.log", encoding="utf-8")
        formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    return logger