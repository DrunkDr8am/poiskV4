import os
import logging
import sys


def setup_logging(log_file='search_log.txt'):
    """Настройка логирования с очисткой предыдущих логов"""
    # Удаляем предыдущий файл логов, если он существует
    if os.path.exists(log_file):
        try:
            os.remove(log_file)
        except Exception as e:
            print(f"Не удалось удалить старый файл логов: {e}")

    # Создаем новый логгер с очисткой предыдущих обработчиков
    logger = logging.getLogger()
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)

    # Настраиваем логирование
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, mode='w', encoding='utf-8'),  # 'w' для перезаписи файла
            logging.StreamHandler(sys.stdout)
        ]
    )