import os
import sys
import time
import logging

from config_loader import load_config, create_default_config
from tesseract_setup import setup_tesseract
from logging_setup import setup_logging
from file_processing import load_keywords
from search_engine import search_files

# Глобальные флаги для доступности функций
HAS_PDF = False
HAS_DOCX = False
HAS_EXCEL = False
HAS_7Z = False
HAS_RAR = False
HAS_OCR = False


def check_dependencies():
    """Проверка доступности опциональных зависимостей"""
    global HAS_PDF, HAS_DOCX, HAS_EXCEL, HAS_7Z, HAS_RAR, HAS_OCR

    # Проверяем доступность модулей
    try:
        import fitz
        HAS_PDF = True
    except ImportError:
        logging.warning("Модуль PyMuPDF не установлен. Поддержка PDF отключена.")

    try:
        import docx2txt
        HAS_DOCX = True
    except ImportError:
        logging.warning("Модули для DOCX не установлены. Поддержка DOCX отключена.")

    try:
        import pandas as pd
        import openpyxl
        HAS_EXCEL = True
    except ImportError:
        logging.warning("Модули для Excel не установлены. Поддержка Excel отключена.")

    try:
        import py7zr
        HAS_7Z = True
    except ImportError:
        logging.warning("Модуль py7zr не установлен. Поддержка 7z архивов отключена.")

    try:
        import rarfile
        HAS_RAR = True
    except ImportError:
        logging.warning("Модуль rarfile не установлен. Поддержка RAR архивов отключена.")

    # Tesseract настраивается отдельно
    HAS_OCR = setup_tesseract()


def main():
    # Загружаем конфигурацию
    config = load_config()

    # Настраиваем логирование
    setup_logging(config.get('log_file', 'search_log.txt'))

    # Проверяем, существует ли файл конфигурации
    if not os.path.exists("config.txt"):
        logging.warning("Файл конфигурации config.txt не найден.")
        create_default_config()
        logging.info("Пожалуйста, настройте config.txt и запустите программу снова.")
        return

    # Проверяем зависимости
    check_dependencies()

    extensions = config['extensions']
    keywords_file = config['keywords_file']
    directory = config['directory']
    threads = config['threads']
    output_file = config['output_file']
    search_images = config['search_images']
    max_file_size = config['max_file_size']

    # Проверяем существование файла с ключевыми словами
    if not os.path.exists(keywords_file):
        logging.error(f"Файл с ключевыми словами '{keywords_file}' не найден.")
        logging.info("Пожалуйста, создайте файл keywords.txt с ключевыми словами (каждое слово с новой строки).")
        return

    # Проверяем существование директории для поиска
    if not os.path.exists(directory):
        logging.error(f"Директория '{directory}' не существует.")
        return

    try:
        keywords = load_keywords(keywords_file)
    except ValueError as e:
        logging.error(e)
        return

    if not keywords:
        logging.error("Не найдено ключевых слов.")
        return

    logging.info(f"Загружена конфигурация из config.txt")
    logging.info(f"Директория для поиска: {directory}")
    logging.info(f"Используемые маски: {extensions}")
    logging.info(f"Файл с ключевыми словами: {keywords_file}")
    logging.info(f"Количество ключевых слов: {len(keywords)}")
    logging.info(f"Используется потоков: {threads}")
    logging.info(f"Файл для результатов: {output_file}")
    logging.info(f"Максимальный размер файла: {max_file_size} МБ")
    logging.info(f"Поиск по изображениям: {'включен' if search_images and HAS_OCR else 'отключен'}")

    # Выводим информацию о доступных функциях
    if not HAS_PDF:
        logging.warning("Обработка PDF недоступна")
    if not HAS_DOCX:
        logging.warning("Обработка DOCX недоступна")
    if not HAS_EXCEL:
        logging.warning("Обработка Excel недоступна")
    if not HAS_OCR and search_images:
        logging.warning("Поиск по изображениям недоступен (Tesseract не установлен)")
    if not HAS_7Z:
        logging.warning("Обработка 7z архивов недоступна")
    if not HAS_RAR:
        logging.warning("Обработка RAR архивов недоступна")

    # Выполняем поиск
    start_time = time.time()
    results = search_files(directory, extensions, threads, output_file, max_file_size, config)
    end_time = time.time()

    if results:
        logging.info(f"Найдено совпадений в {len(results)} файлах:")

        logging.info(f"Результаты также сохранены в файл {output_file}")
    else:
        logging.info("Ничего не найдено.")

    logging.info(f"Время выполнения: {end_time - start_time:.2f} секунд")


if __name__ == '__main__':
    main()