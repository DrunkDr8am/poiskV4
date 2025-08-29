import os
import configparser

def load_config(config_file="config.txt"):
    """Загрузка конфигурации из файла"""
    config = configparser.ConfigParser()

    # Значения по умолчанию
    defaults = {
        'extensions': ['*.txt', '*.pdf', '*.docx', '*.xlsx', '*.jpg', '*.png', '*.zip', '*.rar', '*.7z'],
        'keywords_file': 'keywords.txt',
        'directory': '.',
        'threads': '4',
        'output_file': 'search_results.txt',
        'search_images': 'false',
        'max_file_size': '50',
        'log_file': 'search_log.txt',
        'tesseract_languages': 'rus',
        'tesseract_config': '--oem 3 --psm 6'
    }

    # Если файл конфигурации существует, загружаем его
    if os.path.exists(config_file):
        try:
            config.read(config_file, encoding='utf-8')
        except Exception as e:
            print(f"Ошибка чтения конфигурационного файла: {e}")
            return defaults

    # Получаем значения из конфига или используем значения по умолчанию
    extensions = config.get('Settings', 'extensions', fallback=','.join(defaults['extensions'])).split(',')
    keywords_file = config.get('Settings', 'keywords_file', fallback=defaults['keywords_file'])
    directory = config.get('Settings', 'directory', fallback=defaults['directory'])
    threads = config.getint('Settings', 'threads', fallback=int(defaults['threads']))
    output_file = config.get('Settings', 'output_file', fallback=defaults['output_file'])
    search_images = config.getboolean('Settings', 'search_images', fallback=False)
    max_file_size = config.getint('Settings', 'max_file_size', fallback=int(defaults['max_file_size']))
    log_file = config.get('Settings', 'log_file', fallback=defaults['log_file'])
    tesseract_languages = config.get('Settings', 'tesseract_languages', fallback=defaults['tesseract_languages'])
    tesseract_config = config.get('Settings', 'tesseract_config', fallback=defaults['tesseract_config'])

    # Очищаем значения от пробелов
    extensions = [ext.strip() for ext in extensions]
    keywords_file = keywords_file.strip()
    directory = directory.strip()
    output_file = output_file.strip()
    log_file = log_file.strip()

    # Если поиск по изображениям отключен, убираем изображения из расширений
    if not search_images:
        extensions = [ext for ext in extensions if
                      not ext.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff'))]

    return {
        'extensions': extensions,
        'keywords_file': keywords_file,
        'directory': directory,
        'threads': threads,
        'output_file': output_file,
        'search_images': search_images,
        'max_file_size': max_file_size,
        'log_file': log_file,
        'tesseract_languages': tesseract_languages,
        'tesseract_config': tesseract_config
    }

def create_default_config():
    """Создание файла конфигурации по умолчанию"""
    config_content = """[Settings]
# Расширения файлов для поиска (через запятую)
extensions = *.txt, *.pdf, *.docx, *.xlsx, *.jpg, *.png, *.zip, *.rar, *.7z

# Файл с ключевыми словами (каждое слово с новой строки)
keywords_file = keywords.txt

# Директория для поиска
directory = .

# Количество потоков для обработки
threads = 4

# Файл для сохранения результатов
output_file = search_results.txt

# Поиск по изображениям (требует установленного Tesseract OCR)
search_images = true

# Максимальный размер обрабатываемого файла (МБ)
max_file_size = 50

# Файл для логирования
log_file = search_log.txt

# Настройки Tesseract OCR
tesseract_languages = rus
tesseract_config = --oem 3 --psm 6
"""

    with open("config.txt", "w", encoding="utf-8") as f:
        f.write(config_content)

    print("Создан файл конфигурации config.txt с настройками по умолчанию.")