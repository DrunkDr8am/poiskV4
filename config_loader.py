import os
import configparser

IMAGE_EXTENSION_PATTERNS = (
    '*.jpg', '*.jpeg', '*.jpe', '*.jfif',
    '*.png', '*.bmp', '*.gif', '*.tif', '*.tiff', '*.webp', '*.ico'
)
WORD_EXTENSION_PATTERNS = ('*.doc', '*.docx', '*.docm', '*.dot', '*.dotx', '*.dotm')
EXCEL_EXTENSION_PATTERNS = ('*.xls', '*.xlsx', '*.xlsm', '*.xlt', '*.xltx', '*.xltm')


def _safe_getint(config, section, option, fallback):
    """Безопасно читает int из конфига, поддерживая пустые значения."""
    raw_value = config.get(section, option, fallback=str(fallback))
    if raw_value is None:
        return fallback
    value_str = str(raw_value).strip()
    if value_str == "":
        return fallback
    try:
        return int(value_str)
    except (TypeError, ValueError):
        return fallback


def load_config(config_file="config.txt"):
    """Загрузка конфигурации из файла"""
    config = configparser.ConfigParser()

    # Значения по умолчанию
    defaults = {
        'extensions': ['*.txt', '*.pdf', *WORD_EXTENSION_PATTERNS, *EXCEL_EXTENSION_PATTERNS, *IMAGE_EXTENSION_PATTERNS, '*.zip', '*.rar', '*.7z'],
        'keywords_file': 'keywords.txt',
        'directories': '.',
        'directory': '.',
        'theme': 'Светлая',
        'threads': '4',
        'output_file': 'search_results.txt',
        'search_images': 'false',
        'max_file_size': '50',
        'log_file': 'search_log.txt',
        'tesseract_languages': 'rus',
        'tesseract_config': '--oem 3 --psm 6',
        # Новые параметры для тонкой настройки
        'max_image_size_mb': '10',
        'max_pdf_pages': '0',  # 0 = без ограничения
        'max_excel_rows_per_sheet': '0',  # 0 = без ограничения
        'max_path_length': '240',
    }

    # Если файл конфигурации существует, загружаем его
    if os.path.exists(config_file):
        try:
            config.read(config_file, encoding='utf-8')
        except Exception as e:
            print(f"Ошибка чтения конфигурационного файла: {e}")
            return defaults

    # Пустой или некорректный конфиг может не содержать секцию Settings.
    # Добавляем ее, чтобы безопасно использовать fallback для всех параметров.
    if not config.has_section('Settings'):
        config.add_section('Settings')

    # Получаем значения из конфига или используем значения по умолчанию
    extensions = config.get('Settings', 'extensions', fallback=','.join(defaults['extensions'])).split(',')
    keywords_file = config.get('Settings', 'keywords_file', fallback=defaults['keywords_file'])
    directory = config.get('Settings', 'directory', fallback=defaults['directory'])
    directories_raw = config.get('Settings', 'directories', fallback=defaults['directories'])
    theme = config.get('Settings', 'theme', fallback=defaults['theme'])
    threads = _safe_getint(config, 'Settings', 'threads', int(defaults['threads']))
    output_file = config.get('Settings', 'output_file', fallback=defaults['output_file'])
    search_images = config.getboolean('Settings', 'search_images', fallback=False)
    max_file_size = _safe_getint(config, 'Settings', 'max_file_size', int(defaults['max_file_size']))
    log_file = config.get('Settings', 'log_file', fallback=defaults['log_file'])
    tesseract_languages = config.get('Settings', 'tesseract_languages', fallback=defaults['tesseract_languages'])
    tesseract_config = config.get('Settings', 'tesseract_config', fallback=defaults['tesseract_config'])
    max_image_size_mb = _safe_getint(config, 'Settings', 'max_image_size_mb', int(defaults['max_image_size_mb']))
    max_pdf_pages = _safe_getint(config, 'Settings', 'max_pdf_pages', int(defaults['max_pdf_pages']))
    max_path_length = _safe_getint(config, 'Settings', 'max_path_length', int(defaults['max_path_length']))
    max_excel_rows_per_sheet = _safe_getint(
        config, 'Settings', 'max_excel_rows_per_sheet', int(defaults['max_excel_rows_per_sheet'])
    )

    # Очищаем значения от пробелов
    extensions = [ext.strip() for ext in extensions]
    keywords_file = keywords_file.strip()
    directory = directory.strip()
    directories = [d.strip() for d in directories_raw.split(';') if d.strip()]
    if not directories:
        directories = [directory] if directory else ['.']
    directory = directories[0]
    theme = theme.strip() or defaults['theme']
    output_file = output_file.strip()
    log_file = log_file.strip()

    # Если поиск по изображениям отключен, убираем изображения из расширений
    if not search_images:
        extensions = [ext for ext in extensions if
                      ext.lower() not in IMAGE_EXTENSION_PATTERNS]

    return {
        'extensions': extensions,
        'keywords_file': keywords_file,
        'directories': directories,
        'directory': directory,
        'theme': theme,
        'threads': threads,
        'output_file': output_file,
        'search_images': search_images,
        'max_file_size': max_file_size,
        'log_file': log_file,
        'tesseract_languages': tesseract_languages,
        'tesseract_config': tesseract_config,
        'max_image_size_mb': max_image_size_mb,
        'max_pdf_pages': max_pdf_pages,
        'max_path_length': max_path_length,
        'max_excel_rows_per_sheet': max_excel_rows_per_sheet,
    }

def create_default_config():
    """Создание файла конфигурации по умолчанию"""
    config_content = """[Settings]
# Расширения файлов для поиска (через запятую)
extensions = *.txt, *.pdf, *.doc, *.docx, *.docm, *.dot, *.dotx, *.dotm, *.xls, *.xlsx, *.xlsm, *.xlt, *.xltx, *.xltm, *.jpg, *.jpeg, *.jpe, *.jfif, *.png, *.bmp, *.gif, *.tif, *.tiff, *.webp, *.ico, *.zip, *.rar, *.7z

# Файл с ключевыми словами (каждое слово с новой строки)
keywords_file = keywords.txt

# Директории для поиска (через ; )
directories = .

# Директория для поиска
directory = .

# Тема интерфейса
theme = Светлая

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

# Максимальный размер изображения для OCR (МБ)
max_image_size_mb = 10

# Максимальное количество страниц PDF для анализа (0 = без ограничения)
max_pdf_pages = 0

# Максимальная длина пути к файлу (символов)
max_path_length = 240

# Максимальное количество строк Excel на лист (0 = без ограничения)
max_excel_rows_per_sheet = 0
"""

    with open("config.txt", "w", encoding="utf-8") as f:
        f.write(config_content)

    print("Создан файл конфигурации config.txt с настройками по умолчанию.")