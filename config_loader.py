import os
import configparser
import hashlib
import hmac


ADMIN_AUTH_FILE = "admin_auth.txt"
HASH_SCHEME_VERSION = "v2"
ADMIN_ROLE = "admin"
USER_ROLE = "user"


def hash_password(password):
    """Возвращает legacy SHA-256 хэш пароля без привязки к роли."""
    return hashlib.sha256((password or "").encode("utf-8")).hexdigest()


def make_role_password_hash(password, role):
    """Возвращает ролевой хэш пароля, который нельзя переиспользовать для другой роли."""
    password_bytes = f"zsearch::{role}::{password or ''}".encode("utf-8")
    digest = hashlib.sha256(password_bytes).hexdigest()
    return f"{HASH_SCHEME_VERSION}:{role}:{digest}"


def verify_password_for_role(password, stored_hash, role, conflicting_hash=""):
    """Проверяет пароль для конкретной роли, поддерживая новый и legacy форматы."""
    stored_hash = str(stored_hash or "").strip()
    conflicting_hash = str(conflicting_hash or "").strip()
    if not stored_hash:
        return False

    expected_prefix = f"{HASH_SCHEME_VERSION}:{role}:"
    if stored_hash.startswith(expected_prefix):
        return hmac.compare_digest(make_role_password_hash(password, role), stored_hash)

    if stored_hash.startswith(f"{HASH_SCHEME_VERSION}:"):
        return False

    legacy_hash = hash_password(password)
    if not hmac.compare_digest(legacy_hash, stored_hash):
        return False

    if conflicting_hash and hmac.compare_digest(stored_hash, conflicting_hash):
        return False

    return True


DEFAULT_ADMIN_PASSWORD_HASH = make_role_password_hash("2407", ADMIN_ROLE)
DEFAULT_USER_PASSWORD_HASH = make_role_password_hash("1234", USER_ROLE)


def save_admin_password_hash(password_hash, auth_file=ADMIN_AUTH_FILE):
    """Сохраняет хэш пароля администратора в отдельный файл."""
    auth_config = configparser.ConfigParser()
    auth_config["AdminAuth"] = {
        "admin_password_hash": (password_hash or DEFAULT_ADMIN_PASSWORD_HASH).strip()
    }
    with open(auth_file, "w", encoding="utf-8") as authfile:
        auth_config.write(authfile)


def load_admin_password_hash(auth_file=ADMIN_AUTH_FILE, legacy_config_file="config.txt"):
    """Загружает хэш администратора из отдельного файла с миграцией из старого config.txt."""
    auth_config = configparser.ConfigParser()
    if os.path.exists(auth_file):
        try:
            auth_config.read(auth_file, encoding="utf-8")
            stored_hash = auth_config.get("AdminAuth", "admin_password_hash", fallback="").strip()
            if stored_hash:
                return stored_hash
        except Exception:
            pass

    legacy_hash = ""
    if os.path.exists(legacy_config_file):
        legacy_config = configparser.ConfigParser()
        try:
            legacy_config.read(legacy_config_file, encoding="utf-8")
            legacy_hash = legacy_config.get("Settings", "admin_password_hash", fallback="").strip()
        except Exception:
            legacy_hash = ""

    effective_hash = legacy_hash or DEFAULT_ADMIN_PASSWORD_HASH
    save_admin_password_hash(effective_hash, auth_file)
    return effective_hash

def load_config(config_file="config.txt"):
    """Загрузка конфигурации из файла"""
    config = configparser.ConfigParser()

    # Значения по умолчанию
    defaults = {
        'extensions': ['*.txt', '*.pdf', '*.docx', '*.xlsx', '*.jpg', '*.png', '*.zip', '*.rar', '*.7z'],
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
        'user_password_hash': DEFAULT_USER_PASSWORD_HASH,
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
    directories_raw = config.get('Settings', 'directories', fallback=defaults['directories'])
    theme = config.get('Settings', 'theme', fallback=defaults['theme'])
    threads = config.getint('Settings', 'threads', fallback=int(defaults['threads']))
    output_file = config.get('Settings', 'output_file', fallback=defaults['output_file'])
    search_images = config.getboolean('Settings', 'search_images', fallback=False)
    max_file_size = config.getint('Settings', 'max_file_size', fallback=int(defaults['max_file_size']))
    log_file = config.get('Settings', 'log_file', fallback=defaults['log_file'])
    tesseract_languages = config.get('Settings', 'tesseract_languages', fallback=defaults['tesseract_languages'])
    tesseract_config = config.get('Settings', 'tesseract_config', fallback=defaults['tesseract_config'])
    max_image_size_mb = config.getint('Settings', 'max_image_size_mb', fallback=int(defaults['max_image_size_mb']))
    max_pdf_pages = config.getint('Settings', 'max_pdf_pages', fallback=int(defaults['max_pdf_pages']))
    max_excel_rows_per_sheet = config.getint(
        'Settings',
        'max_excel_rows_per_sheet',
        fallback=int(defaults['max_excel_rows_per_sheet']),
    )
    user_password_hash = config.get('Settings', 'user_password_hash', fallback=defaults['user_password_hash']).strip()

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
                      not ext.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff'))]

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
        'max_excel_rows_per_sheet': max_excel_rows_per_sheet,
        'admin_password_hash': load_admin_password_hash(),
        'user_password_hash': user_password_hash,
    }

def create_default_config():
    """Создание файла конфигурации по умолчанию"""
    config_content = """[Settings]
# Расширения файлов для поиска (через запятую)
extensions = *.txt, *.pdf, *.docx, *.xlsx, *.jpg, *.png, *.zip, *.rar, *.7z

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

# Максимальное количество строк Excel на лист (0 = без ограничения)
max_excel_rows_per_sheet = 0

# SHA-256 хэш пароля пользователя (по умолчанию пароль: 1234)
user_password_hash = 03ac674216f3e15c761ee1a5e255f067953623c8b388b4459e13f978d7c846f4
"""

    with open("config.txt", "w", encoding="utf-8") as f:
        f.write(config_content)
    save_admin_password_hash(DEFAULT_ADMIN_PASSWORD_HASH)

    print("Создан файл конфигурации config.txt с настройками по умолчанию.")