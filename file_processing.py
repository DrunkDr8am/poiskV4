import os
import fnmatch
import zipfile
import tempfile
from io import BytesIO
from typing import Set, Dict, List
import logging
import re

# Глобальные переменные для хранения ключевых слов
# KEYWORDS_LOWER содержит все ключевые слова в нижнем регистре
KEYWORDS_LOWER: Set[str] = set()
# Ключевые слова, которые считаем "словами" (поиск только по целым словам)
KEYWORDS_WORDS: Set[str] = set()
# Ключевые слова, которые ищем как подстроки (короткие/со спецсимволами)
KEYWORDS_SUBSTR: Set[str] = set()
MAX_SAFE_WINDOWS_PATH_LENGTH = 240


def load_keywords(keywords_file: str) -> List[str]:
    """Загрузка ключевых слов из файла с проверкой кодировки

    Формирует три набора:
    - KEYWORDS_LOWER: все ключевые слова в нижнем регистре
    - KEYWORDS_WORDS: "словесные" ключи (буквы/цифры, длина >= 3) для поиска по целым словам
    - KEYWORDS_SUBSTR: короткие и/или содержащие спецсимволы, ищутся как подстроки
    """
    global KEYWORDS_LOWER, KEYWORDS_WORDS, KEYWORDS_SUBSTR
    encodings = ['utf-8', 'cp1251', 'iso-8859-1', 'utf-8-sig']
    for encoding in encodings:
        try:
            with open(keywords_file, 'r', encoding=encoding) as f:
                keywords = [line.strip() for line in f if line.strip()]
                if keywords:
                    # Сохраняем ключевые слова в нижнем регистре для быстрого поиска
                    KEYWORDS_LOWER = {kw.lower() for kw in keywords}

                    word_like: Set[str] = set()
                    substr_like: Set[str] = set()

                    for kw in KEYWORDS_LOWER:
                        # "Словесным" считаем ключ, состоящий только из букв/цифр и длиной >= 3
                        if re.fullmatch(r"[0-9A-Za-zА-Яа-яЁё]+", kw) and len(kw) >= 3:
                            word_like.add(kw)
                        else:
                            substr_like.add(kw)

                    KEYWORDS_WORDS = word_like
                    KEYWORDS_SUBSTR = substr_like
                    return keywords
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Не удалось декодировать файл {keywords_file} с поддержанными кодировками: {encodings}")


def search_in_text(text: str) -> Set[str]:
    """Поиск ключевых слов в тексте.

    Для "словесных" ключей (KEYWORDS_WORDS) выполняется поиск по целым словам.
    Для коротких/специальных ключей (KEYWORDS_SUBSTR) используется поиск по подстроке.
    """
    if not text:
        return set()

    text_lower = text.lower()

    found: Set[str] = set()

    # Поиск по целым словам
    if KEYWORDS_WORDS:
        # Выделяем слова: последовательности букв/цифр (рус/англ)
        words = re.findall(r"[0-9A-Za-zА-Яа-яЁё]+", text_lower)
        words_set = set(words)
        for kw in KEYWORDS_WORDS:
            if kw in words_set:
                found.add(kw)

    # Поиск по подстроке для "особых" ключей
    if KEYWORDS_SUBSTR:
        for kw in KEYWORDS_SUBSTR:
            if kw in text_lower:
                found.add(kw)

    return found


def extract_best_effort_text(file_path: str) -> str:
    """Пытается извлечь читаемый текст из файла даже при некорректном формате."""
    try:
        with open(file_path, 'rb') as f:
            raw = f.read()
    except Exception:
        return ""

    candidates = []
    for encoding in ('utf-8', 'cp1251', 'utf-16le', 'latin1'):
        try:
            decoded = raw.decode(encoding, errors='ignore')
            if decoded:
                candidates.append(decoded)
        except Exception:
            continue

    if not candidates:
        return ""

    # Склеиваем варианты декодирования и немного очищаем шум
    merged_text = "\n".join(candidates)
    merged_text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]+', ' ', merged_text)
    merged_text = re.sub(r'\s+', ' ', merged_text)
    return merged_text


def is_path_too_long(file_path: str) -> bool:
    """Проверяет, что путь превышает безопасную длину для Windows-библиотек."""
    normalized_path = os.path.abspath(file_path)
    if os.name != 'nt':
        return False
    return len(normalized_path) > MAX_SAFE_WINDOWS_PATH_LENGTH


def search_in_image(image_data: BytesIO or str, config: dict) -> Set[str]:
    """Распознавание текста с изображения"""
    # Проверяем доступность OCR через конфиг
    if not config.get('has_ocr', False):
        return set()

    try:
        from PIL import Image
        import pytesseract
    except ImportError:
        logging.warning("Модули для OCR (Pillow/pytesseract) не установлены. Пропуск изображения.")
        return set()

    try:
        img = Image.open(image_data) if isinstance(image_data, BytesIO) else Image.open(image_data)

        # Для палитровых изображений с прозрачностью сначала переводим в RGBA,
        # чтобы избежать предупреждения Pillow и корректно обработать альфа-канал.
        if img.mode == 'P' and 'transparency' in img.info:
            img = img.convert('RGBA')

        if img.mode == 'RGBA':
            img = img.convert('RGB')
        elif img.mode not in ('RGB', 'L'):
            img = img.convert('RGB')

        # Используем настройки из конфига
        languages = config.get('tesseract_languages', 'rus')
        config_param = config.get('tesseract_config', '--oem 3 --psm 6')

        text = pytesseract.image_to_string(img, lang=languages, config=config_param)
        return search_in_text(text)
    except Exception as e:
        logging.error(f"Ошибка обработки изображения {image_data}: {e}")
        return set()


def search_in_pdf(pdf_path: str, config: dict) -> Set[str]:
    """Обработка PDF файлов"""
    try:
        import fitz  # PyMuPDF
    except ImportError:
        return set()

    found = set()
    try:
        with fitz.open(pdf_path) as doc:
            max_pages = int(config.get('max_pdf_pages', 0) or 0)
            for page_index, page in enumerate(doc):
                # 0 означает "без ограничения"
                if max_pages > 0 and page_index >= max_pages:
                    logging.info(
                        f"Пропуск оставшихся страниц PDF {pdf_path} (достигнут лимит {max_pages} страниц)"
                    )
                    break
                # Текст со страницы
                text = page.get_text()
                found.update(search_in_text(text))
                if KEYWORDS_LOWER and found.issuperset(KEYWORDS_LOWER):
                    logging.info(f"Ранний останов PDF {pdf_path} (найдены все ключевые слова)")
                    return found

                # Обработка изображений (только если есть OCR)
                for img in page.get_images(full=True):
                    xref = img[0]
                    base_image = doc.extract_image(xref)
                    if base_image and "image" in base_image:
                        image_data = BytesIO(base_image["image"])
                        found.update(search_in_image(image_data, config))
                        if KEYWORDS_LOWER and found.issuperset(KEYWORDS_LOWER):
                            logging.info(f"Ранний останов PDF {pdf_path} (найдены все ключевые слова)")
                            return found
    except Exception as e:
        logging.error(f"Ошибка обработки PDF {pdf_path}: {e}")
    return found


def search_in_docx(docx_path: str, config: dict) -> Set[str]:
    """Обработка DOCX файлов"""
    try:
        import docx2txt
    except ImportError:
        docx2txt = None

    found = set()
    try:
        base_name = os.path.basename(docx_path)
        file_ext = os.path.splitext(docx_path)[1].lower()

        # Для временных файлов Office, старых .doc и "битых" .docx
        # используем fallback-поиск по извлекаемому тексту.
        use_fallback = (
            base_name.startswith('~$')
            or file_ext == '.doc'
            or not zipfile.is_zipfile(docx_path)
            or docx2txt is None
        )

        if use_fallback:
            fallback_text = extract_best_effort_text(docx_path)
            if fallback_text:
                return search_in_text(fallback_text)
            return set()

        # Текст из документа
        text = docx2txt.process(docx_path)
        found.update(search_in_text(text))

        # Изображения из документа (только если есть OCR)
        if config.get('has_ocr', False):
            with tempfile.TemporaryDirectory() as temp_dir:
                docx2txt.process(docx_path, temp_dir)
                for img_file in os.listdir(temp_dir):
                    if img_file.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff')):
                        img_path = os.path.join(temp_dir, img_file)
                        found.update(search_in_image(img_path, config))
    except Exception as e:
        # Если штатная обработка не удалась, пытаемся хотя бы извлечь текст напрямую.
        logging.warning(f"Ошибка стандартной обработки DOCX {docx_path}: {e}. Переход к fallback-обработке.")
        fallback_text = extract_best_effort_text(docx_path)
        if fallback_text:
            found.update(search_in_text(fallback_text))
    return found


def search_in_excel(excel_path: str, config: dict) -> Set[str]:
    """Обработка Excel файлов с поддержкой старых форматов .xls."""
    found = set()
    try:
        # Пропускаем временные файлы Excel
        if os.path.basename(excel_path).startswith('~$'):
            return set()

        # Определяем расширение файла
        file_ext = os.path.splitext(excel_path)[1].lower()

        if file_ext == '.xlsx':
            # Для новых форматов используем openpyxl
            try:
                import openpyxl
                wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
                for sheet in wb.sheetnames:
                    ws = wb[sheet]
                    for row_index, row in enumerate(ws.iter_rows(values_only=True), start=1):
                        for cell in row:
                            if cell and isinstance(cell, str):
                                found.update(search_in_text(cell))
            except ImportError:
                logging.warning("Модуль openpyxl не установлен. Пропуск файла .xlsx")

        elif file_ext == '.xls':
            # Для старых форматов используем xlrd
            try:
                import xlrd
                workbook = xlrd.open_workbook(excel_path)
                for sheet_index in range(workbook.nsheets):
                    sheet = workbook.sheet_by_index(sheet_index)
                    for row_index in range(sheet.nrows):
                        for col_index in range(sheet.ncols):
                            cell_value = sheet.cell_value(row_index, col_index)
                            if cell_value and isinstance(cell_value, str):
                                found.update(search_in_text(str(cell_value)))
            except ImportError:
                logging.warning("Модуль xlrd не установлен. Пропуск файла .xls")
            except Exception as e:
                logging.error(f"Ошибка обработки Excel .xls {excel_path}: {e}")

        else:
            logging.warning(f"Неизвестное расширение файла Excel: {excel_path}")

    except Exception as e:
        logging.error(f"Ошибка обработки Excel {excel_path}: {e}")

    return found

def search_in_archive(archive_path: str, extensions: List[str], config: dict) -> Set[str]:
    """Обработка архивов с поддержкой изображений"""
    found = set()
    try:
        if archive_path.endswith('.zip'):
            with zipfile.ZipFile(archive_path, 'r') as z:
                with tempfile.TemporaryDirectory() as temp_dir:
                    for file in z.namelist():
                        if any(fnmatch.fnmatch(file, ext) for ext in extensions):
                            # Для текстовых файлов читаем напрямую
                            if file.lower().endswith(('.txt', '.csv', '.log', '.xml', '.html', '.htm')):
                                with z.open(file) as f:
                                    content = f.read().decode('utf-8', errors='ignore')
                                    found.update(search_in_text(content))
                            else:
                                # Извлекаем файл во временную директорию один раз
                                z.extract(file, temp_dir)
                                extracted_file = os.path.join(temp_dir, file)
                                if os.path.isfile(extracted_file):
                                    # Обрабатываем изображения (только если OCR доступен)
                                    if file.lower().endswith(
                                            ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff')) and config.get(
                                        'has_ocr', False):
                                        found.update(search_in_image(extracted_file, config))
                                    # Обрабатываем PDF
                                    elif file.lower().endswith('.pdf') and config.get('has_pdf', False):
                                        found.update(search_in_pdf(extracted_file, config))
                                    # Обрабатываем DOCX
                                    elif file.lower().endswith(('.docx', '.doc')) and config.get('has_docx', False):
                                        found.update(search_in_docx(extracted_file, config))
                                    # Обрабатываем Excel
                                elif file.lower().endswith(('.xls', '.xlsx')) and config.get('has_excel', False):
                                        found.update(search_in_excel(extracted_file, config))


        elif archive_path.endswith('.7z'):
            try:
                import py7zr
            except ImportError:
                return set()
            with tempfile.TemporaryDirectory() as temp_dir:
                try:
                    with py7zr.SevenZipFile(archive_path, mode='r') as z:
                        # Извлекаем все файлы
                        z.extractall(path=temp_dir)
                    # Рекурсивно обходим извлеченные файлы
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            relative_path = os.path.relpath(file_path, temp_dir)
                            if any(fnmatch.fnmatch(relative_path, ext) for ext in extensions):
                                # Обрабатываем файлы в зависимости от типа
                                if file.lower().endswith(('.txt', '.csv', '.log', '.xml', '.html', '.htm')):
                                    try:
                                        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                                            content = f.read()
                                            found.update(search_in_text(content))
                                    except:
                                        try:
                                            with open(file_path, 'rb') as f:
                                                content = f.read().decode('utf-8', errors='ignore')
                                                found.update(search_in_text(content))
                                        except:
                                            pass
                                elif file.lower().endswith(
                                        ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff')) and config.get('has_ocr',
                                                                                                          False):
                                    found.update(search_in_image(file_path, config))
                                elif file.lower().endswith('.pdf') and config.get('has_pdf', False):
                                    found.update(search_in_pdf(file_path, config))
                                elif file.lower().endswith(('.docx', '.doc')) and config.get('has_docx', False):
                                    found.update(search_in_docx(file_path, config))
                                elif file.lower().endswith(('.xls', '.xlsx')) and config.get('has_excel', False):
                                    found.update(search_in_excel(file_path, config))
                except Exception as e:
                    logging.error(f"Ошибка обработки 7z архива {archive_path}: {e}")

        elif archive_path.endswith('.rar'):
            try:
                import rarfile
            except ImportError:
                return set()

            with rarfile.RarFile(archive_path, 'r') as z:
                with tempfile.TemporaryDirectory() as temp_dir:
                    for file in z.namelist():
                        if any(fnmatch.fnmatch(file, ext) for ext in extensions):
                            # Для текстовых файлов читаем напрямую
                            if file.lower().endswith(('.txt', '.csv', '.log', '.xml', '.html', '.htm')):
                                with z.open(file) as f:
                                    content = f.read().decode('utf-8', errors='ignore')
                                    found.update(search_in_text(content))
                            else:
                                # Извлекаем файл во временную директорию один раз
                                z.extract(file, temp_dir)
                                extracted_file = os.path.join(temp_dir, file)
                                if os.path.isfile(extracted_file):
                                    # Обрабатываем изображения (только если OCR доступен)
                                    if file.lower().endswith(
                                            ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff')) and config.get(
                                        'has_ocr', False):
                                        found.update(search_in_image(extracted_file, config))
                                    # Обрабатываем PDF
                                    elif file.lower().endswith('.pdf') and config.get('has_pdf', False):
                                        found.update(search_in_pdf(extracted_file, config))
                                    # Обрабатываем DOCX
                                    elif file.lower().endswith(('.docx', '.doc')) and config.get('has_docx', False):
                                        found.update(search_in_docx(extracted_file, config))
                                    # Обрабатываем Excel
                                    elif file.lower().endswith(('.xls', '.xlsx')) and config.get('has_excel', False):
                                        found.update(search_in_excel(extracted_file, config))

    except Exception as e:
        logging.error(f"Ошибка обработки архива {archive_path}: {e}")
    return found


def process_file(file_path: str, extensions: List[str], max_file_size: int, config: dict) -> Dict[str, Set[str]]:
    """Обработка отдельного файла"""
    found = set()
    try:
        normalized_path = os.path.abspath(file_path)
        if is_path_too_long(normalized_path):
            logging.warning(
                "Пропуск файла %s (длина пути %d символов превышает безопасный лимит %d)",
                normalized_path,
                len(normalized_path),
                MAX_SAFE_WINDOWS_PATH_LENGTH
            )
            return {}

        # Проверяем размер файла
        file_size_mb = os.path.getsize(normalized_path) / (1024 * 1024)
        if file_size_mb > max_file_size:
            logging.warning(
                f"Пропуск файла {normalized_path} (размер {file_size_mb:.2f} МБ превышает лимит {max_file_size} МБ)")
            return {}

        ext = os.path.splitext(normalized_path)[1].lower()

        # Обработка в зависимости от типа файла
        if ext in ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff'):
            # Проверяем доступность OCR через конфиг
            if config.get('has_ocr', False):
                found = search_in_image(normalized_path, config)
            else:
                logging.info(f"Пропуск изображения {normalized_path} (OCR недоступен)")
        elif ext == '.pdf':
            # Проверяем доступность обработки PDF
            if config.get('has_pdf', False):
                found = search_in_pdf(normalized_path, config)
            else:
                logging.info(f"Пропуск PDF {normalized_path} (обработка PDF недоступна)")
        elif ext in ('.doc', '.docx'):
            # Проверяем доступность обработки DOCX
            if config.get('has_docx', False):
                found = search_in_docx(normalized_path, config)
            else:
                logging.info(f"Пропуск DOCX {normalized_path} (обработка DOCX недоступна)")
        elif ext in ('.xls', '.xlsx'):
            # Проверяем доступность обработки Excel
            if config.get('has_excel', False):
                found = search_in_excel(normalized_path, config)
            else:
                logging.info(f"Пропуск Excel {normalized_path} (обработка Excel недоступна)")
        elif ext in ('.zip', '.7z', '.rar'):
            # Для архивов проверяем доступность соответствующих модулей
            if ext == '.7z' and not config.get('has_7z', False):
                logging.info(f"Пропуск 7Z {normalized_path} (обработка 7Z недоступна)")
            elif ext == '.rar' and not config.get('has_rar', False):
                logging.info(f"Пропуск RAR {normalized_path} (обработка RAR недоступна)")
            else:
                found = search_in_archive(normalized_path, extensions, config)  # Передаем config
        else:
            # Обработка текстовых файлов
            try:
                with open(normalized_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                    found = search_in_text(content)
            except UnicodeDecodeError:
                encodings = ['cp1251', 'iso-8859-1', 'latin1']
                for encoding in encodings:
                    try:
                        with open(normalized_path, 'r', encoding=encoding, errors='ignore') as f:
                            content = f.read()
                            found = search_in_text(content)
                            break
                    except UnicodeDecodeError:
                        continue

        return {normalized_path: found} if found else {}
    except Exception as e:
        logging.error(f"Ошибка обработки файла {file_path}: {e}")
        return {}
