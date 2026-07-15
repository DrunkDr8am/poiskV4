import os
import fnmatch
import zipfile
import tempfile
import threading
import hashlib
from io import BytesIO
from typing import Set, Dict, List, Optional, Callable
import logging
import re
import json
import subprocess
import sys
import html
import xml.etree.ElementTree as ET

# Глобальные переменные для хранения ключевых слов
# KEYWORDS_LOWER содержит все ключевые слова в нижнем регистре
KEYWORDS_LOWER: Set[str] = set()
# Ключевые слова, которые считаем "словами" (поиск только по целым словам)
KEYWORDS_WORDS: Set[str] = set()
# Ключевые слова, которые ищем как подстроки (короткие/со спецсимволами)
KEYWORDS_SUBSTR: Set[str] = set()

MAX_SAFE_WINDOWS_PATH_LENGTH = 240
MAX_SAFE_WINDOWS_NAME_LENGTH = 180
IMAGE_EXTENSIONS = ('.png', '.jpg', '.jpeg', '.jpe', '.jfif', '.bmp', '.gif', '.tif', '.tiff', '.webp', '.ico')
ARCHIVE_MEMBER_SEP = "::"
WORD_EXTENSIONS = ('.doc', '.docx', '.docm', '.dot', '.dotx', '.dotm')
EXCEL_EXTENSIONS = ('.xls', '.xlsx', '.xlsm', '.xlt', '.xltx', '.xltm')
PDF_SUBPROCESS_TIMEOUT_SEC = 300
PDF_OCR_TESSERACT_CONFIG = '--oem 3 --psm 3'
PDF_OCR_RENDER_ZOOM = 1.5
WORD_PATTERN = re.compile(r"[0-9A-Za-zА-Яа-яЁё]+")
XML_TAG_PATTERN = re.compile(rb"<[^>]+>")
XLSX_EMPTY_ROW_STREAK_LIMIT = 200

_OCR_CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.ocr_cache')
_ocr_memory_cache: Dict[str, str] = {}
_ocr_memory_cache_lock = threading.Lock()
_OCR_MEMORY_CACHE_MAX = 256


def _make_ocr_cache_key(*parts) -> str:
    payload = "|".join(str(part) for part in parts)
    return hashlib.sha256(payload.encode('utf-8', errors='replace')).hexdigest()


def _ocr_disk_cache_path(cache_key: str) -> str:
    return os.path.join(_OCR_CACHE_DIR, f"{cache_key}.txt")


def ocr_cache_get(cache_key: str) -> Optional[str]:
    with _ocr_memory_cache_lock:
        if cache_key in _ocr_memory_cache:
            return _ocr_memory_cache[cache_key]

    cache_path = _ocr_disk_cache_path(cache_key)
    if not os.path.isfile(cache_path):
        return None

    try:
        with open(cache_path, 'r', encoding='utf-8') as cache_file:
            text = cache_file.read()
    except OSError:
        return None

    with _ocr_memory_cache_lock:
        _ocr_memory_cache[cache_key] = text
        if len(_ocr_memory_cache) > _OCR_MEMORY_CACHE_MAX:
            _ocr_memory_cache.pop(next(iter(_ocr_memory_cache)))
    return text


def ocr_cache_put(cache_key: str, text: str) -> None:
    with _ocr_memory_cache_lock:
        _ocr_memory_cache[cache_key] = text
        if len(_ocr_memory_cache) > _OCR_MEMORY_CACHE_MAX:
            _ocr_memory_cache.pop(next(iter(_ocr_memory_cache)))

    try:
        os.makedirs(_OCR_CACHE_DIR, exist_ok=True)
        with open(_ocr_disk_cache_path(cache_key), 'w', encoding='utf-8') as cache_file:
            cache_file.write(text)
    except OSError as exc:
        logging.debug(f"Не удалось сохранить OCR-кэш {cache_key}: {exc}")


def _execute_ocr(cache_key: str, config: dict, ocr_runner: Callable[[], str]) -> str:
    """OCR в том же потоке поиска; отдельного лимита OCR-потоков нет."""
    cached_text = ocr_cache_get(cache_key)
    if cached_text is not None:
        return cached_text

    text = ocr_runner()
    ocr_cache_put(cache_key, text)
    return text


def _file_ocr_cache_key(file_path: str, config: dict) -> str:
    stat = os.stat(file_path)
    return _make_ocr_cache_key(
        os.path.abspath(file_path),
        stat.st_mtime_ns,
        stat.st_size,
        _ocr_backend_tag(),
        config.get('tesseract_languages', 'rus'),
        config.get('tesseract_config', '--oem 3 --psm 6'),
    )


def _pil_image_cache_key(pil_img, config: dict, preprocess: bool) -> str:
    buffer = BytesIO()
    pil_img.save(buffer, format='PNG')
    return _make_ocr_cache_key(
        buffer.getvalue(),
        _ocr_backend_tag(),
        config.get('tesseract_languages', 'rus'),
        config.get('tesseract_config', '--oem 3 --psm 6'),
        preprocess,
    )


def _pdf_page_cache_key(pdf_path: str, page_index: int, xref: Optional[int], config: dict, mode: str, preprocess: bool) -> str:
    stat = os.stat(pdf_path)
    return _make_ocr_cache_key(
        _ocr_backend_tag(),
        os.path.abspath(pdf_path),
        stat.st_mtime_ns,
        stat.st_size,
        page_index,
        xref if xref is not None else mode,
        mode,
        config.get('tesseract_languages', 'rus'),
        config.get('tesseract_config', '--oem 3 --psm 6'),
        preprocess,
        PDF_OCR_RENDER_ZOOM,
    )


def _extract_json_from_stdout(stdout_text: str) -> Dict:
    """Пытается извлечь JSON-объект даже при «шуме» в stdout."""
    raw = (stdout_text or "").strip()
    if not raw:
        return {}
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass

    start_idx = raw.find("{")
    end_idx = raw.rfind("}")
    if start_idx == -1 or end_idx == -1 or end_idx <= start_idx:
        return {}

    candidate = raw[start_idx:end_idx + 1]
    try:
        return json.loads(candidate)
    except json.JSONDecodeError:
        return {}


def _ocr_backend_tag() -> str:
    try:
        from ocr_engine import get_ocr_backend

        return get_ocr_backend() or "none"
    except Exception:
        return "none"


def _ocr_pil_image_cached(
    pil_img,
    config: dict,
    preprocess: bool = True,
    cache_key: Optional[str] = None,
) -> str:
    from ocr_engine import ocr_pil_image

    resolved_cache_key = cache_key or _pil_image_cache_key(pil_img, config, preprocess)

    def run_ocr() -> str:
        return ocr_pil_image(pil_img, config=config, preprocess=preprocess)

    return _execute_ocr(resolved_cache_key, config, run_ocr)


def _extract_embedded_pdf_image(doc, xref: int, Image):
    """Извлекает встроенное изображение PDF с учётом smask (маски прозрачности)."""
    base_image = doc.extract_image(xref)
    if not base_image or "image" not in base_image:
        return None

    pil_img = Image.open(BytesIO(base_image["image"]))
    smask_bytes = base_image.get("smask")
    if not smask_bytes:
        return pil_img

    mask = Image.open(BytesIO(smask_bytes)).convert('L')
    if mask.size != pil_img.size:
        mask = mask.resize(pil_img.size)
    pil_rgba = pil_img.convert('RGBA')
    pil_rgba.putalpha(mask)
    background = Image.new('RGB', pil_rgba.size, (255, 255, 255))
    background.paste(pil_rgba, mask=pil_rgba.split()[-1])
    return background


def _ocr_pdf_page_pixmap(page, fitz_module, Image, config: dict, pdf_path: str, page_index: int) -> str:
    """OCR всей страницы PDF как растрового изображения."""
    from ocr_engine import ocr_pil_image

    cache_key = _pdf_page_cache_key(pdf_path, page_index, None, config, 'pixmap', True)

    def run_ocr() -> str:
        try:
            zoom = PDF_OCR_RENDER_ZOOM
            matrix = fitz_module.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            if pix.width <= 0 or pix.height <= 0:
                return ""
            pil_img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            return ocr_pil_image(pil_img, config=config, preprocess=True)
        except Exception:
            return ""

    return _execute_ocr(cache_key, config, run_ocr)


def _ocr_pdf_scanned_page(page, doc, fitz_module, Image, config: dict, pdf_path: str, page_index: int) -> str:
    """OCR страницы-скана: растеризация страницы даёт лучший результат, чем smask-изображения."""
    page_text = _ocr_pdf_page_pixmap(page, fitz_module, Image, config, pdf_path, page_index)
    if (page_text or "").strip():
        return page_text

    ocr_chunks = []
    for img in page.get_images(full=True):
        try:
            xref = img[0]
            pil_img = _extract_embedded_pdf_image(doc, xref, Image)
            if pil_img is None:
                continue
            cache_key = _pdf_page_cache_key(pdf_path, page_index, xref, config, 'embedded', True)
            embedded_text = _ocr_pil_image_cached(pil_img, config, True, cache_key)
            if embedded_text:
                ocr_chunks.append(embedded_text)
        except Exception:
            continue

    return "\n".join(ocr_chunks)


def _search_in_pdf_core(pdf_path: str, config: dict, keywords_words: Set[str], keywords_substr: Set[str],
                        keywords_all: Set[str]) -> Set[str]:
    """Базовая логика поиска в PDF (используется worker-ом и fallback-режимом)."""
    found = set()
    import fitz  # type: ignore

    with fitz.open(pdf_path) as doc:
        max_pages = int(config.get('max_pdf_pages', 0) or 0)
        has_ocr = bool(config.get('has_ocr', False))
        ocr_ready = False
        Image = None
        if has_ocr:
            try:
                from PIL import Image as _Image  # type: ignore
                from ocr_engine import setup_ocr

                Image = _Image
                ocr_ready = bool(setup_ocr())
            except Exception:
                ocr_ready = False

        for page_index, page in enumerate(doc):
            if max_pages > 0 and page_index >= max_pages:
                break

            text = page.get_text()
            found.update(_search_in_text_worker(text, keywords_words, keywords_substr))
            if keywords_all and found.issuperset(keywords_all):
                break

            if not ocr_ready:
                continue

            page_text_empty = not (text or "").strip()
            if page_text_empty:
                ocr_text = _ocr_pdf_scanned_page(
                    page, doc, fitz, Image, config, pdf_path, page_index
                )
                found.update(_search_in_ocr_text(ocr_text, keywords_words, keywords_substr))
            else:
                for img in page.get_images(full=True):
                    try:
                        xref = img[0]
                        pil_img = _extract_embedded_pdf_image(doc, xref, Image)
                        if pil_img is None:
                            continue
                        cache_key = _pdf_page_cache_key(pdf_path, page_index, xref, config, 'embedded', False)
                        ocr_text = _ocr_pil_image_cached(
                            pil_img, config, False, cache_key
                        )
                        found.update(_search_in_text_worker(ocr_text, keywords_words, keywords_substr))
                        if keywords_all and found.issuperset(keywords_all):
                            break
                    except Exception:
                        continue

            if keywords_all and found.issuperset(keywords_all):
                break

    return found


def _search_in_text_worker(text: str, words: Set[str], substr: Set[str]) -> Set[str]:
    """Локальный поиск для subprocess-воркера PDF."""
    if not text:
        return set()
    text_lower = text.lower()
    found = set()
    if words:
        token_set = set(WORD_PATTERN.findall(text_lower))
        for kw in words:
            if kw in token_set:
                found.add(kw)
    if substr:
        for kw in substr:
            if kw in text_lower:
                found.add(kw)
    return found


def _search_in_ocr_text(text: str, words: Set[str], substr: Set[str]) -> Set[str]:
    """Поиск в OCR-тексте: целые слова + подстроки (для склонений и ошибок распознавания)."""
    found = _search_in_text_worker(text, words, substr)
    if not text:
        return found
    text_lower = text.lower()
    for kw in words:
        if kw not in found and kw in text_lower:
            found.add(kw)
    # Для "рваного" OCR (разрывы внутри слов) проверяем склеенную версию текста.
    compact_text = "".join(WORD_PATTERN.findall(text_lower))
    if compact_text:
        for kw in words:
            if kw not in found and kw in compact_text:
                found.add(kw)
        for kw in substr:
            if kw not in found and kw in compact_text:
                found.add(kw)
    return found


def run_pdf_worker_cli(argv: List[str]) -> int:
    """CLI-воркер обработки PDF. Возвращает код завершения процесса."""
    if len(argv) < 5:
        print(json.dumps({"ok": False, "error": "invalid_arguments"}), end="")
        return 2

    pdf_path = argv[0]
    try:
        config = json.loads(argv[1])
        keywords_words = set(json.loads(argv[2]))
        keywords_substr = set(json.loads(argv[3]))
        keywords_all = set(json.loads(argv[4]))
    except Exception as e:
        print(json.dumps({"ok": False, "error": f"bad_json: {e}"}), end="")
        return 2

    if config.get("has_ocr", False):
        try:
            from ocr_engine import setup_ocr
            setup_ocr()
        except Exception:
            pass

    try:
        import fitz  # type: ignore
    except Exception as e:
        print(json.dumps({"ok": False, "error": f"fitz_import: {e}"}), end="")
        return 2

    try:
        found = _search_in_pdf_core(pdf_path, config, keywords_words, keywords_substr, keywords_all)

        print(json.dumps({"ok": True, "found": sorted(found)}), end="")
        return 0
    except Exception as e:
        print(json.dumps({"ok": False, "error": str(e)}), end="")
        return 1


def get_long_path_reason(file_path: str, config: dict = None) -> str:
    """Возвращает причину пропуска файла при слишком длинном пути на Windows."""
    if os.name != 'nt':
        return ""

    max_path_length = MAX_SAFE_WINDOWS_PATH_LENGTH
    if isinstance(config, dict):
        try:
            max_path_length = int(config.get('max_path_length', MAX_SAFE_WINDOWS_PATH_LENGTH))
        except (TypeError, ValueError):
            max_path_length = MAX_SAFE_WINDOWS_PATH_LENGTH

    # 0 или отрицательное значение означает "без ограничения по длине пути".
    if max_path_length <= 0:
        return ""

    normalized_path = os.path.normpath(file_path)
    if len(normalized_path) > max_path_length:
        return (
            f"слишком длинный путь ({len(normalized_path)} символов, "
            f"лимит {max_path_length})"
        )

    file_name = os.path.basename(normalized_path)
    if len(file_name) > MAX_SAFE_WINDOWS_NAME_LENGTH:
        return (
            f"слишком длинное имя файла ({len(file_name)} символов, "
            f"лимит {MAX_SAFE_WINDOWS_NAME_LENGTH})"
        )

    return ""


def log_long_path_skip(file_path: str, reason: str) -> None:
    """Логирует пропуск файла из-за слишком длинного пути/имени."""
    logging.warning(f"Пропуск файла {file_path}: {reason}")


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
        words_set = set(WORD_PATTERN.findall(text_lower))
        for kw in KEYWORDS_WORDS:
            if kw in words_set:
                found.add(kw)

    # Поиск по подстроке для "особых" ключей
    if KEYWORDS_SUBSTR:
        for kw in KEYWORDS_SUBSTR:
            if kw in text_lower:
                found.add(kw)

    return found


def _keywords_fully_found(found: Set[str]) -> bool:
    """True, если найдены все ключевые слова."""
    return bool(KEYWORDS_LOWER) and found.issuperset(KEYWORDS_LOWER)


def _search_keywords_in_text(text: str, found: Set[str]) -> None:
    """Добавляет в found ключи, встречающиеся в text (без лишних аллокаций)."""
    if not text or not KEYWORDS_LOWER:
        return

    text_lower = text.lower()
    if KEYWORDS_SUBSTR:
        for kw in KEYWORDS_SUBSTR:
            if kw not in found and kw in text_lower:
                found.add(kw)

    remaining_words = KEYWORDS_WORDS - found
    if remaining_words:
        words_set = set(WORD_PATTERN.findall(text_lower))
        for kw in remaining_words:
            if kw in words_set:
                found.add(kw)


def _excel_row_is_empty(row_values) -> bool:
    for cell in row_values:
        if cell is None or cell == "":
            continue
        if isinstance(cell, str):
            if cell.strip():
                return False
        elif str(cell).strip():
            return False
    return True


def _search_keywords_in_excel_row(row_values, found: Set[str]) -> bool:
    """Поиск ключей в одной строке Excel. True — все ключи уже найдены."""
    parts = []
    for cell in row_values:
        if cell is None or cell == "":
            continue
        if isinstance(cell, str):
            text = cell
        else:
            text = str(cell).strip()
        if not text:
            continue

        if KEYWORDS_SUBSTR:
            text_lower = text.lower()
            for kw in KEYWORDS_SUBSTR:
                if kw not in found and kw in text_lower:
                    found.add(kw)
            if _keywords_fully_found(found):
                return True

        parts.append(text)

    if not parts:
        return _keywords_fully_found(found)

    if len(parts) == 1:
        _search_keywords_in_text(parts[0], found)
    else:
        _search_keywords_in_text(" ".join(parts), found)

    return _keywords_fully_found(found)


def _xlsx_local_tag(tag: str) -> str:
    return tag.rsplit('}', 1)[-1]


def _xlsx_si_text(si_elem) -> str:
    """Собирает текст из элемента shared string (si)."""
    parts = []
    for node in si_elem.iter():
        if node.text:
            parts.append(node.text)
        if node is not si_elem and node.tail:
            parts.append(node.tail)
    return ''.join(parts)


def _search_in_xlsx_shared_strings(zip_file: zipfile.ZipFile, found: Set[str]) -> bool:
    """Быстрый поиск по xl/sharedStrings.xml."""
    shared_name = None
    for name in zip_file.namelist():
        if name.lower() == 'xl/sharedstrings.xml':
            shared_name = name
            break
    if not shared_name:
        return False

    with zip_file.open(shared_name) as xml_file:
        for _, elem in ET.iterparse(xml_file, events=('end',)):
            if _xlsx_local_tag(elem.tag) != 'si':
                continue
            text = _xlsx_si_text(elem)
            if text:
                _search_keywords_in_text(text, found)
            elem.clear()
            if _keywords_fully_found(found):
                return True
    return _keywords_fully_found(found)


def _search_in_xlsx_sheet_xml(raw_xml: bytes, found: Set[str]) -> bool:
    """Поиск по XML листа (inline-строки и значения ячеек)."""
    if not raw_xml:
        return False
    text = XML_TAG_PATTERN.sub(b' ', raw_xml).decode('utf-8', errors='ignore')
    text = html.unescape(text)
    if not text.strip():
        return False
    _search_keywords_in_text(text, found)
    return _keywords_fully_found(found)


def _search_in_xlsx_zip_fast(excel_path: str, found: Set[str]) -> bool:
    """Быстрый поиск в XLSX/XLSM через ZIP+XML (без полного обхода ячеек openpyxl)."""
    with zipfile.ZipFile(excel_path, 'r') as zip_file:
        if _search_in_xlsx_shared_strings(zip_file, found):
            return True

        for name in zip_file.namelist():
            lower_name = name.lower()
            if not lower_name.startswith('xl/worksheets/') or not lower_name.endswith('.xml'):
                continue
            raw_xml = zip_file.read(name)
            if _search_in_xlsx_sheet_xml(raw_xml, found):
                return True
    return _keywords_fully_found(found)


def _search_in_xlsx_openpyxl(excel_path: str, found: Set[str]) -> Set[str]:
    """Резервный поиск через openpyxl (медленнее, но точнее для нестандартных файлов)."""
    import openpyxl

    wb = openpyxl.load_workbook(
        excel_path,
        read_only=True,
        data_only=True,
        keep_links=False,
    )
    try:
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            empty_streak = 0
            for row in ws.iter_rows(values_only=True):
                if _excel_row_is_empty(row):
                    empty_streak += 1
                    if empty_streak >= XLSX_EMPTY_ROW_STREAK_LIMIT:
                        break
                    continue
                empty_streak = 0
                if _search_keywords_in_excel_row(row, found):
                    return found
    finally:
        wb.close()
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


def search_in_image(image_data: BytesIO or str, config: dict) -> Set[str]:
    """Распознавание текста с изображения (RapidOCR / Tesseract)."""
    if not config.get('has_ocr', False):
        return set()

    try:
        from PIL import Image
        from ocr_engine import setup_ocr
    except ImportError:
        logging.warning("Модули для OCR (Pillow/ocr_engine) не установлены. Пропуск изображения.")
        return set()

    if not setup_ocr():
        return set()

    try:
        if isinstance(image_data, BytesIO):
            img = Image.open(image_data)
            cache_key = _pil_image_cache_key(img, config, False)
        else:
            if not os.path.isfile(image_data):
                return set()
            cache_key = _file_ocr_cache_key(image_data, config)
            img = Image.open(image_data)

        if img.mode == 'P' and 'transparency' in img.info:
            img = img.convert('RGBA')

        if img.mode == 'RGBA':
            img = img.convert('RGB')
        elif img.mode not in ('RGB', 'L'):
            img = img.convert('RGB')

        text = _ocr_pil_image_cached(img, config, False, cache_key)
        return search_in_text(text)
    except Exception as e:
        logging.error(f"Ошибка обработки изображения {image_data}: {e}")
        return set()


def _hidden_subprocess_kwargs() -> dict:
    """Параметры subprocess без всплывающего окна (Windows)."""
    if os.name != "nt":
        return {}
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = 0  # SW_HIDE
    kwargs = {"startupinfo": startupinfo}
    if hasattr(subprocess, "CREATE_NO_WINDOW"):
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    return kwargs


def _ensure_ocr_for_pdf(config: dict) -> None:
    """Инициализирует OCR перед обработкой PDF в текущем процессе."""
    if not config.get("has_ocr", False):
        return
    try:
        from ocr_engine import setup_ocr
        setup_ocr()
    except Exception:
        pass


def _search_in_pdf_inprocess(pdf_path: str, config: dict) -> Set[str]:
    """Обработка PDF в текущем процессе (без второго окна приложения)."""
    payload_config = {
        "max_pdf_pages": int(config.get("max_pdf_pages", 0) or 0),
        "has_ocr": bool(config.get("has_ocr", False)),
        "tesseract_languages": config.get("tesseract_languages", "rus"),
        "tesseract_config": config.get("tesseract_config", "--oem 3 --psm 6"),
    }
    _ensure_ocr_for_pdf(payload_config)
    try:
        return _search_in_pdf_core(
            pdf_path,
            payload_config,
            KEYWORDS_WORDS,
            KEYWORDS_SUBSTR,
            KEYWORDS_LOWER,
        )
    except Exception as exc:
        logging.error(f"Ошибка обработки PDF {pdf_path}: {exc}")
        return set()


def _search_in_pdf_subprocess(pdf_path: str, config: dict) -> Set[str]:
    """Обработка PDF в отдельном процессе (только через file_processing.py, без GUI)."""
    payload_config = {
        "max_pdf_pages": int(config.get("max_pdf_pages", 0) or 0),
        "has_ocr": bool(config.get("has_ocr", False)),
        "tesseract_languages": config.get("tesseract_languages", "rus"),
        "tesseract_config": config.get("tesseract_config", "--oem 3 --psm 6"),
    }

    worker_args = [
        pdf_path,
        json.dumps(payload_config, ensure_ascii=False),
        json.dumps(sorted(KEYWORDS_WORDS), ensure_ascii=False),
        json.dumps(sorted(KEYWORDS_SUBSTR), ensure_ascii=False),
        json.dumps(sorted(KEYWORDS_LOWER), ensure_ascii=False),
    ]
    command = [sys.executable, os.path.abspath(__file__), "--pdf-worker", *worker_args]

    try:
        proc = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=PDF_SUBPROCESS_TIMEOUT_SEC,
            **_hidden_subprocess_kwargs(),
        )
    except subprocess.TimeoutExpired:
        logging.error(f"Таймаут обработки PDF в subprocess: {pdf_path}")
        return set()
    except Exception as e:
        logging.error(f"Не удалось запустить subprocess для PDF {pdf_path}: {e}")
        return set()

    if proc.returncode != 0:
        stderr_text = (proc.stderr or "").strip()
        stdout_text = (proc.stdout or "").strip()
        details = stderr_text or stdout_text or f"returncode={proc.returncode}"
        logging.error(f"Subprocess обработки PDF завершился с ошибкой для {pdf_path}: {details}")
        return set()

    response = _extract_json_from_stdout(proc.stdout or "")
    if not response:
        stderr_text = (proc.stderr or "").strip()
        stdout_preview = (proc.stdout or "").strip()[:300]
        logging.warning(
            f"Некорректный ответ subprocess для PDF {pdf_path}. "
            f"Пробуем fallback в текущем процессе. stderr={stderr_text or '-'}, stdout={stdout_preview or '-'}"
        )
        return _search_in_pdf_inprocess(pdf_path, config)

    if not response.get("ok", False):
        logging.error(f"Ошибка обработки PDF {pdf_path} в subprocess: {response.get('error', 'unknown')}")
        return set()

    return set(response.get("found", []))


def search_in_pdf(pdf_path: str, config: dict) -> Set[str]:
    """Обработка PDF: в .exe — в текущем процессе (без моргания окна), иначе — subprocess."""
    if getattr(sys, "frozen", False):
        return _search_in_pdf_inprocess(pdf_path, config)
    return _search_in_pdf_subprocess(pdf_path, config)


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
                    if img_file.lower().endswith(IMAGE_EXTENSIONS):
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
    """Обработка Excel файлов с поддержкой старых и новых форматов."""
    found: Set[str] = set()
    if not KEYWORDS_LOWER:
        return found

    try:
        # Пропускаем временные файлы Excel
        if os.path.basename(excel_path).startswith('~$'):
            return found

        file_ext = os.path.splitext(excel_path)[1].lower()

        if file_ext in ('.xlsx', '.xlsm', '.xltx', '.xltm'):
            used_fast_path = False
            if zipfile.is_zipfile(excel_path):
                try:
                    if _search_in_xlsx_zip_fast(excel_path, found):
                        return found
                    used_fast_path = True
                except Exception as fast_error:
                    logging.warning(
                        f"Быстрый поиск в {excel_path} не удался: {fast_error}. "
                        "Используется openpyxl."
                    )

            if not _keywords_fully_found(found):
                try:
                    found = _search_in_xlsx_openpyxl(excel_path, found)
                except ImportError:
                    if not used_fast_path:
                        logging.warning(f"Модуль openpyxl не установлен. Пропуск файла {file_ext}")
                except Exception as openpyxl_error:
                    logging.error(f"Ошибка openpyxl для {excel_path}: {openpyxl_error}")

        elif file_ext in ('.xls', '.xlt'):
            try:
                import xlrd
                workbook = xlrd.open_workbook(excel_path, on_demand=True)
                try:
                    for sheet_index in range(workbook.nsheets):
                        sheet = workbook.sheet_by_index(sheet_index)
                        empty_streak = 0
                        for row_index in range(sheet.nrows):
                            row = sheet.row_values(row_index)
                            if _excel_row_is_empty(row):
                                empty_streak += 1
                                if empty_streak >= XLSX_EMPTY_ROW_STREAK_LIMIT:
                                    break
                                continue
                            empty_streak = 0
                            if _search_keywords_in_excel_row(row, found):
                                return found
                finally:
                    workbook.release_resources()
            except ImportError:
                logging.warning(f"Модуль xlrd не установлен. Пропуск файла {file_ext}")
            except Exception as e:
                logging.error(f"Ошибка обработки Excel {file_ext} {excel_path}: {e}")

        else:
            logging.warning(f"Неизвестное расширение файла Excel: {excel_path}")

    except Exception as e:
        logging.error(f"Ошибка обработки Excel {excel_path}: {e}")

    return found

def format_archive_member_path(archive_path: str, member_path: str) -> str:
    """Полный отображаемый путь: архив + файл внутри архива."""
    normalized_member = member_path.replace("\\", "/")
    return f"{archive_path}{ARCHIVE_MEMBER_SEP}{normalized_member}"


def _record_archive_hits(results: Dict[str, Set[str]], archive_path: str, member_path: str, keywords: Set[str]) -> None:
    if keywords:
        display_path = format_archive_member_path(archive_path, member_path)
        results.setdefault(display_path, set()).update(keywords)


def _archive_member_matches(member_name: str, extensions: List[str]) -> bool:
    normalized = member_name.replace("\\", "/").rstrip("/")
    if not normalized or normalized.endswith("/"):
        return False
    return any(fnmatch.fnmatch(normalized, ext) for ext in extensions)


def _process_extracted_archive_member(
    results: Dict[str, Set[str]],
    archive_path: str,
    member_path: str,
    extracted_file: str,
    config: dict,
) -> None:
    member_lower = member_path.lower()
    if member_lower.endswith(IMAGE_EXTENSIONS) and config.get('has_ocr', False):
        _record_archive_hits(results, archive_path, member_path, search_in_image(extracted_file, config))
    elif member_lower.endswith('.pdf') and config.get('has_pdf', False):
        _record_archive_hits(results, archive_path, member_path, search_in_pdf(extracted_file, config))
    elif member_lower.endswith(WORD_EXTENSIONS) and config.get('has_docx', False):
        _record_archive_hits(results, archive_path, member_path, search_in_docx(extracted_file, config))
    elif member_lower.endswith(EXCEL_EXTENSIONS) and config.get('has_excel', False):
        _record_archive_hits(results, archive_path, member_path, search_in_excel(extracted_file, config))


def _read_7z_member_text(archive, member_name: str) -> str:
    normalized_name = member_name.replace("\\", "/")
    payload = archive.read([normalized_name])
    if not payload:
        return ""
    member_stream = payload.get(normalized_name) or payload.get(member_name)
    if member_stream is None:
        return ""
    raw_data = member_stream.read()
    if isinstance(raw_data, str):
        return raw_data
    return raw_data.decode('utf-8', errors='ignore')


def _process_7z_member(
    archive,
    archive_path: str,
    member_name: str,
    extensions: List[str],
    config: dict,
    temp_dir: str,
    results: Dict[str, Set[str]],
) -> None:
    normalized_name = member_name.replace("\\", "/")
    if not _archive_member_matches(normalized_name, extensions):
        return

    member_lower = normalized_name.lower()
    if member_lower.endswith(('.txt', '.csv', '.log', '.xml', '.html', '.htm')):
        try:
            content = _read_7z_member_text(archive, normalized_name)
            _record_archive_hits(results, archive_path, normalized_name, search_in_text(content))
        except Exception as exc:
            logging.warning(f"Не удалось прочитать {normalized_name} внутри 7Z {archive_path}: {exc}")
        return

    try:
        archive.extract(targets=[normalized_name], path=temp_dir)
    except Exception as exc:
        logging.warning(f"Не удалось извлечь {normalized_name} из 7Z {archive_path}: {exc}")
        return

    extracted_file = os.path.normpath(os.path.join(temp_dir, normalized_name.replace("/", os.sep)))
    if os.path.isfile(extracted_file):
        _process_extracted_archive_member(results, archive_path, normalized_name, extracted_file, config)


def search_in_archive(archive_path: str, extensions: List[str], config: dict) -> Dict[str, Set[str]]:
    """Обработка архивов с поддержкой изображений."""
    results: Dict[str, Set[str]] = {}
    try:
        if archive_path.endswith('.zip'):
            with zipfile.ZipFile(archive_path, 'r') as z:
                with tempfile.TemporaryDirectory() as temp_dir:
                    for file in z.namelist():
                        if any(fnmatch.fnmatch(file, ext) for ext in extensions):
                            if file.lower().endswith(('.txt', '.csv', '.log', '.xml', '.html', '.htm')):
                                try:
                                    with z.open(file) as f:
                                        content = f.read().decode('utf-8', errors='ignore')
                                    _record_archive_hits(results, archive_path, file, search_in_text(content))
                                except Exception as e:
                                    logging.warning(f"Не удалось открыть файл {file} внутри ZIP {archive_path}: {e}")
                                    continue
                            else:
                                try:
                                    z.extract(file, temp_dir)
                                except Exception as e:
                                    logging.warning(f"Не удалось извлечь файл {file} из ZIP {archive_path}: {e}")
                                    continue
                                extracted_file = os.path.join(temp_dir, file)
                                if os.path.isfile(extracted_file):
                                    _process_extracted_archive_member(results, archive_path, file, extracted_file, config)

        elif archive_path.endswith('.7z'):
            try:
                import py7zr
            except ImportError:
                return results
            with tempfile.TemporaryDirectory() as temp_dir:
                try:
                    with py7zr.SevenZipFile(archive_path, mode='r') as archive:
                        for member_name in archive.getnames():
                            _process_7z_member(
                                archive,
                                archive_path,
                                member_name,
                                extensions,
                                config,
                                temp_dir,
                                results,
                            )
                except Exception as e:
                    logging.error(f"Ошибка обработки 7z архива {archive_path}: {e}")

        elif archive_path.endswith('.rar'):
            try:
                import rarfile
            except ImportError:
                return results

            with rarfile.RarFile(archive_path, 'r') as z:
                with tempfile.TemporaryDirectory() as temp_dir:
                    for file in z.namelist():
                        if any(fnmatch.fnmatch(file, ext) for ext in extensions):
                            if file.lower().endswith(('.txt', '.csv', '.log', '.xml', '.html', '.htm')):
                                try:
                                    with z.open(file) as f:
                                        content = f.read().decode('utf-8', errors='ignore')
                                    _record_archive_hits(results, archive_path, file, search_in_text(content))
                                except Exception as e:
                                    logging.warning(f"Не удалось открыть файл {file} внутри RAR {archive_path}: {e}")
                                    continue
                            else:
                                try:
                                    z.extract(file, temp_dir)
                                except Exception as e:
                                    logging.warning(f"Не удалось извлечь файл {file} из RAR {archive_path}: {e}")
                                    continue
                                extracted_file = os.path.join(temp_dir, file)
                                if os.path.isfile(extracted_file):
                                    _process_extracted_archive_member(results, archive_path, file, extracted_file, config)

    except Exception as e:
        logging.error(f"Ошибка обработки архива {archive_path}: {e}")
    return results


def process_file_with_meta(file_path: str, extensions: List[str], max_file_size: int, config: dict):
    """Обработка отдельного файла с мета-статусом.

    Возвращает кортеж: (result_dict, status, error_text, skip_reason),
    где status: matched | no_match | skipped | error.
    """
    found = set()
    status = "no_match"
    skip_reason = ""
    try:
        long_path_reason = get_long_path_reason(file_path, config)
        if long_path_reason:
            log_long_path_skip(file_path, long_path_reason)
            return {}, "skipped", "", "long_path"

        # Проверяем размер файла только если лимит включен (> 0).
        if max_file_size and max_file_size > 0:
            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
            if file_size_mb > max_file_size:
                logging.warning(
                    f"Пропуск файла {file_path} (размер {file_size_mb:.2f} МБ превышает лимит {max_file_size} МБ)")
                return {}, "skipped", "", "large_file"

        ext = os.path.splitext(file_path)[1].lower()

        # Обработка в зависимости от типа файла
        if ext in IMAGE_EXTENSIONS:
            if config.get('has_ocr', False):
                found = search_in_image(file_path, config)
            else:
                logging.info(f"Пропуск изображения {file_path} (OCR недоступен)")
                return {}, "skipped", "", "module_unavailable"
        elif ext == '.pdf':
            if config.get('has_pdf', False):
                found = search_in_pdf(file_path, config)
            else:
                logging.info(f"Пропуск PDF {file_path} (обработка PDF недоступна)")
                return {}, "skipped", "", "module_unavailable"
        elif ext in WORD_EXTENSIONS:
            if config.get('has_docx', False):
                found = search_in_docx(file_path, config)
            else:
                logging.info(f"Пропуск DOCX {file_path} (обработка DOCX недоступна)")
                return {}, "skipped", "", "module_unavailable"
        elif ext in EXCEL_EXTENSIONS:
            if config.get('has_excel', False):
                found = search_in_excel(file_path, config)
            else:
                logging.info(f"Пропуск Excel {file_path} (обработка Excel недоступна)")
                return {}, "skipped", "", "module_unavailable"
        elif ext in ('.zip', '.7z', '.rar'):
            if ext == '.7z' and not config.get('has_7z', False):
                logging.info(f"Пропуск 7Z {file_path} (обработка 7Z недоступна)")
                return {}, "skipped", "", "module_unavailable"
            if ext == '.rar' and not config.get('has_rar', False):
                logging.info(f"Пропуск RAR {file_path} (обработка RAR недоступна)")
                return {}, "skipped", "", "module_unavailable"
            archive_results = search_in_archive(file_path, extensions, config)
            if archive_results:
                return archive_results, "matched", "", ""
            return {}, status, "", ""
        else:
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                    found = search_in_text(content)
            except UnicodeDecodeError:
                encodings = ['cp1251', 'iso-8859-1', 'latin1']
                decoded = False
                for encoding in encodings:
                    try:
                        with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                            content = f.read()
                            found = search_in_text(content)
                            decoded = True
                            break
                    except UnicodeDecodeError:
                        continue
                if not decoded:
                    logging.warning(f"Пропуск файла {file_path} (ошибка чтения текста)")
                    return {}, "skipped", "", "read_error"
            except OSError as read_error:
                logging.warning(f"Пропуск файла {file_path} (ошибка чтения: {read_error})")
                return {}, "skipped", "", "read_error"

        if found:
            status = "matched"
            return {file_path: found}, status, "", ""
        return {}, status, "", ""
    except Exception as e:
        error_text = f"Ошибка обработки файла {file_path}: {e}"
        logging.error(error_text)
        return {}, "error", error_text, ""


def process_file(file_path: str, extensions: List[str], max_file_size: int, config: dict) -> Dict[str, Set[str]]:
    """Обратная совместимость: возвращает только словарь совпадений."""
    result, _, _, _ = process_file_with_meta(file_path, extensions, max_file_size, config)
    return result


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--pdf-worker":
        raise SystemExit(run_pdf_worker_cli(sys.argv[2:]))
