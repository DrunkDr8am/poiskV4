"""OCR-движок на базе Tesseract."""

from __future__ import annotations

import logging
import threading
from collections import OrderedDict
from hashlib import blake2b
from typing import Optional

# Картинки меньше этого размера (по любой стороне) пропускаются — обычно иконки/декор.
OCR_MIN_IMAGE_SIDE = 48

_ocr_ready: Optional[bool] = None
_ocr_backend: str = "none"  # tesseract | none
_init_lock = threading.Lock()
_ocr_cache_lock = threading.Lock()
_ocr_text_cache: "OrderedDict[tuple, str]" = OrderedDict()
_ocr_in_progress: dict[tuple, threading.Event] = {}
_OCR_CACHE_MAX_ITEMS = 256


def get_ocr_backend() -> str:
    return _ocr_backend


def is_ocr_available() -> bool:
    return bool(_ocr_ready)


def setup_ocr() -> bool:
    """Инициализирует Tesseract OCR."""
    global _ocr_ready, _ocr_backend
    with _init_lock:
        if _ocr_ready is not None:
            return _ocr_ready

        try:
            from tesseract_setup import setup_tesseract

            if setup_tesseract():
                _ocr_backend = "tesseract"
                _ocr_ready = True
                logging.info("OCR: используется Tesseract")
                return True
        except Exception as exc:
            logging.warning(f"OCR: не удалось инициализировать Tesseract: {exc}")

        _ocr_backend = "none"
        _ocr_ready = False
        logging.warning("OCR недоступен: Tesseract не инициализирован")
        return False


def _normalize_pil(pil_img, preprocess: bool):
    from PIL import ImageOps  # type: ignore

    if pil_img.mode == "P" and "transparency" in pil_img.info:
        pil_img = pil_img.convert("RGBA")
    if pil_img.mode == "RGBA":
        pil_img = pil_img.convert("RGB")
    elif pil_img.mode not in ("RGB", "L"):
        pil_img = pil_img.convert("RGB")

    if preprocess:
        gray = pil_img.convert("L")
        return ImageOps.autocontrast(gray).convert("RGB")
    if pil_img.mode == "L":
        return pil_img.convert("RGB")
    return pil_img


def _is_image_too_small(pil_img) -> bool:
    try:
        width, height = pil_img.size
    except Exception:
        return False
    return width < OCR_MIN_IMAGE_SIDE or height < OCR_MIN_IMAGE_SIDE


def _ocr_with_tesseract(pil_img, ocr_lang: str, ocr_cfg: str) -> str:
    import pytesseract  # type: ignore

    lang = (ocr_lang or "rus").strip() or "rus"
    parts = [part.strip() for part in lang.split("+") if part.strip()]
    if "rus" in parts and "eng" not in parts:
        parts.append("eng")
    lang = "+".join(parts) if parts else "rus+eng"
    return pytesseract.image_to_string(
        pil_img,
        lang=lang,
        config=ocr_cfg or "--oem 3 --psm 6",
    )


def ocr_pil_image(
    pil_img,
    config: Optional[dict] = None,
    preprocess: bool = True,
) -> str:
    """Распознаёт текст с PIL-изображения через Tesseract."""
    if pil_img is None:
        return ""
    if _is_image_too_small(pil_img):
        logging.debug(
            f"OCR: пропуск маленького изображения {getattr(pil_img, 'size', '?')} "
            f"(мин. сторона {OCR_MIN_IMAGE_SIDE}px)"
        )
        return ""

    if not setup_ocr():
        return ""

    config = config or {}
    image = _normalize_pil(pil_img, preprocess=preprocess)
    ocr_lang = config.get("tesseract_languages", "rus")
    ocr_cfg = config.get("tesseract_config", "--oem 3 --psm 6")

    try:
        raw_pixels = image.tobytes()
        pixel_digest = blake2b(raw_pixels, digest_size=16).digest()
        del raw_pixels
        cache_key = (
            _ocr_backend,
            image.mode,
            image.size,
            bool(preprocess),
            ocr_lang,
            ocr_cfg,
            pixel_digest,
        )
    except Exception:
        cache_key = None

    owner = True
    wait_event = None
    if cache_key is not None:
        with _ocr_cache_lock:
            cached = _ocr_text_cache.get(cache_key)
            if cached is not None:
                _ocr_text_cache.move_to_end(cache_key)
                return cached
            wait_event = _ocr_in_progress.get(cache_key)
            if wait_event is None:
                wait_event = threading.Event()
                _ocr_in_progress[cache_key] = wait_event
            else:
                owner = False

    if not owner and wait_event is not None:
        wait_event.wait()
        with _ocr_cache_lock:
            return _ocr_text_cache.get(cache_key, "")

    text = ""
    try:
        text = _ocr_with_tesseract(image, ocr_lang, ocr_cfg)
    except Exception as exc:
        logging.error(f"Ошибка OCR (Tesseract): {exc}")
    finally:
        if cache_key is not None:
            with _ocr_cache_lock:
                _ocr_text_cache[cache_key] = text
                _ocr_text_cache.move_to_end(cache_key)
                while len(_ocr_text_cache) > _OCR_CACHE_MAX_ITEMS:
                    _ocr_text_cache.popitem(last=False)
                event = _ocr_in_progress.pop(cache_key, None)
                if event is not None:
                    event.set()
    return text
