"""Единый OCR-движок: RapidOCR (основной) + Tesseract (fallback)."""

from __future__ import annotations

import logging
import threading
from typing import Optional, Tuple

# Картинки меньше этого размера (по любой стороне) пропускаются — обычно иконки/декор.
OCR_MIN_IMAGE_SIDE = 48

_ocr_ready: Optional[bool] = None
_ocr_backend: str = "none"  # rapidocr_cyrillic | tesseract | none
_rapidocr_local = threading.local()
_tesseract_ready = False
_init_lock = threading.Lock()
_onnx_accel: Optional[Tuple[bool, bool]] = None  # (cuda, dml)


def get_ocr_backend() -> str:
    return _ocr_backend


def is_ocr_available() -> bool:
    return bool(_ocr_ready)


def setup_ocr() -> bool:
    """Инициализирует OCR: сначала RapidOCR (кириллица), иначе Tesseract."""
    global _ocr_ready, _ocr_backend, _tesseract_ready
    with _init_lock:
        if _ocr_ready is not None:
            return _ocr_ready

        if _try_init_rapidocr():
            _ocr_backend = "rapidocr_cyrillic"
            _ocr_ready = True
            logging.info("OCR: используется RapidOCR (кириллица, ONNX Runtime)")
            return True

        try:
            from tesseract_setup import setup_tesseract

            _tesseract_ready = bool(setup_tesseract())
        except Exception as exc:
            logging.warning(f"OCR: не удалось инициализировать Tesseract: {exc}")
            _tesseract_ready = False

        if _tesseract_ready:
            _ocr_backend = "tesseract"
            _ocr_ready = True
            logging.info("OCR: используется Tesseract (fallback)")
            return True

        _ocr_backend = "none"
        _ocr_ready = False
        logging.warning("OCR недоступен: ни RapidOCR, ни Tesseract не инициализированы")
        return False


def _detect_onnx_accelerators() -> Tuple[bool, bool]:
    """Проверяет наличие CUDA / DirectML в onnxruntime и пишет результат в лог."""
    global _onnx_accel
    if _onnx_accel is not None:
        return _onnx_accel

    has_cuda = False
    has_dml = False
    providers = []
    try:
        import onnxruntime as ort  # type: ignore

        providers = list(ort.get_available_providers() or [])
        has_cuda = "CUDAExecutionProvider" in providers
        has_dml = "DmlExecutionProvider" in providers
    except Exception as exc:
        logging.warning(f"OCR GPU: не удалось получить список providers onnxruntime ({exc})")
        _onnx_accel = (False, False)
        return _onnx_accel

    if has_cuda:
        logging.info("OCR GPU: CUDA найден")
    else:
        logging.info("OCR GPU: CUDA отсутствует")

    if has_dml:
        logging.info("OCR GPU: DirectML найден")
    else:
        logging.info("OCR GPU: DirectML отсутствует")

    if not has_cuda and not has_dml:
        logging.info(
            "OCR GPU: ускорение недоступно, используется CPU "
            f"(providers={providers})"
        )
    else:
        logging.info(f"OCR GPU: доступные providers={providers}")

    _onnx_accel = (has_cuda, has_dml)
    return _onnx_accel


def _build_rapidocr_engine():
    """Создаёт RapidOCR с моделью распознавания кириллицы (русский и др.)."""
    from rapidocr import LangRec, ModelType, OCRVersion, RapidOCR  # type: ignore

    has_cuda, has_dml = _detect_onnx_accelerators()
    # Предпочитаем CUDA; DirectML включаем, если CUDA нет.
    use_cuda = bool(has_cuda)
    use_dml = bool(has_dml and not has_cuda)
    if use_cuda:
        logging.info("OCR GPU: RapidOCR будет использовать CUDA")
    elif use_dml:
        logging.info("OCR GPU: RapidOCR будет использовать DirectML")
    else:
        logging.info("OCR GPU: RapidOCR будет использовать CPU")

    return RapidOCR(
        params={
            "Rec.lang_type": LangRec.CYRILLIC,
            "Rec.ocr_version": OCRVersion.PPOCRV5,
            "Rec.model_type": ModelType.MOBILE,
            "EngineConfig.onnxruntime.use_cuda": use_cuda,
            "EngineConfig.onnxruntime.use_dml": use_dml,
        }
    )


def _try_init_rapidocr() -> bool:
    try:
        import numpy as np  # type: ignore
        from PIL import Image  # type: ignore

        engine = _build_rapidocr_engine()
        # Прогрев: убеждаемся, что модели грузятся без падения.
        probe = Image.new("RGB", (64, 32), color="white")
        engine(np.asarray(probe))
        _rapidocr_local.engine = engine
        return True
    except Exception as exc:
        logging.info(f"OCR: RapidOCR недоступен ({exc}), будет проверен Tesseract")
        return False


def _get_rapidocr_engine():
    engine = getattr(_rapidocr_local, "engine", None)
    if engine is not None:
        return engine
    engine = _build_rapidocr_engine()
    _rapidocr_local.engine = engine
    return engine


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


def _rapidocr_to_text(result) -> str:
    if result is None:
        return ""
    txts = getattr(result, "txts", None)
    if txts:
        return "\n".join(str(item) for item in txts if item)
    # Совместимость со старым API rapidocr_onnxruntime: (boxes, txts, scores), elapse
    if isinstance(result, (list, tuple)) and result:
        first = result[0]
        if isinstance(first, (list, tuple)) and first and isinstance(first[0], (list, tuple)):
            # [[box, text, score], ...]
            lines = []
            for item in first:
                if isinstance(item, (list, tuple)) and len(item) >= 2:
                    lines.append(str(item[1]))
            return "\n".join(lines)
        if all(isinstance(item, str) for item in first):
            return "\n".join(first)
    return str(result) if result else ""


def _ocr_with_rapidocr(pil_img) -> str:
    import numpy as np  # type: ignore

    engine = _get_rapidocr_engine()
    arr = np.asarray(pil_img)
    result = engine(arr)
    return _rapidocr_to_text(result)


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
    """Распознаёт текст с PIL-изображения текущим OCR-движком."""
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

    try:
        if _ocr_backend.startswith("rapidocr"):
            return _ocr_with_rapidocr(image)
        if _ocr_backend == "tesseract":
            return _ocr_with_tesseract(
                image,
                config.get("tesseract_languages", "rus"),
                config.get("tesseract_config", "--oem 3 --psm 6"),
            )
    except Exception as exc:
        logging.error(f"Ошибка OCR ({_ocr_backend}): {exc}")
    return ""
