import os
import sys
import logging
import subprocess

_TESSERACT_READY = None


def _hidden_subprocess_kwargs():
    """Параметры subprocess для гарантированно скрытого окна на Windows."""
    if os.name != "nt":
        return {}
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = 0
    kwargs = {"startupinfo": startupinfo}
    if hasattr(subprocess, "CREATE_NO_WINDOW"):
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    return kwargs


def _patch_pytesseract_no_window(pytesseract_module):
    """
    На Windows принудительно скрывает окно процесса tesseract.exe.
    На Win11 одного STARTUPINFO иногда недостаточно, поэтому добавляем CREATE_NO_WINDOW.
    """
    if os.name != "nt":
        return

    pt = pytesseract_module.pytesseract
    original_subprocess_args = getattr(pt, "subprocess_args", None)
    if not callable(original_subprocess_args):
        return

    # Чтобы не оборачивать повторно при повторных вызовах setup_tesseract().
    if getattr(pt, "_zsearch_no_window_patch", False):
        return

    hidden_kwargs = _hidden_subprocess_kwargs()

    def _wrapped_subprocess_args(include_stdout=True):
        kwargs = original_subprocess_args(include_stdout=include_stdout)
        kwargs.update(hidden_kwargs)
        if "creationflags" in hidden_kwargs:
            kwargs["creationflags"] = kwargs.get("creationflags", 0) | hidden_kwargs["creationflags"]
        return kwargs

    pt.subprocess_args = _wrapped_subprocess_args

    # Дополнительно патчим проверку версии, т.к. там используется check_output
    # без subprocess_args и на Win11 может мигать окно.
    original_get_version = getattr(pt, "get_tesseract_version", None)
    if callable(original_get_version):
        def _wrapped_get_tesseract_version(*args, **kwargs):
            original_check_output = pt.subprocess.check_output
            try:
                def _wrapped_check_output(*c_args, **c_kwargs):
                    merged = dict(c_kwargs)
                    merged.update(hidden_kwargs)
                    if "creationflags" in c_kwargs and "creationflags" in hidden_kwargs:
                        merged["creationflags"] = c_kwargs.get("creationflags", 0) | hidden_kwargs["creationflags"]
                    return original_check_output(*c_args, **merged)

                pt.subprocess.check_output = _wrapped_check_output
                return original_get_version(*args, **kwargs)
            finally:
                pt.subprocess.check_output = original_check_output

        pt.get_tesseract_version = _wrapped_get_tesseract_version

    pt._zsearch_no_window_patch = True


def setup_tesseract():
    """Настройка портативного Tesseract OCR"""
    global _TESSERACT_READY
    if _TESSERACT_READY is not None:
        return _TESSERACT_READY

    try:
        # Определяем базовый путь в зависимости от того, запущено ли как exe
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        tesseract_path = os.path.join(base_path, "tesseract", "tesseract.exe")
        tessdata_path = os.path.join(base_path, "tesseract", "tessdata")

        # Проверяем, существует ли портативный Tesseract
        if os.path.exists(tesseract_path) and os.path.exists(tessdata_path):
            # Устанавливаем путь к Tesseract
            import pytesseract
            pytesseract.pytesseract.tesseract_cmd = tesseract_path
            _patch_pytesseract_no_window(pytesseract)

            # Устанавливаем путь к данным
            os.environ['TESSDATA_PREFIX'] = tessdata_path

            # Проверяем, работает ли Tesseract
            try:
                version = pytesseract.get_tesseract_version()
                logging.info(f"Портативный Tesseract OCR найден: версия {version}")
                _TESSERACT_READY = True
                return _TESSERACT_READY
            except Exception as e:
                logging.error(f"Портативный Tesseract найден, но не работает корректно: {e}")
                _TESSERACT_READY = False
                return _TESSERACT_READY
        else:
            logging.warning("Портативный Tesseract не найден или неполная установка")
            _TESSERACT_READY = False
            return _TESSERACT_READY
    except Exception as e:
        logging.error(f"Ошибка при настройке портативного Tesseract: {e}")
        _TESSERACT_READY = False
        return _TESSERACT_READY