import os
import sys
import logging


def setup_tesseract():
    """Настройка портативного Tesseract OCR"""
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

            # Устанавливаем путь к данным
            os.environ['TESSDATA_PREFIX'] = tessdata_path

            # Проверяем, работает ли Tesseract
            try:
                version = pytesseract.get_tesseract_version()
                logging.info(f"Портативный Tesseract OCR найден: версия {version}")
                return True
            except Exception as e:
                logging.error(f"Портативный Tesseract найден, но не работает корректно: {e}")
                return False
        else:
            logging.warning("Портативный Tesseract не найден или неполная установка")
            return False
    except Exception as e:
        logging.error(f"Ошибка при настройке портативного Tesseract: {e}")
        return False