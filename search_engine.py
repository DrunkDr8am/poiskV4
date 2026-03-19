import os
import fnmatch
import time
from concurrent.futures import ThreadPoolExecutor, as_completed, TimeoutError as FuturesTimeoutError
from typing import List, Dict, Set

import logging

from file_processing import process_file  # Импортируем функцию обработки файла

SEARCH_RESULTS_ENCODING = 'utf-8-sig'


def _wait_if_paused(is_paused_func: callable = None, is_searching_func: callable = None) -> bool:
    """Ожидание снятия паузы. Возвращает False, если поиск остановлен."""
    while is_paused_func and is_paused_func():
        if is_searching_func and not is_searching_func():
            return False
        time.sleep(0.1)
    return True


def _process_file_with_start(file_path: str, extensions: List[str], max_file_size: int, config: dict,
                             progress_callback: callable = None, is_searching_func: callable = None,
                             is_paused_func: callable = None):
    """Обертка для отправки статуса старта обработки файла."""
    if is_searching_func and not is_searching_func():
        return {}

    if not _wait_if_paused(is_paused_func, is_searching_func):
        return {}

    if progress_callback and callable(progress_callback):
        try:
            progress_callback(f"Начат: {os.path.basename(file_path)}", None)
        except Exception as e:
            logging.error(f"Ошибка в callback старта обработки: {e}")

    if not _wait_if_paused(is_paused_func, is_searching_func):
        return {}

    if is_searching_func and not is_searching_func():
        return {}

    return process_file(file_path, extensions, max_file_size, config)


def search_files(root_dir: str, extensions: List[str], max_workers: int = 4, output_file: str = None,
                 max_file_size: int = 10, config: dict = None, progress_callback: callable = None,
                 start_count: int = 0, is_searching_func: callable = None,
                 result_callback: callable = None, is_paused_func: callable = None) -> Dict[str, Set[str]]:
    """Многопоточный поиск файлов с поддержкой offset и проверкой флага остановки"""
    results: Dict[str, Set[str]] = {}

    # Собираем все файлы для обработки
    files_to_process: List[str] = []
    for root, _, files in os.walk(root_dir):
        if not _wait_if_paused(is_paused_func, is_searching_func):
            break

        # Проверяем флаг остановки перед обработкой каждой папки
        if is_searching_func and not is_searching_func():
            logging.info("Поиск остановлен пользователем при сборе файлов")
            break

        for file in files:
            file_path = os.path.join(root, file)
            if any(fnmatch.fnmatch(file, ext_pattern) for ext_pattern in extensions):
                files_to_process.append(file_path)

    logging.info(f"Найдено файлов для обработки в {root_dir}: {len(files_to_process)}")

    # Открываем файл для записи результатов
    output_handle = None
    if output_file:
        try:
            file_is_empty = (not os.path.exists(output_file)) or os.path.getsize(output_file) == 0
            if file_is_empty:
                output_handle = open(output_file, 'w', encoding=SEARCH_RESULTS_ENCODING)
            else:
                output_handle = open(output_file, 'a', encoding='utf-8')

            if file_is_empty:
                output_handle.write("Результаты поиска:\n\n")
                output_handle.write(f"Директория: {root_dir}\n")
                output_handle.write(f"Время начала: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        except OSError as e:
            logging.error(f"Не удалось проверить размер файла результатов {output_file}: {e}")

    # Обрабатываем файлы в несколько потоков
    executor = ThreadPoolExecutor(max_workers=max_workers)
    stop_requested = False
    try:
        future_to_file = {
            executor.submit(
                _process_file_with_start,
                file_path,
                extensions,
                max_file_size,
                config,
                progress_callback,
                is_searching_func,
                is_paused_func
            ): file_path
            for file_path in files_to_process
        }

        for i, future in enumerate(as_completed(future_to_file)):
            if not _wait_if_paused(is_paused_func, is_searching_func):
                logging.info("Поиск остановлен пользователем во время паузы")
                stop_requested = True
                for f in future_to_file:
                    f.cancel()
                break

            # Проверяем флаг остановки перед обработкой каждого файла
            if is_searching_func and not is_searching_func():
                logging.info("Поиск остановлен пользователем во время обработки файлов")
                stop_requested = True
                # Отменяем все оставшиеся задачи
                for f in future_to_file:
                    f.cancel()
                break

            file_path = future_to_file[future]
            total_processed = start_count + i + 1

            # Вызываем callback для обновления прогресса в GUI
            if progress_callback and callable(progress_callback):
                try:
                    # Для завершения файла обновляем только счетчик прогресса,
                    # чтобы не перезатирать статус "Начат: <файл>".
                    progress_callback("", total_processed)
                except Exception as e:
                    logging.error(f"Ошибка в callback обновления прогресса: {e}")

            try:
                result = future.result(timeout=300)
                if result:
                    results.update(result)
                    if result_callback and callable(result_callback):
                        try:
                            for path, keywords_found in result.items():
                                result_callback(path, keywords_found)
                        except Exception as e:
                            logging.error(f"Ошибка в callback результата: {e}")
                    if output_handle:
                        for path, keywords_found in result.items():
                            output_handle.write(f"Файл: {path}\n")
                            output_handle.write(f"Найденные ключевые слова: {', '.join(keywords_found)}\n\n")
                            output_handle.flush()
            except FuturesTimeoutError:
                logging.error(f"Таймаут при обработке файла {file_path}")
            except Exception as e:
                logging.error(f"Ошибка при обработке файла {file_path}: {e}")
    finally:
        # При остановке не блокируемся, ожидая завершения всех worker'ов.
        if stop_requested:
            executor.shutdown(wait=False, cancel_futures=True)
        else:
            executor.shutdown(wait=True)

    if output_handle:
        if not results and not stop_requested:
            output_handle.write("Совпадения не найдены.\n")
            output_handle.flush()
        output_handle.close()

    logging.info(f"Завершена обработка директории {root_dir}. Найдено совпадений: {len(results)}")
    return results