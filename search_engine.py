import os
import fnmatch
import time
from concurrent.futures import ThreadPoolExecutor, as_completed, TimeoutError as FuturesTimeoutError
from typing import List, Dict, Set

import logging

from file_processing import process_file  # Импортируем функцию обработки файла


def search_files(root_dir: str, extensions: List[str], max_workers: int = 4, output_file: str = None,
                 max_file_size: int = 10, config: dict = None, progress_callback: callable = None,
                 start_count: int = 0, is_searching_func: callable = None,
                 result_callback: callable = None) -> Dict[str, Set[str]]:
    """Многопоточный поиск файлов с поддержкой offset и проверкой флага остановки"""
    results: Dict[str, Set[str]] = {}

    # Собираем все файлы для обработки
    files_to_process: List[str] = []
    for root, _, files in os.walk(root_dir):
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
        output_handle = open(output_file, 'a', encoding='utf-8')
        try:
            # Пишем заголовок только если файл пустой и это самое начало
            if start_count == 0 and os.path.getsize(output_file) == 0:
                output_handle.write("Результаты поиска:\n\n")
                output_handle.write(f"Время начала: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        except OSError as e:
            logging.error(f"Не удалось проверить размер файла результатов {output_file}: {e}")

    # Обрабатываем файлы в несколько потоков
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_file = {
            executor.submit(process_file, file_path, extensions, max_file_size, config): file_path
            for file_path in files_to_process
        }

        for i, future in enumerate(as_completed(future_to_file)):
            # Проверяем флаг остановки перед обработкой каждого файла
            if is_searching_func and not is_searching_func():
                logging.info("Поиск остановлен пользователем во время обработки файлов")
                # Отменяем все оставшиеся задачи
                for f in future_to_file:
                    f.cancel()
                break

            file_path = future_to_file[future]
            total_processed = start_count + i + 1

            # Вызываем callback для обновления прогресса в GUI
            if progress_callback and callable(progress_callback):
                try:
                    # Передаем только имя файла, не специальные сообщения
                    progress_callback(os.path.basename(file_path), total_processed)
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

    if output_handle:
        output_handle.close()

    logging.info(f"Завершена обработка директории {root_dir}. Найдено совпадений: {len(results)}")
    return results