import os
import fnmatch
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Set

from tqdm import tqdm
import logging

from file_processing import process_file  # Импортируем функцию обработки файла


def search_files(root_dir: str, extensions: List[str], max_workers: int = 4, output_file: str = None,
                 max_file_size: int = 10, config: dict = None, progress_callback: callable = None,
                 start_count: int = 0) -> Dict[str, Set[str]]:
    """Многопоточный поиск файлов с поддержкой offset"""
    results = {}

    # Собираем все файлы для обработки
    files_to_process = []
    for root, _, files in os.walk(root_dir):
        for file in files:
            file_path = os.path.join(root, file)
            if any(fnmatch.fnmatch(file, ext_pattern) for ext_pattern in extensions):
                files_to_process.append(file_path)

    logging.info(f"Найдено файлов для обработки в {root_dir}: {len(files_to_process)}")

    # Открываем файл для записи результатов
    output_handle = None
    if output_file:
        output_handle = open(output_file, 'a', encoding='utf-8')
        if start_count == 0 and os.path.getsize(output_file) == 0:
            output_handle.write("Результаты поиска:\n\n")
            output_handle.write(f"Время начала: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n")

    # Обрабатываем файлы в несколько потоков
    with tqdm(total=len(files_to_process), desc=f"Обработка {os.path.basename(root_dir)}", unit="файл") as pbar:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_file = {
                executor.submit(process_file, file_path, extensions, max_file_size, config): file_path
                for file_path in files_to_process
            }

            for i, future in enumerate(as_completed(future_to_file)):
                file_path = future_to_file[future]
                total_processed = start_count + i + 1

                # Вызываем callback для обновления прогресса в GUI
                if progress_callback and callable(progress_callback):
                    try:
                        # Передаем только имя файла, не специальные сообщения
                        progress_callback(os.path.basename(file_path), total_processed)
                    except Exception as e:
                        logging.error(f"Ошибка в callback: {e}")

                try:
                    result = future.result(timeout=300)
                    if result:
                        results.update(result)
                        if output_handle:
                            for path, keywords_found in result.items():
                                output_handle.write(f"Файл: {path}\n")
                                output_handle.write(f"Найденные ключевые слова: {', '.join(keywords_found)}\n\n")
                                output_handle.flush()
                except TimeoutError:
                    logging.error(f"Таймаут при обработке файла {file_path}")
                except Exception as e:
                    logging.error(f"Ошибка при обработке файла {file_path}: {e}")
                finally:
                    pbar.update(1)
                    pbar.set_postfix(file=os.path.basename(file_path)[:20])

    if output_handle:
        output_handle.close()

    logging.info(f"Завершена обработка директории {root_dir}. Найдено совпадений: {len(results)}")
    return results