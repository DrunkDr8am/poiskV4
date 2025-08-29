import os
import fnmatch
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Set

from tqdm import tqdm
import logging

from file_processing import process_file  # Импортируем функцию обработки файла

def search_files(root_dir: str, extensions: List[str], max_workers: int = 4, output_file: str = None,
                 max_file_size: int = 10, config: dict = None) -> Dict[str, Set[str]]:
    """Многопоточный поиск файлов"""
    results = {}

    # Собираем все файлы для обработки
    files_to_process = []
    for root, _, files in os.walk(root_dir):
        for file in files:
            file_path = os.path.join(root, file)
            if any(fnmatch.fnmatch(file, ext_pattern) for ext_pattern in extensions):
                files_to_process.append(file_path)

    logging.info(f"Найдено файлов для обработки: {len(files_to_process)}")

    # Открываем файл для записи результатов
    output_handle = None
    if output_file:
        output_handle = open(output_file, 'w', encoding='utf-8')
        output_handle.write("Результаты поиска:\n\n")
        output_handle.write(f"Время начала: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n")

    # Обрабатываем файлы в несколько потоков с прогрессбаром
    with tqdm(total=len(files_to_process), desc="Обработка файлов", unit="файл") as pbar:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_file = {
                executor.submit(process_file, file_path, extensions, max_file_size, config): file_path
                for file_path in files_to_process
            }

            for future in as_completed(future_to_file):
                file_path = future_to_file[future]
                try:
                    result = future.result()
                    if result:
                        results.update(result)
                        # Записываем результат в файл сразу
                        if output_handle:
                            for path, keywords_found in result.items():
                                output_handle.write(f"Файл: {path}\n")
                                output_handle.write(f"Найденные ключевые слова: {', '.join(keywords_found)}\n\n")
                                output_handle.flush()  # Принудительно записываем в файл
                except Exception as e:
                    logging.error(f"Ошибка при обработке файла {file_path}: {e}")
                finally:
                    # Обновляем прогрессбар
                    pbar.update(1)
                    pbar.set_postfix(file=os.path.basename(file_path)[:20])  # Показываем имя текущего файла

    # Закрываем файл результатов
    if output_handle:
        output_handle.write(f"\nВремя завершения: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        output_handle.close()

    return results