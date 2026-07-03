import os
import time
from concurrent.futures import ThreadPoolExecutor, wait, FIRST_COMPLETED, TimeoutError as FuturesTimeoutError
from typing import List, Dict, Set, Callable, Optional

import logging

from file_processing import process_file_with_meta  # Импортируем функцию обработки файла

SEARCH_RESULTS_ENCODING = 'utf-8-sig'
PRECOUNT_PROGRESS_INTERVAL = 500


def extension_patterns_to_suffixes(extensions: List[str]) -> Set[str]:
    """Преобразует маски вида *.pdf в множество суффиксов (.pdf)."""
    suffixes = set()
    for pattern in extensions:
        normalized = str(pattern).strip().lower()
        if not normalized:
            continue
        if normalized.startswith('*.'):
            suffixes.add(normalized[1:])
        elif normalized.startswith('*'):
            suffix = normalized[1:]
            if suffix and not suffix.startswith('.'):
                suffix = f".{suffix}"
            if suffix:
                suffixes.add(suffix)
        elif normalized.startswith('.'):
            suffixes.add(normalized)
        else:
            suffixes.add(f".{normalized}")
    return suffixes


def file_matches_extension_suffixes(filename: str, suffixes: Set[str]) -> bool:
    """Быстрая проверка расширения файла по набору суффиксов."""
    if not suffixes:
        return False
    name_lower = filename.lower()
    return any(name_lower.endswith(suffix) for suffix in suffixes)


def collect_matching_files(
    directory: str,
    extensions: List[str],
    progress_callback: Optional[Callable[[int], None]] = None,
    cancel_check: Optional[Callable[[], bool]] = None,
) -> List[str]:
    """Собирает пути файлов с нужными расширениями через os.scandir."""
    suffixes = extension_patterns_to_suffixes(extensions)
    if not suffixes:
        return []

    collected: List[str] = []
    pending_dirs = [directory]
    matched_count = 0

    while pending_dirs:
        if cancel_check and cancel_check():
            break

        current_dir = pending_dirs.pop()
        try:
            with os.scandir(current_dir) as entries:
                for entry in entries:
                    if cancel_check and cancel_check():
                        return collected
                    try:
                        if entry.is_dir(follow_symlinks=False):
                            pending_dirs.append(entry.path)
                        elif entry.is_file(follow_symlinks=False):
                            if file_matches_extension_suffixes(entry.name, suffixes):
                                collected.append(entry.path)
                                matched_count += 1
                                if (
                                    progress_callback
                                    and matched_count % PRECOUNT_PROGRESS_INTERVAL == 0
                                ):
                                    progress_callback(matched_count)
                    except OSError:
                        continue
        except OSError as exc:
            logging.warning(f"Не удалось прочитать каталог {current_dir}: {exc}")

    if progress_callback and matched_count:
        progress_callback(matched_count)
    return collected


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

    return process_file_with_meta(file_path, extensions, max_file_size, config)


def search_files(root_dir: str, extensions: List[str], max_workers: int = 4, output_file: str = None,
                 max_file_size: int = 10, config: dict = None, progress_callback: callable = None,
                 start_count: int = 0, is_searching_func: callable = None,
                 result_callback: callable = None, is_paused_func: callable = None,
                 processed_files_set: set = None, file_completed_callback: callable = None,
                 candidate_files: List[str] = None) -> Dict[str, Set[str]]:
    """Многопоточный поиск файлов с поддержкой offset и проверкой флага остановки"""
    results: Dict[str, Set[str]] = {}

    # Список файлов может быть передан заранее (чтобы избежать повторного обхода ФС).
    files_to_process: List[str] = []
    already_processed_set = processed_files_set or set()
    skipped_processed = 0

    if candidate_files is not None:
        for file_path in candidate_files:
            if file_path in already_processed_set:
                skipped_processed += 1
                continue
            files_to_process.append(file_path)
    else:
        suffixes = extension_patterns_to_suffixes(extensions)
        for root, _, files in os.walk(root_dir):
            if not _wait_if_paused(is_paused_func, is_searching_func):
                break

            # Проверяем флаг остановки перед обработкой каждой папки
            if is_searching_func and not is_searching_func():
                logging.info("Поиск остановлен пользователем при сборе файлов")
                break

            for file in files:
                if not file_matches_extension_suffixes(file, suffixes):
                    continue
                file_path = os.path.join(root, file)
                if file_path in already_processed_set:
                    skipped_processed += 1
                    continue
                files_to_process.append(file_path)

    logging.info(f"Найдено файлов для обработки в {root_dir}: {len(files_to_process)}")
    if skipped_processed:
        logging.info(
            f"Пропущено уже обработанных файлов в {root_dir}: {skipped_processed}"
        )

    # Открываем файл для записи результатов
    output_handle = None
    if output_file:
        try:
            file_is_empty = (not os.path.exists(output_file)) or os.path.getsize(output_file) == 0
            if start_count == 0 and file_is_empty:
                output_handle = open(output_file, 'w', encoding=SEARCH_RESULTS_ENCODING)
            else:
                output_handle = open(output_file, 'a', encoding=SEARCH_RESULTS_ENCODING)

            # Пишем заголовок только если файл пустой и это самое начало
            if start_count == 0 and file_is_empty:
                output_handle.write("Результаты поиска:\n\n")
                output_handle.write(f"Время начала: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        except OSError as e:
            logging.error(f"Не удалось проверить размер файла результатов {output_file}: {e}")

    # Обрабатываем файлы в несколько потоков, подавая задачи порциями.
    executor = ThreadPoolExecutor(max_workers=max_workers)
    stop_requested = False
    completed_count = start_count
    try:
        files_iter = iter(files_to_process)
        in_flight = {}
        max_in_flight = max(1, max_workers * 2)

        def submit_next():
            while len(in_flight) < max_in_flight:
                try:
                    file_path = next(files_iter)
                except StopIteration:
                    break
                future = executor.submit(
                    _process_file_with_start,
                    file_path,
                    extensions,
                    max_file_size,
                    config,
                    progress_callback,
                    is_searching_func,
                    is_paused_func
                )
                in_flight[future] = file_path

        submit_next()

        while in_flight:
            if not _wait_if_paused(is_paused_func, is_searching_func):
                logging.info("Поиск остановлен пользователем во время паузы")
                stop_requested = True
                for f in in_flight:
                    f.cancel()
                break

            # Проверяем флаг остановки перед обработкой каждого файла
            if is_searching_func and not is_searching_func():
                logging.info("Поиск остановлен пользователем во время обработки файлов")
                stop_requested = True
                # Отменяем все оставшиеся задачи
                for f in in_flight:
                    f.cancel()
                break

            done, _ = wait(in_flight.keys(), return_when=FIRST_COMPLETED)
            for future in done:
                file_path = in_flight.pop(future)
                completed_count += 1
                total_processed = completed_count

                # Вызываем callback для обновления прогресса в GUI
                if progress_callback and callable(progress_callback):
                    try:
                        progress_callback(
                            f"Готово: {os.path.basename(file_path)}",
                            total_processed,
                            len(in_flight),
                        )
                    except Exception as e:
                        logging.error(f"Ошибка в callback обновления прогресса: {e}")

                file_had_matches = False
                file_error = ""
                file_status = "no_match"
                skip_reason = ""
                try:
                    payload = future.result(timeout=300)
                    if isinstance(payload, tuple) and len(payload) >= 4:
                        result, file_status, file_error, skip_reason = payload[0], payload[1], payload[2] or "", payload[3] or ""
                    elif isinstance(payload, tuple) and len(payload) >= 3:
                        result, file_status, file_error = payload[0], payload[1], payload[2] or ""
                    else:
                        result = payload or {}
                        file_status = "matched" if result else "no_match"
                    if result:
                        file_had_matches = True
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
                    file_status = "error"
                    file_error = f"Таймаут при обработке файла {file_path}"
                    logging.error(file_error)
                except Exception as e:
                    file_status = "error"
                    file_error = f"Ошибка при обработке файла {file_path}: {e}"
                    logging.error(file_error)
                finally:
                    if file_completed_callback and callable(file_completed_callback):
                        try:
                            file_completed_callback(file_path, file_had_matches, file_error, file_status, skip_reason)
                        except Exception as callback_error:
                            logging.error(f"Ошибка в callback завершения файла: {callback_error}")

            submit_next()
    finally:
        # При остановке не блокируемся, ожидая завершения всех worker'ов.
        if stop_requested:
            executor.shutdown(wait=False, cancel_futures=True)
        else:
            executor.shutdown(wait=True)

    if output_handle:
        output_handle.close()

    logging.info(f"Завершена обработка директории {root_dir}. Найдено совпадений: {len(results)}")
    return results