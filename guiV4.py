import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
import threading
import os
import logging
import time  # Добавляем импорт модуля time
import subprocess
import sys
import traceback
import json
import hashlib
import tkinter.font as tkfont
try:
    import sqlite3
except ImportError:  # pragma: no cover - крайне редкий случай для стандартного Python
    sqlite3 = None
from config_loader import load_config, create_default_config
from tesseract_setup import setup_tesseract
from file_processing import load_keywords, run_pdf_worker_cli, ARCHIVE_MEMBER_SEP
from search_engine import search_files
from configparser import ConfigParser

import fnmatch
import zipfile
import tempfile
import shutil

# Глобальные флаги для доступности функций
HAS_PDF = False
HAS_DOCX = False
HAS_EXCEL = False
HAS_7Z = False
HAS_RAR = False
HAS_OCR = False
APP_PASSWORD = "1511"
INVALID_PASSWORD_CLOSE_MS = 300_000
INVALID_PASSWORD_TICK_MS = 1000
CRASH_LOG_FILE = "crash_log.txt"
SEARCH_STATE_DB_FILE = "search_state.db"
SEARCH_STATE_VERSION = 1


def write_crash_report(error_title, exc_value, exc_traceback):
    """Сохраняет подробности необработанной ошибки в отдельный лог-файл."""
    report_time = time.strftime('%Y-%m-%d %H:%M:%S')
    traceback_text = "".join(traceback.format_exception(type(exc_value), exc_value, exc_traceback))
    report = (
        f"[{report_time}] {error_title}\n"
        f"{traceback_text}\n"
        f"{'-' * 80}\n"
    )
    log_path = os.path.abspath(CRASH_LOG_FILE)
    with open(log_path, "a", encoding="utf-8") as crash_file:
        crash_file.write(report)
    return log_path


class SearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Поиск файлов по ключевым словам")
        self.root.geometry("1000x800")
        self.root.minsize(900, 700)

        # Переменные для хранения состояний
        self.extension_options = [
            '*.txt', '*.pdf',
            '*.doc', '*.docx', '*.docm', '*.dot', '*.dotx', '*.dotm',
            '*.xls', '*.xlsx', '*.xlsm', '*.xlt', '*.xltx', '*.xltm',
            '*.jpg', '*.jpeg', '*.jpe', '*.jfif', '*.png', '*.bmp', '*.gif', '*.tif', '*.tiff', '*.webp', '*.ico',
            '*.zip', '*.rar', '*.7z'
        ]
        self.extension_vars = {ext: tk.BooleanVar(value=False) for ext in self.extension_options}
        self.directories_list = []
        self.is_searching = False
        self.is_paused = False
        self.search_thread = None
        self.directory_add_thread = None
        self.is_adding_directory = False
        self.is_closing = False
        self.config_dirty = False
        self._updating_threads_var = False
        self._updating_ocr_threads_var = False
        self.invalid_password_timer_id = None
        self.invalid_password_deadline_ms = None
        self.search_file_handler = None
        self.search_state_lock = threading.Lock()
        self.active_search_state = None
        self.resume_processed_files = set()
        self.directory_files_map = {}
        self._pending_state_updates = 0
        self._state_flush_interval = 25
        self._reset_processed_files_table = False
        self._saved_processed_count = 0
        self._unsaved_processed_paths = []
        self.resume_paths_pending_load = False
        self.resume_start_count = 0
        self.search_session_status_var = tk.StringVar(value="Статус сессии: нет данных")
        self.search_session_remaining_var = tk.StringVar(value="Осталось файлов: -")
        self.search_session_last_file_var = tk.StringVar(value="Последний файл: -")
        self.dashboard_processed_var = tk.StringVar(value="0")
        self.dashboard_found_var = tk.StringVar(value="0")
        self.dashboard_skipped_var = tk.StringVar(value="0")
        self.dashboard_errors_var = tk.StringVar(value="0")
        self.dashboard_found = 0
        self.found_results = {}
        self.dashboard_skipped = 0
        self.dashboard_errors = 0
        self.dashboard_skip_reasons = {
            "long_path": 0,
            "large_file": 0,
            "module_unavailable": 0,
            "read_error": 0,
            "other": 0,
        }
        self.progress_value = tk.DoubleVar(value=0.0)
        self.current_file = tk.StringVar(value="")
        self.theme_var = tk.StringVar(value="Светлая")
        self.total_files = 0
        self.processed_files = 0
        self.search_in_flight_count = 0
        self.search_start_time = None  # Время начала поиска
        self.search_end_time = None  # Время окончания поиска
        self.max_result_file_text_px = 0

        # Загружаем конфигурацию ДО создания интерфейса
        self.config = self.load_configuration()

        # Проверяем зависимости
        self.check_dependencies()

        # Создаем интерфейс
        self.create_widgets()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.init_search_state_storage()
        self.found_results = self._load_found_results_from_db()
        self.dashboard_found = len(self.found_results)
        self.refresh_search_session_ui()

        # Центрируем окно
        self.center_window()

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    @staticmethod
    def get_auto_threads_count():
        """Автовыбор потоков: CPU-2, либо 1 при CPU<=2."""
        cpu_count = os.cpu_count() or 1
        return cpu_count - 2 if cpu_count > 2 else 1

    @staticmethod
    def get_default_ocr_threads_count():
        """Рекомендуемое число OCR-потоков: 2 или меньше при слабом CPU."""
        cpu_count = os.cpu_count() or 1
        return min(2, max(1, cpu_count))

    @staticmethod
    def get_max_threads_count():
        """Максимально допустимое количество потоков на текущем компьютере."""
        return max(1, os.cpu_count() or 1)

    def normalize_threads_value(self):
        """Нормализует количество потоков в диапазон [1, cpu_count]."""
        max_threads = self.get_max_threads_count()
        try:
            threads = int(self.threads_var.get())
        except (TypeError, ValueError):
            threads = self.get_auto_threads_count()
        threads = max(1, min(threads, max_threads))
        self.threads_var.set(str(threads))
        return threads

    def normalize_ocr_threads_value(self):
        """Нормализует количество OCR-потоков в диапазон [1, cpu_count]."""
        max_threads = self.get_max_threads_count()
        try:
            ocr_threads = int(self.ocr_threads_var.get())
        except (TypeError, ValueError):
            ocr_threads = self.get_default_ocr_threads_count()
        ocr_threads = max(1, min(ocr_threads, max_threads))
        self.ocr_threads_var.set(str(ocr_threads))
        return ocr_threads

    def on_ocr_threads_var_change(self, *_args):
        """Не дает вручную ввести число OCR-потоков больше доступного."""
        if self._updating_ocr_threads_var:
            return

        value = self.ocr_threads_var.get()
        if value == "":
            return

        filtered = "".join(ch for ch in value if ch.isdigit())
        max_threads = self.get_max_threads_count()

        if not filtered:
            new_value = "1"
        else:
            new_value = str(min(max(1, int(filtered)), max_threads))

        if new_value != value:
            self._updating_ocr_threads_var = True
            try:
                self.ocr_threads_var.set(new_value)
            finally:
                self._updating_ocr_threads_var = False

    def on_threads_var_change(self, *_args):
        """Не дает вручную ввести число потоков больше доступного."""
        if self._updating_threads_var:
            return

        value = self.threads_var.get()
        if value == "":
            return

        filtered = "".join(ch for ch in value if ch.isdigit())
        max_threads = self.get_max_threads_count()

        if not filtered:
            new_value = "1"
        else:
            new_value = str(min(max(1, int(filtered)), max_threads))

        if new_value != value:
            self._updating_threads_var = True
            try:
                self.threads_var.set(new_value)
            finally:
                self._updating_threads_var = False

    def load_configuration(self):
        """Загрузка конфигурации"""
        if not os.path.exists("config.txt"):
            create_default_config()

        config = load_config()
        selected_extensions = [ext.strip() for ext in config.get('extensions', []) if ext.strip()]
        selected_set = set(selected_extensions)
        self.theme_var.set(config.get('theme', 'Светлая'))

        # Сохраняем пользовательские расширения из конфига и показываем их как отдельные чекбоксы
        for ext in selected_extensions:
            if ext not in self.extension_vars:
                self.extension_options.append(ext)
                self.extension_vars[ext] = tk.BooleanVar(value=False)

        for ext, ext_var in self.extension_vars.items():
            ext_var.set(ext in selected_set)

        return {
            'config': config
        }

    def check_dependencies(self):
        """Проверка доступности опциональных зависимостей"""
        global HAS_PDF, HAS_DOCX, HAS_EXCEL, HAS_7Z, HAS_RAR, HAS_OCR

        # Проверяем доступность модулей
        try:
            import fitz
            HAS_PDF = True
        except ImportError:
            logging.warning("Модуль PyMuPDF не установлен. Поддержка PDF отключена.")

        try:
            import docx2txt
            HAS_DOCX = True
        except ImportError:
            logging.warning("Модули для DOCX не установлены. Поддержка DOCX отключена.")

        try:
            import pandas as pd
            import openpyxl
            HAS_EXCEL = True
        except ImportError:
            logging.warning("Модули для Excel не установлены. Поддержка Excel отключена.")

        try:
            import py7zr
            HAS_7Z = True
        except ImportError:
            logging.warning("Модуль py7zr не установлен. Поддержка 7z архивов отключена.")

        try:
            import rarfile
            HAS_RAR = True
        except ImportError:
            logging.warning("Модуль rarfile не установлен. Поддержка RAR архивов отключена.")

        # Tesseract настраивается отдельно
        HAS_OCR = setup_tesseract()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        search_tab = ttk.Frame(self.notebook, padding="8")
        settings_tab = ttk.Frame(self.notebook, padding="8")
        self.notebook.add(search_tab, text="Поиск")
        self.notebook.add(settings_tab, text="Настройки поиска")

        # ---------------- Поиск ----------------
        search_tab.columnconfigure(1, weight=1)
        search_tab.rowconfigure(5, weight=1)

        ttk.Label(search_tab, text="Директории для поиска:").grid(row=0, column=0, sticky=tk.NW, pady=5)
        dir_frame = ttk.Frame(search_tab)
        dir_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        dir_frame.columnconfigure(0, weight=1)
        dir_frame.rowconfigure(1, weight=0)

        self.dirs_listbox = tk.Listbox(dir_frame, height=5)
        self.dirs_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        scrollbar = ttk.Scrollbar(dir_frame, orient=tk.VERTICAL, command=self.dirs_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.dirs_listbox.configure(yscrollcommand=scrollbar.set)

        dir_btn_frame = ttk.Frame(dir_frame)
        dir_btn_frame.grid(row=0, column=2, padx=(6, 0), sticky=tk.N)
        ttk.Button(dir_btn_frame, text="Добавить", command=self.add_directory).pack(pady=2)
        ttk.Button(dir_btn_frame, text="Удалить", command=self.remove_directory).pack(pady=2)

        self.adding_dir_frame = ttk.Frame(dir_frame)
        self.adding_dir_frame.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(6, 0))
        self.adding_dir_label = ttk.Label(self.adding_dir_frame, text="Добавление директории...")
        self.adding_dir_label.pack(side=tk.LEFT, padx=(0, 8))
        self.adding_dir_spinner = ttk.Progressbar(self.adding_dir_frame, mode="indeterminate", length=150)
        self.adding_dir_spinner.pack(side=tk.LEFT)
        self.adding_dir_frame.grid_remove()

        initial_directories = self.config['config'].get('directories', [])
        if not initial_directories:
            fallback_directory = self.config['config'].get('directory', '.')
            initial_directories = [fallback_directory]

        for directory in initial_directories:
            if directory not in self.directories_list:
                self.directories_list.append(directory)
                self.dirs_listbox.insert(tk.END, directory)

        ttk.Label(search_tab, text="Дашборд:").grid(row=1, column=0, sticky=tk.W, pady=5)
        dashboard_frame = ttk.Frame(search_tab, style="DashboardWrap.TFrame")
        dashboard_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        for col_idx in range(4):
            dashboard_frame.columnconfigure(col_idx, weight=1)

        dashboard_cards = [
            ("Проверено", self.dashboard_processed_var),
            ("Найдено", self.dashboard_found_var),
            ("Пропущено", self.dashboard_skipped_var),
            ("Ошибки", self.dashboard_errors_var),
        ]
        for index, (title, value_var) in enumerate(dashboard_cards):
            card = ttk.Frame(dashboard_frame, style="DashboardCard.TFrame", padding=(10, 8))
            card.grid(row=0, column=index, sticky=(tk.W, tk.E), padx=4)
            title_label = ttk.Label(card, text=title, style="DashboardCardTitle.TLabel")
            title_label.grid(row=0, column=0, sticky=tk.W)
            value_label = ttk.Label(card, textvariable=value_var, style="DashboardCardValue.TLabel")
            value_label.grid(
                row=1, column=0, sticky=tk.W, pady=(4, 0)
            )
            if title == "Пропущено":
                self.dashboard_skipped_card = card
                self.dashboard_skipped_card.configure(cursor="hand2")
                card.bind("<Button-1>", self.show_skip_details)
                title_label.bind("<Button-1>", self.show_skip_details)
                value_label.bind("<Button-1>", self.show_skip_details)
            elif title == "Найдено":
                self.dashboard_found_card = card
                self.dashboard_found_card.configure(cursor="hand2")
                card.bind("<Button-1>", self.show_found_results)
                title_label.bind("<Button-1>", self.show_found_results)
                value_label.bind("<Button-1>", self.show_found_results)

        ttk.Label(search_tab, text="Прогресс:").grid(row=2, column=0, sticky=tk.W, pady=5)
        progress_frame = ttk.Frame(search_tab)
        progress_frame.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)
        progress_frame.columnconfigure(0, weight=1)
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_value, maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Label(progress_frame, textvariable=self.current_file).grid(row=1, column=0, sticky=(tk.W, tk.E))

        session_cards_frame = ttk.Frame(progress_frame, style="SessionCardsWrap.TFrame")
        session_cards_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(8, 0))
        session_cards_frame.columnconfigure(0, weight=1)
        session_cards_frame.columnconfigure(1, weight=1)

        status_card = ttk.Frame(session_cards_frame, style="SessionCard.TFrame", padding=(12, 8))
        status_card.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 6))
        ttk.Label(status_card, text="Статус сессии", style="SessionCardTitle.TLabel").grid(row=0, column=0, sticky=tk.W)
        ttk.Label(status_card, textvariable=self.search_session_status_var, style="SessionCardValue.TLabel").grid(
            row=1, column=0, sticky=tk.W, pady=(4, 0)
        )

        remaining_card = ttk.Frame(session_cards_frame, style="SessionCard.TFrame", padding=(12, 8))
        remaining_card.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(6, 0))
        ttk.Label(remaining_card, text="Осталось файлов", style="SessionCardTitle.TLabel").grid(row=0, column=0, sticky=tk.W)
        ttk.Label(remaining_card, textvariable=self.search_session_remaining_var, style="SessionCardValue.TLabel").grid(
            row=1, column=0, sticky=tk.W, pady=(4, 0)
        )

        button_frame = ttk.Frame(search_tab)
        button_frame.grid(row=3, column=0, columnspan=2, pady=10)
        self.start_button = ttk.Button(button_frame, text="Начать поиск", command=self.start_search)
        self.start_button.pack(side=tk.LEFT, padx=5)
        self.pause_button = ttk.Button(button_frame, text="Пауза", command=self.toggle_pause, state=tk.DISABLED)
        self.pause_button.pack(side=tk.LEFT, padx=5)
        self.stop_button = ttk.Button(button_frame, text="Закончить поиск", command=self.stop_search, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Сохранить результаты", command=self.save_results).pack(side=tk.LEFT, padx=5)

        ttk.Label(search_tab, text="Результаты поиска:").grid(row=4, column=0, sticky=tk.NW, pady=5)
        results_frame = ttk.Frame(search_tab)
        results_frame.grid(row=4, column=1, rowspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        self.results_table = ttk.Treeview(results_frame, columns=("keywords", "file"), show="headings", height=14)
        self.results_table.heading("keywords", text="Найденные слова")
        self.results_table.heading("file", text="Ссылка на файл")
        self.results_table.column("keywords", width=150, minwidth=100, stretch=False, anchor=tk.W)
        self.results_table.column("file", width=850, minwidth=300, stretch=False, anchor=tk.W)
        results_scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_table.yview)
        results_h_scrollbar = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL, command=self.results_table.xview)
        self.results_table.configure(yscrollcommand=results_scrollbar.set, xscrollcommand=results_h_scrollbar.set)
        self.results_table.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        results_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        results_h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.results_h_scrollbar = results_h_scrollbar
        self.results_h_scrollbar.grid_remove()
        self.hovered_result_item = None
        self.results_context_menu = tk.Menu(self.root, tearoff=0)
        self.results_context_menu.add_command(label="Перейти к расположению файла", command=self.open_result_location)
        self.results_context_menu.add_command(label="Открыть файл", command=self.open_selected_result_file)
        self.results_table.bind("<Configure>", self.on_results_table_resize)
        self.results_table.bind("<Double-Button-1>", self.open_result_file)
        self.results_table.bind("<Button-3>", self.show_results_context_menu)
        self.results_table.bind("<Motion>", self.on_results_hover)
        self.results_table.bind("<Leave>", self.on_results_leave)

        # ---------------- Настройки поиска ----------------
        settings_tab.columnconfigure(1, weight=1)
        settings_tab.rowconfigure(11, weight=1)
        settings_tab.grid_anchor("nw")

        # Поля настроек идут по порядку
        ttk.Label(settings_tab, text="Расширения файлов:").grid(row=0, column=0, sticky=tk.NW, pady=3)
        extensions_frame = ttk.Frame(settings_tab)
        extensions_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=3)
        columns = 5
        for index, ext in enumerate(self.extension_options):
            row = index // columns
            col = index % columns
            ttk.Checkbutton(extensions_frame, text=ext, variable=self.extension_vars[ext]).grid(
                row=row, column=col, sticky=tk.W, padx=(0, 12), pady=2
            )

        ttk.Label(settings_tab, text="Ключевые слова:").grid(row=1, column=0, sticky=tk.NW, pady=3)
        keywords_frame = ttk.Frame(settings_tab)
        keywords_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=3)
        keywords_frame.columnconfigure(0, weight=1)
        self.keywords_text = scrolledtext.ScrolledText(keywords_frame, height=5)
        self.keywords_text.grid(row=0, column=0, sticky=(tk.W, tk.E))

        if os.path.exists("keywords.txt"):
            try:
                with open("keywords.txt", "r", encoding="utf-8") as f:
                    self.keywords_text.insert("1.0", f.read())
            except Exception:
                pass

        ttk.Label(settings_tab, text="Потоки:").grid(row=2, column=0, sticky=tk.W, pady=3)
        auto_threads = self.get_auto_threads_count()
        self.threads_var = tk.StringVar(value=str(auto_threads))
        self.threads_var.trace_add("write", self.on_threads_var_change)
        threads_spin = ttk.Spinbox(
            settings_tab, from_=1, to=self.get_max_threads_count(), textvariable=self.threads_var, width=8
        )
        threads_spin.grid(row=2, column=1, sticky=tk.W, pady=3)

        ttk.Label(settings_tab, text="OCR-потоки (Tesseract):").grid(row=3, column=0, sticky=tk.W, pady=3)
        default_ocr_threads = str(self.config['config'].get('ocr_threads', self.get_default_ocr_threads_count()))
        self.ocr_threads_var = tk.StringVar(value=default_ocr_threads)
        self.ocr_threads_var.trace_add("write", self.on_ocr_threads_var_change)
        ocr_threads_spin = ttk.Spinbox(
            settings_tab, from_=1, to=self.get_max_threads_count(), textvariable=self.ocr_threads_var, width=8
        )
        ocr_threads_spin.grid(row=3, column=1, sticky=tk.W, pady=3)

        ttk.Label(settings_tab, text="Макс. размер файла (МБ, 0=без лимита):").grid(row=4, column=0, sticky=tk.W, pady=3)
        self.max_size_var = tk.StringVar(value=str(self.config['config'].get('max_file_size', 50)))
        max_size_spin = ttk.Spinbox(settings_tab, from_=0, to=1000, textvariable=self.max_size_var, width=8)
        max_size_spin.grid(row=4, column=1, sticky=tk.W, pady=3)

        self.search_images_var = tk.BooleanVar(value=self.config['config'].get('search_images', False))
        ttk.Label(settings_tab, text="Поиск по изображениям (OCR):").grid(row=5, column=0, sticky=tk.W, pady=3)
        ttk.Checkbutton(settings_tab, text="Включить OCR", variable=self.search_images_var).grid(
            row=5, column=1, sticky=tk.W, pady=3
        )

        ttk.Label(settings_tab, text="Макс. страниц PDF (0=без лимита):").grid(row=6, column=0, sticky=tk.W, pady=3)
        self.max_pdf_pages_var = tk.StringVar(value=str(self.config['config'].get('max_pdf_pages', 0)))
        max_pdf_pages_spin = ttk.Spinbox(settings_tab, from_=0, to=100000, textvariable=self.max_pdf_pages_var, width=10)
        max_pdf_pages_spin.grid(row=6, column=1, sticky=tk.W, pady=3)

        ttk.Label(settings_tab, text="Макс. длина пути (0=без лимита):").grid(row=7, column=0, sticky=tk.W, pady=3)
        self.max_path_length_var = tk.StringVar(value=str(self.config['config'].get('max_path_length', 240)))
        max_path_spin = ttk.Spinbox(settings_tab, from_=0, to=10000, textvariable=self.max_path_length_var, width=10)
        max_path_spin.grid(row=7, column=1, sticky=tk.W, pady=3)

        ttk.Label(settings_tab, text="Режим запуска поиска:").grid(row=8, column=0, sticky=tk.NW, pady=(8, 3))
        self.pre_count_files_var = tk.BooleanVar(value=bool(self.config['config'].get('pre_count_files', True)))
        launch_mode_frame = ttk.Frame(settings_tab)
        launch_mode_frame.grid(row=8, column=1, sticky=tk.W, pady=(8, 3))
        ttk.Radiobutton(
            launch_mode_frame,
            text="Сначала считать файлы, затем искать",
            variable=self.pre_count_files_var,
            value=True
        ).grid(row=0, column=0, sticky=tk.W, pady=(0, 2))
        ttk.Radiobutton(
            launch_mode_frame,
            text="Сразу запускать поиск (без подсчета)",
            variable=self.pre_count_files_var,
            value=False
        ).grid(row=1, column=0, sticky=tk.W)

        ttk.Label(settings_tab, text="Тема интерфейса:").grid(row=9, column=0, sticky=tk.W, pady=(8, 3))
        self.theme_toggle_button = ttk.Button(settings_tab, text="", command=self.toggle_theme, width=14)
        self.theme_toggle_button.grid(row=9, column=1, sticky=tk.W, pady=(8, 3))

        ttk.Button(
            settings_tab,
            text="Сбросить состояние поиска",
            command=self.reset_search_state
        ).grid(row=10, column=1, sticky=tk.W, pady=(8, 3))

        footer_info = ttk.Label(
            settings_tab,
            text="Версия: v.2.0.6 | Автор: Андрей ОБИС 2026"
        )
        footer_info.grid(row=12, column=0, columnspan=2, sticky=(tk.W, tk.S), pady=(18, 0))

        self.setup_logging()
        self.apply_theme(self.theme_var.get())

    def toggle_theme(self):
        if self.theme_var.get() == "Светлая":
            self.theme_var.set("Темная")
        else:
            self.theme_var.set("Светлая")
        self.apply_theme(self.theme_var.get())
        self.config_dirty = True
        self.update_config(reload_after_save=False)

    def apply_theme(self, theme_name: str):
        style = ttk.Style()
        style.theme_use("clam")

        if theme_name == "Темная":
            colors = {
                "bg": "#1f1f1f",
                "panel": "#2a2a2a",
                "fg": "#e8e8e8",
                "entry_bg": "#333333",
                "accent": "#a855f7",
                "hover": "#4a3b5c",
                "btn_enabled_bg": "#3a3a3a",
                "btn_enabled_fg": "#f3f3f3",
                "btn_disabled_bg": "#262626",
                "btn_disabled_fg": "#7a7a7a",
            }
        else:
            colors = {
                "bg": "#f2f5fa",
                "panel": "#ffffff",
                "fg": "#1f2937",
                "entry_bg": "#ffffff",
                "accent": "#8b5cf6",
                "hover": "#efe7ff",
                "btn_enabled_bg": "#ffffff",
                "btn_enabled_fg": "#1f2937",
                "btn_disabled_bg": "#e5e7eb",
                "btn_disabled_fg": "#9ca3af",
            }

        self.root.configure(bg=colors["bg"])
        style.configure("TFrame", background=colors["bg"])
        style.configure("TLabel", background=colors["bg"], foreground=colors["fg"])
        style.configure("TButton", background=colors["btn_enabled_bg"], foreground=colors["btn_enabled_fg"])
        style.map(
            "TButton",
            background=[
                ("disabled", colors["btn_disabled_bg"]),
                ("active", colors["accent"]),
                ("!disabled", colors["btn_enabled_bg"]),
            ],
            foreground=[
                ("disabled", colors["btn_disabled_fg"]),
                ("active", "#ffffff"),
                ("!disabled", colors["btn_enabled_fg"]),
            ]
        )
        style.configure("TCheckbutton", background=colors["bg"], foreground=colors["fg"])
        style.configure("TNotebook", background=colors["bg"], borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            background=colors["panel"],
            foreground=colors["fg"],
            padding=(12, 7),
            font=("Segoe UI", 9)
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", colors["accent"]), ("!selected", colors["panel"])],
            foreground=[("selected", "#ffffff"), ("!selected", colors["fg"])],
            padding=[("selected", (12, 7)), ("!selected", (12, 7))],
            font=[("selected", ("Segoe UI", 9)), ("!selected", ("Segoe UI", 9))]
        )
        style.configure("Treeview", background=colors["entry_bg"], foreground=colors["fg"], fieldbackground=colors["entry_bg"])
        style.configure("Treeview.Heading", background=colors["panel"], foreground=colors["fg"])
        style.configure("TProgressbar", troughcolor=colors["panel"], background=colors["accent"])
        style.configure("TCombobox", fieldbackground=colors["entry_bg"], background=colors["panel"], foreground=colors["fg"])
        style.configure("DashboardWrap.TFrame", background=colors["bg"])
        style.configure("DashboardCard.TFrame", background=colors["panel"], borderwidth=1, relief="solid")
        style.configure("DashboardCardTitle.TLabel", background=colors["panel"], foreground=colors["fg"], font=("Segoe UI", 9))
        style.configure("DashboardCardValue.TLabel", background=colors["panel"], foreground=colors["accent"], font=("Segoe UI", 10, "bold"))
        style.configure("SessionCardsWrap.TFrame", background=colors["bg"])
        style.configure(
            "SessionCard.TFrame",
            background=colors["panel"],
            borderwidth=1,
            relief="solid"
        )
        style.configure(
            "SessionCardTitle.TLabel",
            background=colors["panel"],
            foreground=colors["fg"],
            font=("Segoe UI", 9)
        )
        style.configure(
            "SessionCardValue.TLabel",
            background=colors["panel"],
            foreground=colors["accent"],
            font=("Segoe UI", 10, "bold")
        )

        self.dirs_listbox.configure(bg=colors["entry_bg"], fg=colors["fg"], selectbackground=colors["accent"], selectforeground="#ffffff")
        self.keywords_text.configure(bg=colors["entry_bg"], fg=colors["fg"], insertbackground=colors["fg"])
        self.results_table.tag_configure('hover', background=colors["hover"])
        if self.theme_var.get() == "Темная":
            self.theme_toggle_button.config(text="🌙 Темная")
        else:
            self.theme_toggle_button.config(text="☀ Светлая")

    def setup_logging(self):
        """Настройка базового логирования приложения"""
        # Очищаем существующие обработчики
        logger = logging.getLogger()
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)
        logger.setLevel(logging.INFO)

    def build_search_signature(self, extensions, keywords_text, max_file_size):
        """Формирует подпись параметров поиска для проверки совместимости resume-сессии только по выбранным директориям."""
        signature_payload = {
            "directories": sorted(os.path.normcase(os.path.normpath(path)) for path in self.directories_list),
            "keywords_sha256": hashlib.sha256(keywords_text.encode("utf-8")).hexdigest(),
        }
        payload_json = json.dumps(signature_payload, ensure_ascii=False, sort_keys=True)
        return hashlib.sha256(payload_json.encode("utf-8")).hexdigest()

    def init_search_state_storage(self):
        """Инициализирует хранилище состояния поиска (SQLite-only)."""
        if sqlite3 is None:
            raise RuntimeError("Модуль sqlite3 недоступен в текущей среде Python")
        self._init_search_state_db()

    def _init_search_state_db(self):
        """Создает SQLite-таблицы для состояния поиска."""
        with sqlite3.connect(SEARCH_STATE_DB_FILE) as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS state_kv (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS processed_files (
                    path TEXT PRIMARY KEY
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS search_results (
                    path TEXT PRIMARY KEY,
                    keywords TEXT NOT NULL
                )
                """
            )
            conn.commit()

    def _load_found_results_from_db(self):
        """Загружает сохранённые совпадения из SQLite."""
        if sqlite3 is None or not os.path.exists(SEARCH_STATE_DB_FILE):
            return {}
        try:
            with sqlite3.connect(SEARCH_STATE_DB_FILE) as conn:
                rows = conn.execute("SELECT path, keywords FROM search_results").fetchall()
            found = {}
            for path, keywords_json in rows:
                try:
                    keywords = json.loads(keywords_json)
                except (TypeError, json.JSONDecodeError):
                    continue
                if isinstance(keywords, list):
                    found[path] = set(keywords)
            return found
        except Exception as exc:
            logging.warning(f"Не удалось загрузить search_results из {SEARCH_STATE_DB_FILE}: {exc}")
            return {}

    def _persist_found_result(self, file_path, keywords_set):
        """Сохраняет или обновляет одно совпадение в SQLite."""
        if sqlite3 is None:
            return
        try:
            with self.search_state_lock:
                with sqlite3.connect(SEARCH_STATE_DB_FILE) as conn:
                    conn.execute(
                        "INSERT OR REPLACE INTO search_results(path, keywords) VALUES(?, ?)",
                        (file_path, json.dumps(sorted(keywords_set), ensure_ascii=False)),
                    )
                    conn.commit()
        except Exception as exc:
            logging.error(f"Не удалось сохранить результат поиска {file_path}: {exc}")

    def _clear_found_results_db(self):
        """Удаляет все сохранённые совпадения."""
        if sqlite3 is None:
            return
        try:
            with sqlite3.connect(SEARCH_STATE_DB_FILE) as conn:
                conn.execute("DELETE FROM search_results")
                conn.commit()
        except Exception as exc:
            logging.error(f"Не удалось очистить search_results: {exc}")

    def load_search_state(self, include_processed_paths=False):
        """Читает состояние предыдущего поиска из SQLite."""
        return self._load_search_state_from_sqlite(include_processed_paths=include_processed_paths)

    def _load_processed_files_set(self):
        """Загружает множество уже обработанных путей из SQLite."""
        if sqlite3 is None or not os.path.exists(SEARCH_STATE_DB_FILE):
            return set()
        try:
            with sqlite3.connect(SEARCH_STATE_DB_FILE) as conn:
                return {row[0] for row in conn.execute("SELECT path FROM processed_files")}
        except Exception as exc:
            logging.warning(f"Не удалось загрузить processed_files из {SEARCH_STATE_DB_FILE}: {exc}")
            return set()

    def _load_search_state_from_sqlite(self, include_processed_paths=False):
        """Читает состояние из SQLite."""
        if sqlite3 is None or not os.path.exists(SEARCH_STATE_DB_FILE):
            return None
        try:
            with sqlite3.connect(SEARCH_STATE_DB_FILE) as conn:
                row = conn.execute(
                    "SELECT value FROM state_kv WHERE key = ?",
                    ("state_json",)
                ).fetchone()
                if not row:
                    return None
                state = json.loads(row[0])
                if not isinstance(state, dict):
                    return None
                if state.get("version") != SEARCH_STATE_VERSION:
                    return None
                processed_count_row = conn.execute("SELECT COUNT(*) FROM processed_files").fetchone()
                db_processed_count = int(processed_count_row[0]) if processed_count_row else 0
                state_processed_count = int(state.get("processed_count", 0))
                processed_count = max(state_processed_count, db_processed_count)
                state["processed_count"] = processed_count
                total_files = int(state.get("total_files", 0))
                state["remaining_count"] = max(0, total_files - processed_count)
                if include_processed_paths:
                    state["processed_files"] = [row[0] for row in conn.execute("SELECT path FROM processed_files")]
                else:
                    state["processed_files"] = []
                return state
        except Exception as exc:
            logging.warning(f"Не удалось прочитать {SEARCH_STATE_DB_FILE}: {exc}")
            return None

    def save_search_state(self):
        """Сохраняет текущее состояние поиска в SQLite."""
        with self.search_state_lock:
            if not self.active_search_state:
                return
            state_copy = dict(self.active_search_state)
        self._save_search_state_to_sqlite(state_copy)

    def _save_search_state_to_sqlite(self, state_copy):
        """Сохраняет состояние в SQLite."""
        try:
            with self.search_state_lock:
                pending_paths = list(self._unsaved_processed_paths)
                self._unsaved_processed_paths = []
                should_reset_table = self._reset_processed_files_table
                saved_processed_count = self._saved_processed_count
            state_for_db = dict(state_copy)
            state_for_db["processed_files"] = []
            with sqlite3.connect(SEARCH_STATE_DB_FILE) as conn:
                conn.execute("BEGIN")
                conn.execute(
                    "INSERT OR REPLACE INTO state_kv(key, value) VALUES(?, ?)",
                    ("state_json", json.dumps(state_for_db, ensure_ascii=False))
                )
                if should_reset_table:
                    conn.execute("DELETE FROM processed_files")
                    saved_processed_count = 0
                if pending_paths:
                    conn.executemany(
                        "INSERT OR REPLACE INTO processed_files(path) VALUES(?)",
                        [(path,) for path in pending_paths]
                    )
                    saved_processed_count += len(pending_paths)
                conn.commit()
            with self.search_state_lock:
                self._reset_processed_files_table = False
                self._saved_processed_count = saved_processed_count
        except Exception as exc:
            logging.error(f"Не удалось сохранить {SEARCH_STATE_DB_FILE}: {exc}")

    @staticmethod
    def format_session_status(status):
        status_map = {
            "running": "выполняется",
            "stopped": "остановлена",
            "completed": "завершена",
            "crashed": "завершилась с ошибкой",
        }
        return status_map.get(status, "нет данных")

    def update_search_session_labels(self, state):
        """Обновляет блок статуса сессии в интерфейсе."""
        if not state:
            self.search_session_status_var.set("нет данных")
            self.search_session_remaining_var.set("-")
            self.search_session_last_file_var.set("Последний файл: -")
            return

        status = self.format_session_status(state.get("status"))
        remaining = state.get("remaining_count", "-")
        last_file = state.get("last_processed_file") or "-"

        self.search_session_status_var.set(status)
        self.search_session_remaining_var.set(str(remaining))
        if len(last_file) > 120:
            last_file = "..." + last_file[-117:]
        self.search_session_last_file_var.set(f"Последний файл: {last_file}")

    def refresh_search_session_ui(self):
        """Подтягивает состояние поиска из памяти/файла и обновляет блок статуса."""
        state = None
        with self.search_state_lock:
            if self.active_search_state:
                state = dict(self.active_search_state)
        if state is None:
            state = self.load_search_state()
        self.found_results = self._load_found_results_from_db()
        self.update_search_session_labels(state)
        if state:
            self.reset_dashboard(
                state.get("total_files", 0),
                state.get("processed_count", 0),
                len(self.found_results),
                state.get("skipped_count", 0),
                state.get("error_count", 0),
            )
            skip_reasons = state.get("skip_reasons", {})
            if isinstance(skip_reasons, dict):
                for reason_key in self.dashboard_skip_reasons:
                    self.dashboard_skip_reasons[reason_key] = int(skip_reasons.get(reason_key, 0))
        else:
            self.reset_dashboard(0, 0, 0, 0, 0)

    def reset_dashboard(self, total_files, processed_count=0, found_count=0, skipped_count=0, error_count=0):
        """Сбрасывает счетчики дашборда перед запуском/возобновлением."""
        self.processed_files = max(0, int(processed_count))
        self.dashboard_found = max(0, int(found_count))
        self.dashboard_skipped = max(0, int(skipped_count))
        self.dashboard_errors = max(0, int(error_count))
        self.total_files = max(0, int(total_files))
        self.dashboard_skip_reasons = {
            "long_path": 0,
            "large_file": 0,
            "module_unavailable": 0,
            "read_error": 0,
            "other": 0,
        }
        self.update_dashboard_labels()

    def update_dashboard_labels(self):
        """Обновляет значения карточек дашборда."""
        self.dashboard_processed_var.set(str(self.processed_files))
        found_count = len(self.found_results) if self.found_results else self.dashboard_found
        self.dashboard_found_var.set(str(found_count))
        self.dashboard_skipped_var.set(str(self.dashboard_skipped))
        self.dashboard_errors_var.set(str(self.dashboard_errors))

    def show_found_results(self, _event=None):
        """Показывает все найденные файлы в таблице результатов."""
        self.found_results = self._load_found_results_from_db()
        self.dashboard_found = len(self.found_results)
        self.update_dashboard_labels()

        if not self.found_results:
            messagebox.showinfo("Найдено", "Совпадений пока нет.", parent=self.root)
            return

        for item in self.results_table.get_children():
            self.results_table.delete(item)
        self.hovered_result_item = None
        self.max_result_file_text_px = 0

        for file_path in sorted(self.found_results.keys()):
            keywords_str = ', '.join(sorted(self.found_results[file_path]))
            self._insert_live_result_row(keywords_str, file_path)

        children = self.results_table.get_children()
        if children:
            first_item = children[0]
            self.results_table.selection_set(first_item)
            self.results_table.focus(first_item)
            self.results_table.see(first_item)
            self.results_table.update_idletasks()

    def show_skip_details(self, _event=None):
        """Показывает детальную разбивку причин пропусков."""
        reason_names = {
            "long_path": "Слишком длинный путь",
            "large_file": "Слишком большой файл",
            "module_unavailable": "Модуль недоступен",
            "read_error": "Ошибка чтения",
            "other": "Другая причина",
        }
        lines = [
            f"{reason_names[key]}: {self.dashboard_skip_reasons.get(key, 0)}"
            for key in ("long_path", "large_file", "module_unavailable", "read_error", "other")
        ]
        details_text = "Детализация пропущенных файлов:\n\n" + "\n".join(lines)
        messagebox.showinfo("Причины пропуска", details_text, parent=self.root)

    def reset_search_state(self):
        """Удаляет сохраненное состояние поиска и сбрасывает UI-блок."""
        if self.is_searching:
            messagebox.showwarning("Нельзя выполнить", "Остановите текущий поиск перед сбросом состояния.")
            return

        try:
            self._init_search_state_db()
            with sqlite3.connect(SEARCH_STATE_DB_FILE, timeout=5) as conn:
                conn.execute("DELETE FROM state_kv")
                conn.execute("DELETE FROM processed_files")
                conn.execute("DELETE FROM search_results")
                conn.commit()
        except Exception as exc:
            messagebox.showerror("Ошибка", f"Не удалось очистить SQLite-состояние:\n{exc}")
            return

        with self.search_state_lock:
            self.active_search_state = None
            self.resume_processed_files = set()
            self.resume_start_count = 0
            self._pending_state_updates = 0
            self._reset_processed_files_table = False
            self._saved_processed_count = 0

        self.found_results = {}
        self.dashboard_found = 0
        self.refresh_search_session_ui()
        messagebox.showinfo("Готово", "Состояние поиска успешно сброшено.")

    def prepare_search_state(self, signature, total_files, resume_state=None):
        """Инициализирует состояние поиска (новое или продолжение)."""
        now_str = time.strftime('%Y-%m-%d %H:%M:%S')
        self._pending_state_updates = 0
        if resume_state:
            processed_count = int(resume_state.get("processed_count", 0))
            self._reset_processed_files_table = False
            self._saved_processed_count = processed_count
            self._unsaved_processed_paths = []
            self.found_results = self._load_found_results_from_db()
            resolved_total = max(int(total_files), int(resume_state.get("total_files", total_files)))
            self.active_search_state = {
                **{key: value for key, value in resume_state.items() if key != "processed_files"},
                "version": SEARCH_STATE_VERSION,
                "signature": signature,
                "status": "running",
                "updated_at": now_str,
                "last_error": "",
                "processed_files": [],
                "processed_count": processed_count,
                "matched_count": int(resume_state.get("matched_count", 0)),
                "skipped_count": int(resume_state.get("skipped_count", 0)),
                "error_count": int(resume_state.get("error_count", 0)),
                "skip_reasons": dict(resume_state.get("skip_reasons", {})) if isinstance(resume_state.get("skip_reasons", {}), dict) else {},
                "remaining_count": max(0, resolved_total - processed_count),
                "total_files": resolved_total,
            }
        else:
            self._reset_processed_files_table = True
            self._saved_processed_count = 0
            self._unsaved_processed_paths = []
            self._clear_found_results_db()
            self.found_results = {}
            self.active_search_state = {
                "version": SEARCH_STATE_VERSION,
                "status": "running",
                "created_at": now_str,
                "updated_at": now_str,
                "signature": signature,
                "total_files": int(total_files),
                "processed_count": 0,
                "matched_count": 0,
                "skipped_count": 0,
                "error_count": 0,
                "skip_reasons": dict(self.dashboard_skip_reasons),
                "remaining_count": int(total_files),
                "last_processed_file": "",
                "last_error": "",
                "processed_files": [],
            }
        self.save_search_state()
        self.safe_after(0, self.refresh_search_session_ui)

    def update_search_state_checkpoint(self, file_path, had_matches, error_text="", file_status="no_match", skip_reason=""):
        """Обновляет состояние после обработки файла."""
        should_flush = False
        with self.search_state_lock:
            if not self.active_search_state:
                return

            processed_files = self.active_search_state.setdefault("processed_files", [])
            if file_path not in self.resume_processed_files:
                self.resume_processed_files.add(file_path)
                self._unsaved_processed_paths.append(file_path)
                if len(processed_files) < 32:
                    processed_files.append(file_path)

            self.active_search_state["processed_count"] = len(self.resume_processed_files)
            if had_matches:
                self.active_search_state["matched_count"] = int(self.active_search_state.get("matched_count", 0)) + 1
            if file_status == "skipped":
                self.active_search_state["skipped_count"] = int(self.active_search_state.get("skipped_count", 0)) + 1
                normalized_reason = skip_reason if skip_reason in self.dashboard_skip_reasons else "other"
                self.dashboard_skip_reasons[normalized_reason] = int(self.dashboard_skip_reasons.get(normalized_reason, 0)) + 1
            if error_text:
                self.active_search_state["error_count"] = int(self.active_search_state.get("error_count", 0)) + 1
                self.active_search_state["last_error"] = error_text
            self.active_search_state["last_processed_file"] = file_path
            total_files = int(self.active_search_state.get("total_files", self.total_files))
            self.active_search_state["remaining_count"] = max(0, total_files - len(self.resume_processed_files))
            self.active_search_state["updated_at"] = time.strftime('%Y-%m-%d %H:%M:%S')
            self.active_search_state["skip_reasons"] = dict(self.dashboard_skip_reasons)
            matched_count = int(self.active_search_state.get("matched_count", 0))
            skipped_count = int(self.active_search_state.get("skipped_count", 0))
            error_count = int(self.active_search_state.get("error_count", 0))
            self._pending_state_updates += 1
            should_flush = self._pending_state_updates >= self._state_flush_interval

        self.dashboard_found = matched_count
        self.dashboard_skipped = skipped_count
        self.dashboard_errors = error_count
        if should_flush:
            self.save_search_state()
            with self.search_state_lock:
                self._pending_state_updates = 0
        self.safe_after(0, self.refresh_search_session_ui)
        self.safe_after(0, self.update_dashboard_labels)

    def finalize_search_state(self, status, error_text=""):
        """Фиксирует финальный статус сессии поиска."""
        with self.search_state_lock:
            if not self.active_search_state:
                return
            self.active_search_state["status"] = status
            self.active_search_state["updated_at"] = time.strftime('%Y-%m-%d %H:%M:%S')
            self.active_search_state["processed_count"] = len(self.resume_processed_files)
            total_files = int(self.active_search_state.get("total_files", self.total_files))
            self.active_search_state["remaining_count"] = max(0, total_files - len(self.resume_processed_files))
            if error_text:
                self.active_search_state["last_error"] = error_text
            self._pending_state_updates = 0
        self.save_search_state()
        self.safe_after(0, self.refresh_search_session_ui)

    def mark_search_crashed(self, error_title, exc_value):
        """Помечает текущую сессию как аварийно завершенную."""
        crash_reason = f"{error_title}: {exc_value}"
        self.finalize_search_state("crashed", crash_reason)

    def ensure_search_file_logging(self, log_file='search_log.txt'):
        """Гарантирует ровно один файловый обработчик логов для поиска."""
        logger = logging.getLogger()
        if self.search_file_handler is not None:
            try:
                logger.removeHandler(self.search_file_handler)
                self.search_file_handler.close()
            except Exception:
                pass
            self.search_file_handler = None

        file_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
        file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(file_handler)
        self.search_file_handler = file_handler

    def safe_after(self, delay_ms, callback, *args, allow_when_closing=False, **kwargs):
        """Безопасно планирует вызов в UI-потоке и не дает падать из-за after/callback."""
        if self.is_closing and not allow_when_closing:
            return False

        def wrapped():
            try:
                callback(*args, **kwargs)
            except Exception as exc:
                self.report_runtime_error(
                    f"Ошибка интерфейса в {getattr(callback, '__name__', 'callback')}",
                    exc,
                    show_dialog=True
                )

        try:
            self.root.after(delay_ms, wrapped)
            return True
        except (RuntimeError, tk.TclError) as exc:
            if not self.is_closing:
                self.report_runtime_error("Не удалось запланировать обновление интерфейса", exc, show_dialog=False)
            return False

    def _show_runtime_error_dialog(self, error_title, error_message):
        """Показывает пользователю подробности ошибки, если окно еще активно."""
        if self.is_closing or not self.root.winfo_exists():
            return
        messagebox.showerror(error_title, error_message, parent=self.root)

    def report_runtime_error(self, error_title, exc, show_dialog=True):
        """Сохраняет ошибку, пишет ее в лог и сообщает пользователю без остановки приложения."""
        if self.is_closing and isinstance(exc, (RuntimeError, tk.TclError)):
            return

        log_path = write_crash_report(error_title, exc, exc.__traceback__)
        logging.exception("%s: %s", error_title, exc)

        short_message = f"{error_title}: {exc}"
        try:
            self.safe_after(0, self.current_file.set, short_message[:300])
        except Exception:
            pass

        if show_dialog and not self.is_closing:
            error_message = (
                f"{error_title}\n\n"
                f"{exc}\n\n"
                f"Подробности сохранены в файле:\n{log_path}\n\n"
                "Приложение продолжит работу, если это возможно."
            )
            try:
                self.safe_after(0, self._show_runtime_error_dialog, "Ошибка", error_message)
            except Exception:
                pass

    def add_directory(self):
        """Добавление директории для поиска"""
        if self.is_adding_directory:
            return

        directory = filedialog.askdirectory(title="Выберите директорию для поиска")
        if not directory or directory in self.directories_list:
            return

        self.is_adding_directory = True
        self.start_button.config(state=tk.DISABLED)
        self.adding_dir_frame.grid()
        self.adding_dir_spinner.start(10)
        worker = threading.Thread(
            target=self._add_directory_worker,
            args=(directory,),
            daemon=True,
            name="directory-add-worker"
        )
        self.directory_add_thread = worker
        worker.start()

    def _add_directory_worker(self, directory):
        """Фоновая подготовка директории перед добавлением в UI."""
        try:
            normalized_directory = os.path.normpath(directory)
            self.safe_after(0, self._finish_add_directory, normalized_directory)
        except Exception as e:
            self.safe_after(0, self._finish_add_directory, None, error=e)

    def _finish_add_directory(self, directory, error=None):
        """Завершение добавления директории в основном потоке."""
        try:
            if error:
                messagebox.showerror("Ошибка", f"Не удалось добавить директорию:\n{error}")
                return

            if directory and directory not in self.directories_list:
                self.directories_list.append(directory)
                self.dirs_listbox.insert(tk.END, directory)
                # Помечаем конфиг измененным; сохраним позже, чтобы не тормозить UI
                self.config_dirty = True
        finally:
            self.adding_dir_spinner.stop()
            self.adding_dir_frame.grid_remove()
            self.is_adding_directory = False
            self.directory_add_thread = None
            if not self.is_searching:
                self.start_button.config(state=tk.NORMAL)

    def remove_directory(self):
        """Удаление выбранной директории"""
        selection = self.dirs_listbox.curselection()
        if selection:
            index = selection[0]
            self.dirs_listbox.delete(index)
            del self.directories_list[index]
            # Помечаем конфиг измененным; сохраним позже, чтобы не тормозить UI
            self.config_dirty = True

    def clear_all(self):
        """Очистка всех полей"""
        for item in self.results_table.get_children():
            self.results_table.delete(item)
        self.hovered_result_item = None
        self.max_result_file_text_px = 0
        self.results_table.column("file", width=850)
        self._update_results_h_scrollbar_visibility()

        self.progress_value.set(0)
        self.current_file.set("")
        self.processed_files = 0

    def update_progress(self, file_name="", in_flight=None):
        """Обновление прогресса с информацией о прогрессе"""
        if in_flight is not None:
            self.search_in_flight_count = max(0, int(in_flight))

        logging.debug(
            f"Updating progress: {self.processed_files}/{self.total_files}, "
            f"file: {file_name}, in_flight: {self.search_in_flight_count}"
        )

        if self.total_files > 0:
            progress = (self.processed_files / self.total_files) * 100
            self.progress_value.set(progress)

            progress_text = f"Обработано: {self.processed_files}/{self.total_files} файлов"
            if self.search_in_flight_count > 0:
                progress_text += f" | в работе: {self.search_in_flight_count}"

            if file_name:
                display_name = file_name
                if len(file_name) > 50:
                    display_name = "..." + file_name[-47:]

                if file_name.startswith("Завершена обработка:"):
                    progress_text += f" | {file_name}"
                elif file_name.startswith("Начат:"):
                    progress_text += f" | {file_name.replace('Начат:', 'Текущий:', 1)}"
                elif file_name.startswith("Подготовка:"):
                    progress_text += f" | {file_name}"
                elif file_name.startswith("Запуск поиска:"):
                    progress_text += f" | {file_name}"
                elif file_name.startswith("Готово:"):
                    progress_text += f" | {file_name}"
                elif file_name == "Поиск завершен":
                    progress_text = "Поиск завершен! Обработано всех файлов."
                else:
                    progress_text += f" | {display_name}"

            self.current_file.set(progress_text)
        else:
            status_text = f"Обработано: {self.processed_files} файлов"
            if self.search_in_flight_count > 0:
                status_text += f" | в работе: {self.search_in_flight_count}"
            if file_name:
                if file_name.startswith("Завершена обработка:"):
                    status_text += f" | {file_name}"
                elif file_name.startswith("Начат:"):
                    status_text += f" | {file_name.replace('Начат:', 'Текущий:', 1)}"
                elif file_name.startswith("Готово:"):
                    status_text += f" | {file_name}"
                elif file_name == "Поиск завершен":
                    status_text = f"Поиск завершен! Обработано файлов: {self.processed_files}"
                else:
                    status_text += f" | {file_name}"
            self.current_file.set(status_text)

        # Принудительно обновляем прогрессбар
        self.progress_bar.update_idletasks()

    def add_live_result(self, file_path, keywords):
        """Добавление найденного результата в таблицу в реальном времени."""
        keywords_set = set(keywords) if keywords else set()
        if file_path in self.found_results:
            self.found_results[file_path].update(keywords_set)
        else:
            self.found_results[file_path] = keywords_set
        self.dashboard_found = len(self.found_results)
        self._persist_found_result(file_path, self.found_results[file_path])
        keywords_str = ', '.join(sorted(self.found_results[file_path]))
        self.safe_after(0, self._upsert_live_result_row, keywords_str, file_path)
        self.safe_after(0, self.update_dashboard_labels)

    def _upsert_live_result_row(self, keywords_str, file_path):
        """Добавляет или обновляет строку результата для указанного пути."""
        for item_id in self.results_table.get_children():
            values = self.results_table.item(item_id, "values")
            if len(values) >= 2 and str(values[1]) == file_path:
                self.results_table.item(item_id, values=(keywords_str, file_path))
                self._update_file_column_width_for_path(file_path)
                return
        self._insert_live_result_row(keywords_str, file_path)

    def _insert_live_result_row(self, keywords_str, file_path):
        """Вставляет строку результата и подстраивает ширину колонки ссылки."""
        self.results_table.insert('', tk.END, values=(keywords_str, file_path))
        self._update_file_column_width_for_path(file_path)

    def _update_file_column_width_for_path(self, file_path):
        """Расширяет колонку ссылки по фактической длине текста."""
        try:
            table_font = tkfont.nametofont("TkDefaultFont")
            text_px = table_font.measure(str(file_path)) + 40
            if text_px > self.max_result_file_text_px:
                self.max_result_file_text_px = text_px
                new_width = max(850, min(6000, self.max_result_file_text_px))
                self.results_table.column("file", width=new_width)
            self._update_results_h_scrollbar_visibility()
        except Exception:
            pass

    def _update_results_h_scrollbar_visibility(self):
        """Показывает нижний скролл только если ссылка выходит за видимую ширину."""
        if not hasattr(self, "results_h_scrollbar"):
            return
        table_width = self.results_table.winfo_width()
        if table_width <= 1:
            return
        total_columns_width = int(self.results_table.column("keywords", "width")) + int(self.results_table.column("file", "width"))
        if total_columns_width > table_width:
            self.results_h_scrollbar.grid()
        else:
            self.results_h_scrollbar.grid_remove()

    def on_results_table_resize(self, event):
        """При ресайзе обновляет видимость горизонтального скролла."""
        _ = event
        self._update_results_h_scrollbar_visibility()

    def open_result_file(self, event):
        """Открытие файла из выбранной строки таблицы по двойному клику."""
        try:
            item_id = self.results_table.identify_row(event.y)
            if not item_id:
                return

            self.results_table.selection_set(item_id)
            self._open_file_by_path(self._get_result_file_path(item_id))
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{e}")

    def _get_result_file_path(self, item_id=None):
        """Возвращает путь к файлу из выбранной строки таблицы."""
        if item_id is None:
            selection = self.results_table.selection()
            item_id = selection[0] if selection else ""

        if not item_id:
            return ""

        values = self.results_table.item(item_id, "values")
        if not values or len(values) < 2:
            return ""

        return str(values[1]).strip()

    @staticmethod
    def split_result_path(file_path):
        """Разделяет путь к файлу внутри архива и путь к самому архиву."""
        if ARCHIVE_MEMBER_SEP in file_path:
            archive_path, member_path = file_path.split(ARCHIVE_MEMBER_SEP, 1)
            return archive_path.strip(), member_path.strip()
        return file_path, None

    def _extract_archive_member(self, archive_path, member_path):
        """Извлекает файл из архива во временную папку и возвращает путь к копии."""
        member_path = member_path.replace("\\", "/")
        archive_ext = os.path.splitext(archive_path)[1].lower()
        temp_dir = tempfile.mkdtemp(prefix="zsearch_open_")

        if archive_ext == '.zip':
            with zipfile.ZipFile(archive_path, 'r') as archive:
                archive.extract(member_path, temp_dir)
        elif archive_ext == '.rar':
            import rarfile
            with rarfile.RarFile(archive_path, 'r') as archive:
                archive.extract(member_path, temp_dir)
        elif archive_ext == '.7z':
            import py7zr
            with py7zr.SevenZipFile(archive_path, mode='r') as archive:
                archive.extract(targets=[member_path.replace("/", os.sep)], path=temp_dir)
        else:
            shutil.rmtree(temp_dir, ignore_errors=True)
            raise ValueError(f"Неподдерживаемый тип архива: {archive_path}")

        extracted_path = os.path.normpath(os.path.join(temp_dir, member_path.replace("/", os.sep)))
        if not os.path.isfile(extracted_path):
            shutil.rmtree(temp_dir, ignore_errors=True)
            raise FileNotFoundError(f"Файл не найден внутри архива: {member_path}")
        return extracted_path

    def _open_file_by_path(self, file_path):
        """Открывает файл в системе."""
        if not file_path:
            messagebox.showwarning("Файл не найден", "Путь к файлу не указан.")
            return

        archive_path, member_path = self.split_result_path(file_path)
        if member_path:
            if not os.path.exists(archive_path):
                messagebox.showwarning("Файл не найден", f"Архив не существует:\n{archive_path}")
                return
            try:
                extracted_path = self._extract_archive_member(archive_path, member_path)
            except Exception as exc:
                messagebox.showerror("Ошибка", f"Не удалось извлечь файл из архива:\n{exc}")
                return
            file_path = extracted_path
        elif not os.path.exists(file_path):
            messagebox.showwarning("Файл не найден", f"Файл не существует:\n{file_path}")
            return

        if hasattr(os, "startfile"):
            os.startfile(file_path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", file_path])
        else:
            subprocess.Popen(["xdg-open", file_path])

    def open_selected_result_file(self):
        """Открывает файл из выбранной строки таблицы."""
        try:
            self._open_file_by_path(self._get_result_file_path())
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{e}")

    def open_result_location(self):
        """Открывает папку с файлом из выбранной строки таблицы."""
        try:
            file_path = self._get_result_file_path()
            if not file_path:
                return

            archive_path, member_path = self.split_result_path(file_path)
            target_path = archive_path if member_path else file_path
            if not os.path.exists(target_path):
                messagebox.showwarning("Файл не найден", f"Файл не существует:\n{target_path}")
                return

            if sys.platform.startswith("win"):
                subprocess.Popen(["explorer", "/select,", os.path.normpath(target_path)])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", target_path])
            else:
                subprocess.Popen(["xdg-open", os.path.dirname(target_path) or "."])
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть расположение файла:\n{e}")

    def show_results_context_menu(self, event):
        """Показывает контекстное меню для строки таблицы результатов."""
        item_id = self.results_table.identify_row(event.y)
        if not item_id:
            return

        self.results_table.selection_set(item_id)
        self.results_table.focus(item_id)
        try:
            self.results_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.results_context_menu.grab_release()

    def on_results_hover(self, event):
        """Подсветка строки таблицы при наведении."""
        item_id = self.results_table.identify_row(event.y)
        if item_id == self.hovered_result_item:
            return

        if self.hovered_result_item:
            self.results_table.item(self.hovered_result_item, tags=())

        self.hovered_result_item = item_id if item_id else None
        if self.hovered_result_item:
            self.results_table.item(self.hovered_result_item, tags=('hover',))

    def on_results_leave(self, _event):
        """Сброс подсветки строки, когда курсор покинул таблицу."""
        if self.hovered_result_item:
            self.results_table.item(self.hovered_result_item, tags=())
            self.hovered_result_item = None

    def collect_files_to_process(self, directory, extensions):
        """Собирает список файлов для обработки в директории."""
        collected = []
        for root, _, files in os.walk(directory):
            for file in files:
                file_path = os.path.join(root, file)
                # Используем fnmatch для проверки соответствия расширениям
                if any(fnmatch.fnmatch(file, ext) for ext in extensions):
                    collected.append(file_path)
        return collected

    def get_selected_extensions(self):
        """Возвращает список расширений, выбранных галочками."""
        return [ext for ext in self.extension_options if self.extension_vars[ext].get()]

    def start_search(self):
        """Запуск поиска в отдельном потоке"""
        if self.is_searching:
            return
        if self.is_adding_directory:
            messagebox.showwarning("Подождите", "Дождитесь завершения добавления директории.")
            return

        # Автоматически выставляем потоки по формуле: max_cpu-2, иначе 1
        self.threads_var.set(str(self.get_auto_threads_count()))
        self.normalize_threads_value()
        self.normalize_ocr_threads_value()

        extensions = self.get_selected_extensions()

        if not extensions:
            messagebox.showerror("Ошибка", "Не выбрано ни одного расширения файлов!")
            return

        # Получаем ключевые слова
        keywords = self.keywords_text.get("1.0", tk.END).strip()
        if not keywords:
            messagebox.showerror("Ошибка", "Не введены ключевые слова для поиска!")
            return

        # Сохраняем ключевые слова в файл
        try:
            with open("keywords.txt", "w", encoding="utf-8") as f:
                f.write(keywords)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить ключевые слова: {e}")
            return

        # Проверяем директории
        if not self.directories_list:
            messagebox.showerror("Ошибка", "Не выбрано ни одной директории для поиска!")
            return

        max_size_raw = str(self.max_size_var.get()).strip()
        if max_size_raw == "":
            max_file_size = 0
        else:
            try:
                max_file_size = int(max_size_raw)
            except ValueError:
                messagebox.showerror("Ошибка", "Некорректный лимит размера файла (МБ).")
                return
        if max_file_size < 0:
            messagebox.showerror("Ошибка", "Макс. размер файла не может быть отрицательным.")
            return

        # В зависимости от режима либо считаем файлы заранее, либо запускаем сразу.
        pre_count_enabled = bool(self.pre_count_files_var.get())
        calculated_total = 0
        self.directory_files_map = {}
        if pre_count_enabled:
            for directory in self.directories_list:
                files_in_directory = self.collect_files_to_process(directory, extensions)
                self.directory_files_map[directory] = files_in_directory
                calculated_total += len(files_in_directory)

        signature = self.build_search_signature(extensions, keywords, max_file_size)
        resume_state = None
        self.resume_processed_files = set()
        self.resume_start_count = 0
        self.total_files = calculated_total if pre_count_enabled else 0
        resume_matched_count = 0
        resume_skipped_count = 0
        resume_error_count = 0
        resume_skip_reasons = {}

        previous_state = self.load_search_state(include_processed_paths=False)
        if (
            previous_state
            and previous_state.get("status") in ("running", "stopped", "crashed")
            and previous_state.get("signature") == signature
        ):
            previous_processed_count = int(previous_state.get("processed_count", 0))
            previous_total = max(int(previous_state.get("total_files", 0)), calculated_total)
            remaining = max(0, previous_total - previous_processed_count)
            resume_choice = messagebox.askyesnocancel(
                "Найден незавершенный поиск",
                (
                    "Найдена предыдущая незавершенная сессия поиска.\n\n"
                    f"Уже обработано: {previous_processed_count}\n"
                    f"Осталось: {remaining}\n"
                    f"Статус прошлой сессии: {previous_state.get('status', 'unknown')}\n\n"
                    "Продолжить с прошлого места?"
                ),
                parent=self.root
            )
            if resume_choice is None:
                return
            if resume_choice:
                resume_state = previous_state
                self.resume_paths_pending_load = True
                self.resume_start_count = previous_processed_count
                self.total_files = previous_total
                resume_matched_count = int(previous_state.get("matched_count", 0))
                resume_skipped_count = int(previous_state.get("skipped_count", 0))
                resume_error_count = int(previous_state.get("error_count", 0))
                if isinstance(previous_state.get("skip_reasons"), dict):
                    resume_skip_reasons = dict(previous_state.get("skip_reasons"))
                logging.info(
                    f"Возобновление прошлой сессии: обработано {self.resume_start_count}, "
                    f"осталось {max(0, self.total_files - self.resume_start_count)}"
                )

        if pre_count_enabled and self.total_files == 0:
            messagebox.showwarning("Предупреждение", "Не найдено файлов для обработки в указанных директориях!")
            return

        # Очищаем результаты и лог
        self.clear_all()

        # Записываем время начала поиска
        self.search_start_time = time.strftime('%Y-%m-%d %H:%M:%S')
        start_message = f"Поиск начат: {self.search_start_time}"
        logging.info(start_message)
        self.prepare_search_state(signature, self.total_files, resume_state=resume_state)
        self.reset_dashboard(
            self.total_files,
            self.resume_start_count,
            len(self.found_results),
            resume_skipped_count,
            resume_error_count,
        )
        if resume_skip_reasons:
            for reason_key in self.dashboard_skip_reasons:
                self.dashboard_skip_reasons[reason_key] = int(resume_skip_reasons.get(reason_key, 0))

        # Обновляем конфиг
        self.update_config()

        # Добавляем информацию о доступности модулей в конфиг
        self.config['config']['has_pdf'] = HAS_PDF
        self.config['config']['has_docx'] = HAS_DOCX
        self.config['config']['has_excel'] = HAS_EXCEL
        self.config['config']['has_7z'] = HAS_7Z
        self.config['config']['has_rar'] = HAS_RAR
        self.config['config']['has_ocr'] = HAS_OCR
        self.config['config']['ocr_threads'] = self.normalize_ocr_threads_value()

        # Настраиваем логирование с защитой от дублирующихся обработчиков
        try:
            self.ensure_search_file_logging('search_log.txt')
        except Exception as e:
            logging.error(f"Не удалось настроить файловое логирование: {e}")

        # Загружаем ключевые слова
        try:
            load_keywords("keywords.txt")
        except ValueError as e:
            self.finalize_search_state("stopped", str(e))
            messagebox.showerror("Ошибка", str(e))
            return

        # Меняем состояние кнопок
        self.start_button.config(state=tk.DISABLED)
        self.pause_button.config(state=tk.NORMAL, text="Пауза")
        self.stop_button.config(state=tk.NORMAL)
        self.is_searching = True
        self.is_paused = False
        self.processed_files = self.resume_start_count
        self.search_in_flight_count = 0
        if pre_count_enabled:
            self.progress_bar.config(mode="determinate")
            self.progress_value.set(0)
        else:
            self.progress_bar.config(mode="indeterminate")
            self.progress_bar.start(10)
            self.current_file.set(f"Обработано: {self.processed_files} файлов | Поиск запущен без предварительного подсчета")

        # Запускаем поиск в отдельном потоке
        self.search_thread = threading.Thread(
            target=self.run_search,
            args=(extensions, self.update_progress_callback),  # Передаем callback
            name="search-worker"
        )
        self.search_thread.daemon = True
        self._begin_search_preparation_ui()
        self.search_thread.start()

    def _begin_search_preparation_ui(self):
        """Показывает, что поиск запускается после подсчёта файлов."""
        self.current_file.set("Запуск поиска: подготовка очереди файлов...")
        self.progress_bar.update_idletasks()

    def stop_search(self):
        """Остановка поиска"""
        if self.is_searching:
            self.is_searching = False
            self.is_paused = False
            # Не включаем "Старт" до полного завершения фонового потока
            self.start_button.config(state=tk.DISABLED)
            self.pause_button.config(state=tk.DISABLED, text="Пауза")
            self.stop_button.config(state=tk.DISABLED)
            logging.info("Запрошена остановка поиска пользователем")

            # Обновляем статус: процесс еще завершает текущие задачи
            self.current_file.set("Останавливаем поиск...")

    def toggle_pause(self):
        """Пауза/продолжение поиска."""
        if not self.is_searching:
            return

        self.is_paused = not self.is_paused
        if self.is_paused:
            self.pause_button.config(text="Продолжить")
            self.current_file.set("Поиск на паузе")
            logging.info("Поиск поставлен на паузу")
        else:
            self.pause_button.config(text="Пауза")
            logging.info("Поиск продолжен")

    def run_search(self, extensions, progress_callback):
        """Выполнение поиска"""
        search_completed = False
        critical_error = ""
        try:
            if self.resume_paths_pending_load:
                self.safe_after(0, self.update_progress, "Подготовка: загрузка состояния возобновления...")
                self.resume_processed_files = self._load_processed_files_set()
                self.resume_paths_pending_load = False
                self.processed_files = len(self.resume_processed_files)
                self.safe_after(0, self.update_dashboard_labels)

            self.processed_files = self.resume_start_count
            threads_count = self.normalize_threads_value()
            logging.info(f"Начинаем поиск. Всего файлов: {self.total_files}")

            # Выполняем поиск для каждой директории с накоплением счетчика
            for directory in self.directories_list:
                if not self.is_searching:
                    logging.info("Поиск остановлен пользователем")
                    break

                try:
                    logging.info(f"Начинаем поиск в директории: {directory}")

                    # Обновляем статус - начало обработки директории
                    self.safe_after(0, self.update_progress, f"Начата обработка: {os.path.basename(directory)}")

                    # Используем модифицированную функцию поиска с прогрессом
                    results = search_files(
                        directory,
                        extensions,
                        threads_count,
                        "search_results.txt",
                        int(self.max_size_var.get()),
                        self.config['config'],
                        progress_callback,
                        self.processed_files,  # Передаем текущее значение как offset
                        lambda: self.is_searching,
                        self.add_live_result,
                        lambda: self.is_paused,
                        self.resume_processed_files,
                        self.update_search_state_checkpoint,
                        self.directory_files_map.get(directory)
                    )

                    # Показываем результаты для текущей директории
                    if results:
                        logging.info(f"Найдено совпадений в {len(results)} файлах в директории {directory}:")
                        for file_path, keywords in results.items():
                            logging.info(f"Файл: {file_path}")
                            logging.info(f"Ключевые слова: {', '.join(keywords)}")
                    else:
                        logging.info(f"В директории {directory} ничего не найдено.")

                    # Обновляем прогресс после обработки каждой директории
                    if self.is_searching:
                        self.safe_after(0, self.update_progress, f"Завершена обработка: {os.path.basename(directory)}")
                except Exception as directory_error:
                    self.report_runtime_error(
                        f"Ошибка при обработке директории {directory}",
                        directory_error,
                        show_dialog=True
                    )
                    continue

            # Только после ВСЕХ директорий показываем завершение поиска
            if self.is_searching:
                logging.info("Поиск завершен!")
                self.safe_after(0, self.update_progress, "Поиск завершен")
                search_completed = True

        except Exception as e:
            critical_error = str(e)
            self.report_runtime_error("Критическая ошибка во время поиска", e, show_dialog=True)

        finally:
            was_searching = self.is_searching
            # Записываем время окончания поиска
            self.search_end_time = time.strftime('%Y-%m-%d %H:%M:%S')
            if self.is_searching:
                end_message = f"Поиск завершен: {self.search_end_time}"
            else:
                end_message = f"Поиск остановлен: {self.search_end_time}"
            logging.info(end_message)

            # Добавляем информацию о продолжительности поиска
            if self.search_start_time:
                try:
                    start_time_obj = time.strptime(self.search_start_time, '%Y-%m-%d %H:%M:%S')
                    end_time_obj = time.strptime(self.search_end_time, '%Y-%m-%d %H:%M:%S')
                    start_seconds = time.mktime(start_time_obj)
                    end_seconds = time.mktime(end_time_obj)
                    duration = end_seconds - start_seconds
                    hours = int(duration // 3600)
                    minutes = int((duration % 3600) // 60)
                    seconds = int(duration % 60)
                    duration_message = f"Продолжительность поиска: {hours:02d}:{minutes:02d}:{seconds:02d}"
                    logging.info(duration_message)
                except ValueError:
                    pass

            if search_completed:
                self.finalize_search_state("completed")
            elif was_searching and not self.is_closing:
                self.finalize_search_state("crashed", critical_error)
            else:
                self.finalize_search_state("stopped")

            self.is_searching = False
            self.is_paused = False
            self.safe_after(0, self.on_search_finished)

    def update_progress_callback(self, file_name, processed_count, in_flight=None):
        """Callback для обновления прогресса из search_engine"""
        self.safe_after(0, self._update_progress_in_main_thread, file_name, processed_count, in_flight)

    def _update_progress_in_main_thread(self, file_name, processed_count, in_flight=None):
        """Обновление прогресса в основном потоке"""
        if isinstance(processed_count, int):
            self.processed_files = processed_count
            self.update_dashboard_labels()
        self.update_progress(file_name, in_flight=in_flight)

    def on_search_finished(self):
        """Вызывается при завершении поиска"""
        self.progress_bar.stop()
        self.progress_bar.config(mode="determinate")
        self.search_in_flight_count = 0
        self.directory_files_map = {}
        self.start_button.config(state=tk.NORMAL)
        self.pause_button.config(state=tk.DISABLED, text="Пауза")
        self.stop_button.config(state=tk.DISABLED)
        self.update_dashboard_labels()
        # Если остановка была пользователем, финализируем человекочитаемый статус
        if self.current_file.get() == "Останавливаем поиск...":
            self.current_file.set("Поиск остановлен пользователем")

    def update_config(self, reload_after_save=True):
        """Обновление конфигурации"""
        config = ConfigParser()
        extensions = self.get_selected_extensions()

        current_cfg = self.config['config']
        threads_count = self.normalize_threads_value()
        ocr_threads_count = self.normalize_ocr_threads_value()
        max_size_to_save = str(self.max_size_var.get()).strip() or "0"
        max_path_to_save = str(self.max_path_length_var.get()).strip() or "0"

        # Обновляем конфиг
        config['Settings'] = {
            'extensions': ', '.join(extensions),
            'keywords_file': current_cfg.get('keywords_file', 'keywords.txt'),
            'directories': '; '.join(self.directories_list) if self.directories_list else '.',
            'directory': self.directories_list[0] if self.directories_list else current_cfg.get('directory', '.'),
            'theme': self.theme_var.get(),
            'pre_count_files': 'true' if self.pre_count_files_var.get() else 'false',
            'threads': str(threads_count),
            'ocr_threads': str(ocr_threads_count),
            'output_file': current_cfg.get('output_file', 'search_results.txt'),
            'search_images': 'true' if self.search_images_var.get() else 'false',
            'max_file_size': max_size_to_save,
            'max_path_length': max_path_to_save,
            'log_file': current_cfg.get('log_file', 'search_log.txt'),
            'tesseract_languages': current_cfg.get('tesseract_languages', 'rus'),
            'tesseract_config': current_cfg.get('tesseract_config', '--oem 3 --psm 6'),
            # Новые параметры для тонкой настройки поиска
            'max_pdf_pages': self.max_pdf_pages_var.get(),
        }

        # Сохраняем конфиг
        with open('config.txt', 'w', encoding='utf-8') as configfile:
            config.write(configfile)
        self.config_dirty = False

        # Обновляем self.config без сброса значений в интерфейсе
        if reload_after_save:
            new_config = load_config()
            self.config['config'] = new_config

    def save_keywords_to_file(self):
        """Сохраняет текущий текст ключевых слов в keywords.txt."""
        try:
            keywords = self.keywords_text.get("1.0", tk.END).strip()
            with open("keywords.txt", "w", encoding="utf-8") as f:
                f.write(keywords)
        except Exception as exc:
            logging.error(f"Не удалось сохранить keywords.txt при закрытии: {exc}")

    def on_close(self):
        """Корректное завершение приложения с остановкой фоновых потоков."""
        if self.is_closing:
            return

        self.is_closing = True
        self.is_searching = False
        self.is_paused = False
        self.start_button.config(state=tk.DISABLED)
        self.pause_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.DISABLED)
        self.current_file.set("Завершение приложения...")
        self._finish_close_when_ready()

    def _finish_close_when_ready(self):
        """Дожидается завершения фоновых потоков и закрывает окно."""
        search_alive = self.search_thread is not None and self.search_thread.is_alive()
        add_alive = self.directory_add_thread is not None and self.directory_add_thread.is_alive()

        if search_alive or add_alive:
            self.safe_after(100, self._finish_close_when_ready, allow_when_closing=True)
            return

        try:
            if self.invalid_password_timer_id is not None:
                self.root.after_cancel(self.invalid_password_timer_id)
                self.invalid_password_timer_id = None
        except Exception:
            pass

        try:
            self.save_keywords_to_file()
            self.update_config(reload_after_save=False)
        finally:
            if self.search_file_handler is not None:
                try:
                    logging.getLogger().removeHandler(self.search_file_handler)
                    self.search_file_handler.close()
                except Exception:
                    pass
                self.search_file_handler = None
            self.root.destroy()

    def start_invalid_password_timer(self):
        """Запускает обратный отсчет до автозакрытия при неверном пароле."""
        self.invalid_password_deadline_ms = int(time.time() * 1000) + INVALID_PASSWORD_CLOSE_MS
        self._update_invalid_password_timer()

    def _update_invalid_password_timer(self):
        """Обновляет таймер обратного отсчета при неверном пароле."""
        if self.invalid_password_deadline_ms is None:
            return

        remaining_ms = self.invalid_password_deadline_ms - int(time.time() * 1000)
        if remaining_ms <= 0:
            self.invalid_password_timer_id = None
            self.on_close()
            return

        total_seconds = max(0, remaining_ms // 1000)
        minutes = total_seconds // 60
        seconds = total_seconds % 60
        timer_text = f"Ограниченный доступ. До закрытия: {minutes:02d}:{seconds:02d}"
        self.current_file.set(timer_text)
        self.root.title(f"Поиск файлов по ключевым словам [Ограниченный доступ {minutes:02d}:{seconds:02d}]")
        self.invalid_password_timer_id = self.root.after(INVALID_PASSWORD_TICK_MS, self._update_invalid_password_timer)

    def save_results(self):
        """Сохранение результатов в файл"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8-sig') as f:
                    f.write("Результаты поиска:\n\n")
                    for item_id in self.results_table.get_children():
                        keywords, file_path = self.results_table.item(item_id, "values")
                        f.write(f"Найденные слова: {keywords}\n")
                        f.write(f"Файл: {file_path}\n\n")
                messagebox.showinfo("Успех", "Результаты сохранены!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить результаты: {e}")


def main():
    """Основная функция"""
    root = tk.Tk()
    root.withdraw()

    entered_password = simpledialog.askstring(
        "Вход в приложение",
        "Введите пароль:",
        show="*",
        parent=root
    )
    password_is_valid = entered_password == APP_PASSWORD

    app = SearchApp(root)

    def handle_unhandled_exception(error_title, exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return

        log_path = write_crash_report(error_title, exc_value, exc_traceback)
        logging.error("%s: %s", error_title, exc_value)
        app.mark_search_crashed(error_title, exc_value)
        try:
            app.safe_after(
                0,
                app._show_runtime_error_dialog,
                "Критическая ошибка",
                (
                    f"{error_title}\n\n"
                    f"{exc_value}\n\n"
                    f"Подробности сохранены в файле:\n{log_path}\n\n"
                    "Если ошибка была нефатальной, приложение продолжит работу."
                )
            )
        except Exception:
            pass

    sys.excepthook = lambda exc_type, exc_value, exc_traceback: handle_unhandled_exception(
        "Необработанная ошибка приложения",
        exc_type,
        exc_value,
        exc_traceback
    )

    if hasattr(threading, "excepthook"):
        def thread_exception_handler(args):
            thread_name = args.thread.name if args.thread else "unknown-thread"
            handle_unhandled_exception(
                f"Необработанная ошибка в потоке {thread_name}",
                args.exc_type,
                args.exc_value,
                args.exc_traceback
            )

        threading.excepthook = thread_exception_handler

    root.report_callback_exception = (
        lambda exc_type, exc_value, exc_traceback: handle_unhandled_exception(
            "Необработанная ошибка интерфейса",
            exc_type,
            exc_value,
            exc_traceback
        )
    )

    root.deiconify()

    if not password_is_valid:
        root.title("Поиск файлов по ключевым словам [Ограниченный доступ]")
        messagebox.showwarning(
            "Неверный пароль",
            "Пароль неверный. Приложение будет закрыто через 5 минут.",
            parent=root
        )
        app.start_invalid_password_timer()

    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--pdf-worker":
        raise SystemExit(run_pdf_worker_cli(sys.argv[2:]))
    main()