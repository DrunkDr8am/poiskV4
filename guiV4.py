import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import logging
import time  # Добавляем импорт модуля time
import subprocess
import sys
from config_loader import load_config, create_default_config
from tesseract_setup import setup_tesseract
from file_processing import load_keywords
from search_engine import search_files
from configparser import ConfigParser

import fnmatch

# Глобальные флаги для доступности функций
HAS_PDF = False
HAS_DOCX = False
HAS_EXCEL = False
HAS_7Z = False
HAS_RAR = False
HAS_OCR = False


class SearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Поиск файлов по ключевым словам")
        self.root.geometry("1000x800")
        self.root.minsize(900, 700)

        # Переменные для хранения состояний
        self.extension_options = ['*.txt', '*.pdf', '*.docx', '*.xlsx', '*.jpg', '*.png', '*.zip', '*.rar', '*.7z']
        self.extension_vars = {ext: tk.BooleanVar(value=False) for ext in self.extension_options}
        self.directories_list = []
        self.is_searching = False
        self.is_paused = False
        self.search_thread = None
        self.is_adding_directory = False
        self.config_dirty = False
        self._updating_threads_var = False
        self.progress_value = tk.DoubleVar(value=0.0)
        self.current_file = tk.StringVar(value="")
        self.theme_var = tk.StringVar(value="Светлая")
        self.total_files = 0
        self.processed_files = 0
        self.search_start_time = None  # Время начала поиска
        self.search_end_time = None  # Время окончания поиска

        # Загружаем конфигурацию ДО создания интерфейса
        self.config = self.load_configuration()

        # Проверяем зависимости
        self.check_dependencies()

        # Создаем интерфейс
        self.create_widgets()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

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
        search_tab.rowconfigure(4, weight=1)

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

        ttk.Label(search_tab, text="Прогресс:").grid(row=1, column=0, sticky=tk.W, pady=5)
        progress_frame = ttk.Frame(search_tab)
        progress_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        progress_frame.columnconfigure(0, weight=1)
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_value, maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Label(progress_frame, textvariable=self.current_file).grid(row=1, column=0, sticky=(tk.W, tk.E))

        button_frame = ttk.Frame(search_tab)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        self.start_button = ttk.Button(button_frame, text="Начать поиск", command=self.start_search)
        self.start_button.pack(side=tk.LEFT, padx=5)
        self.pause_button = ttk.Button(button_frame, text="Пауза", command=self.toggle_pause, state=tk.DISABLED)
        self.pause_button.pack(side=tk.LEFT, padx=5)
        self.stop_button = ttk.Button(button_frame, text="Закончить поиск", command=self.stop_search, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Сохранить результаты", command=self.save_results).pack(side=tk.LEFT, padx=5)

        ttk.Label(search_tab, text="Результаты поиска:").grid(row=3, column=0, sticky=tk.NW, pady=5)
        results_frame = ttk.Frame(search_tab)
        results_frame.grid(row=3, column=1, rowspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        self.results_table = ttk.Treeview(results_frame, columns=("keywords", "file"), show="headings", height=14)
        self.results_table.heading("keywords", text="Найденные слова")
        self.results_table.heading("file", text="Ссылка на файл")
        self.results_table.column("keywords", width=320, minwidth=320, stretch=False, anchor=tk.W)
        self.results_table.column("file", width=620, minwidth=620, stretch=False, anchor=tk.W)
        results_scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_table.yview)
        self.results_table.configure(yscrollcommand=results_scrollbar.set)
        self.results_table.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        results_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.hovered_result_item = None
        self.results_table.bind("<Double-Button-1>", self.open_result_file)
        self.results_table.bind("<Motion>", self.on_results_hover)
        self.results_table.bind("<Leave>", self.on_results_leave)

        # ---------------- Настройки поиска ----------------
        settings_tab.columnconfigure(1, weight=1)
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

        ttk.Label(settings_tab, text="Макс. размер файла (МБ):").grid(row=3, column=0, sticky=tk.W, pady=3)
        self.max_size_var = tk.StringVar(value=str(self.config['config'].get('max_file_size', 50)))
        max_size_spin = ttk.Spinbox(settings_tab, from_=1, to=1000, textvariable=self.max_size_var, width=8)
        max_size_spin.grid(row=3, column=1, sticky=tk.W, pady=3)

        self.search_images_var = tk.BooleanVar(value=self.config['config'].get('search_images', False))
        ttk.Label(settings_tab, text="Поиск по изображениям (OCR):").grid(row=4, column=0, sticky=tk.W, pady=3)
        ttk.Checkbutton(settings_tab, text="Включить OCR", variable=self.search_images_var).grid(
            row=4, column=1, sticky=tk.W, pady=3
        )

        ttk.Label(settings_tab, text="Макс. страниц PDF (0=без лимита):").grid(row=5, column=0, sticky=tk.W, pady=3)
        self.max_pdf_pages_var = tk.StringVar(value=str(self.config['config'].get('max_pdf_pages', 0)))
        max_pdf_pages_spin = ttk.Spinbox(settings_tab, from_=0, to=100000, textvariable=self.max_pdf_pages_var, width=10)
        max_pdf_pages_spin.grid(row=5, column=1, sticky=tk.W, pady=3)

        ttk.Label(settings_tab, text="Тема интерфейса:").grid(row=6, column=0, sticky=tk.W, pady=(8, 3))
        self.theme_toggle_button = ttk.Button(settings_tab, text="", command=self.toggle_theme, width=14)
        self.theme_toggle_button.grid(row=6, column=1, sticky=tk.W, pady=(8, 3))

        self.setup_logging()
        self.apply_theme(self.theme_var.get())

    def toggle_theme(self):
        if self.theme_var.get() == "Светлая":
            self.theme_var.set("Темная")
        else:
            self.theme_var.set("Светлая")
        self.apply_theme(self.theme_var.get())

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
        worker = threading.Thread(target=self._add_directory_worker, args=(directory,), daemon=True)
        worker.start()

    def _add_directory_worker(self, directory):
        """Фоновая подготовка директории перед добавлением в UI."""
        try:
            normalized_directory = os.path.normpath(directory)
            self.root.after(0, lambda: self._finish_add_directory(normalized_directory))
        except Exception as e:
            self.root.after(0, lambda: self._finish_add_directory(None, error=e))

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

        self.progress_value.set(0)
        self.current_file.set("")
        self.processed_files = 0

    def update_progress(self, file_name=""):
        """Обновление прогресса с информацией о прогрессе"""
        logging.debug(f"Updating progress: {self.processed_files}/{self.total_files}, file: {file_name}")

        if self.total_files > 0:
            progress = (self.processed_files / self.total_files) * 100
            self.progress_value.set(progress)

            # Обновляем текст с информацией о прогрессе
            if file_name:
                # Обрезаем длинное имя файла для отображения
                display_name = file_name
                if len(file_name) > 50:
                    display_name = "..." + file_name[-47:]

                progress_text = f"Обработано: {self.processed_files}/{self.total_files} файлов"

                # Различаем разные типы сообщений
                if file_name.startswith("Завершена обработка:"):
                    progress_text += f" | {file_name}"
                elif file_name.startswith("Начат:"):
                    progress_text += f" | {file_name.replace('Начат:', 'Текущий:', 1)}"
                elif file_name == "Поиск завершен":
                    progress_text = "Поиск завершен! Обработано всех файлов."
                else:
                    progress_text += f" | {display_name}"

                self.current_file.set(progress_text)

        # Принудительно обновляем прогрессбар
        self.progress_bar.update_idletasks()

    def add_live_result(self, file_path, keywords):
        """Добавление найденного результата в таблицу в реальном времени."""
        keywords_str = ', '.join(sorted(keywords)) if keywords else ''
        self.root.after(0, lambda: self.results_table.insert('', tk.END, values=(keywords_str, file_path)))

    def open_result_file(self, event):
        """Открытие файла из выбранной строки таблицы по двойному клику."""
        try:
            item_id = self.results_table.identify_row(event.y)
            if not item_id:
                return

            values = self.results_table.item(item_id, "values")
            if not values or len(values) < 2:
                return

            file_path = str(values[1]).strip()
            if not file_path or not os.path.exists(file_path):
                messagebox.showwarning("Файл не найден", f"Файл не существует:\n{file_path}")
                return

            if hasattr(os, "startfile"):
                os.startfile(file_path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", file_path])
            else:
                subprocess.Popen(["xdg-open", file_path])
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{e}")

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

    def count_files_to_process(self, directory, extensions):
        """Подсчет общего количества файлов для обработки"""
        count = 0
        for root, _, files in os.walk(directory):
            for file in files:
                file_path = os.path.join(root, file)
                # Используем fnmatch для проверки соответствия расширениям
                if any(fnmatch.fnmatch(file, ext) for ext in extensions):
                    count += 1
        return count

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

        # Сбрасываем счетчики
        self.processed_files = 0
        self.total_files = 0

        # Подсчитываем общее количество файлов для прогресса
        for directory in self.directories_list:
            self.total_files += self.count_files_to_process(directory, extensions)

        if self.total_files == 0:
            messagebox.showwarning("Предупреждение", "Не найдено файлов для обработки в указанных директориях!")
            return

        # Очищаем результаты и лог
        self.clear_all()

        # Записываем время начала поиска
        self.search_start_time = time.strftime('%Y-%m-%d %H:%M:%S')
        start_message = f"Поиск начат: {self.search_start_time}"
        logging.info(start_message)

        # Обновляем конфиг
        self.update_config()

        # Добавляем информацию о доступности модулей в конфиг
        self.config['config']['has_pdf'] = HAS_PDF
        self.config['config']['has_docx'] = HAS_DOCX
        self.config['config']['has_excel'] = HAS_EXCEL
        self.config['config']['has_7z'] = HAS_7Z
        self.config['config']['has_rar'] = HAS_RAR
        self.config['config']['has_ocr'] = HAS_OCR

        # Настраиваем логирование без удаления файла
        try:
            # Просто добавляем обработчик, не удаляем старый файл
            file_handler = logging.FileHandler('search_log.txt', mode='a', encoding='utf-8')
            file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
            logging.getLogger().addHandler(file_handler)
        except Exception as e:
            logging.error(f"Не удалось настроить файловое логирование: {e}")

        # Загружаем ключевые слова
        try:
            load_keywords("keywords.txt")
        except ValueError as e:
            messagebox.showerror("Ошибка", str(e))
            return

        # Меняем состояние кнопок
        self.start_button.config(state=tk.DISABLED)
        self.pause_button.config(state=tk.NORMAL, text="Пауза")
        self.stop_button.config(state=tk.NORMAL)
        self.is_searching = True
        self.is_paused = False
        self.processed_files = 0

        # Запускаем поиск в отдельном потоке
        self.search_thread = threading.Thread(
            target=self.run_search,
            args=(extensions, self.update_progress_callback)  # Передаем callback
        )
        self.search_thread.daemon = True
        self.search_thread.start()

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
        try:
            # Сбрасываем только processed_files при начале нового поиска
            self.processed_files = 0
            threads_count = self.normalize_threads_value()
            logging.info(f"Начинаем поиск. Всего файлов: {self.total_files}")

            # Выполняем поиск для каждой директории с накоплением счетчика
            for directory in self.directories_list:
                if not self.is_searching:
                    logging.info("Поиск остановлен пользователем")
                    break

                logging.info(f"Начинаем поиск в директории: {directory}")

                # Обновляем статус - начало обработки директории
                self.root.after(0, lambda: self.update_progress(f"Начата обработка: {os.path.basename(directory)}"))

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
                    lambda: self.is_paused
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
                # Используем другое сообщение, не "Поиск завершен"
                if self.is_searching:
                    self.root.after(0,
                                    lambda: self.update_progress(f"Завершена обработка: {os.path.basename(directory)}"))

            # Только после ВСЕХ директорий показываем завершение поиска
            if self.is_searching:
                logging.info("Поиск завершен!")
                self.root.after(0, lambda: self.update_progress("Поиск завершен"))

        except Exception as e:
            logging.error(f"Ошибка при поиске: {e}")

        finally:
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

            self.is_searching = False
            self.is_paused = False
            self.root.after(0, self.on_search_finished)

    def update_progress_callback(self, file_name, processed_count):
        """Callback для обновления прогресса из search_engine"""
        # Обновляем в основном потоке через after
        self.root.after(0, lambda: self._update_progress_in_main_thread(file_name, processed_count))

    def _update_progress_in_main_thread(self, file_name, processed_count):
        """Обновление прогресса в основном потоке"""
        if isinstance(processed_count, int):
            self.processed_files = processed_count
        self.update_progress(file_name)

    def on_search_finished(self):
        """Вызывается при завершении поиска"""
        self.start_button.config(state=tk.NORMAL)
        self.pause_button.config(state=tk.DISABLED, text="Пауза")
        self.stop_button.config(state=tk.DISABLED)
        # Если остановка была пользователем, финализируем человекочитаемый статус
        if self.current_file.get() == "Останавливаем поиск...":
            self.current_file.set("Поиск остановлен пользователем")

    def update_config(self, reload_after_save=True):
        """Обновление конфигурации"""
        config = ConfigParser()
        extensions = self.get_selected_extensions()

        current_cfg = self.config['config']
        threads_count = self.normalize_threads_value()

        # Обновляем конфиг
        config['Settings'] = {
            'extensions': ', '.join(extensions),
            'keywords_file': current_cfg.get('keywords_file', 'keywords.txt'),
            'directories': '; '.join(self.directories_list) if self.directories_list else '.',
            'directory': self.directories_list[0] if self.directories_list else current_cfg.get('directory', '.'),
            'threads': str(threads_count),
            'output_file': current_cfg.get('output_file', 'search_results.txt'),
            'search_images': 'true' if self.search_images_var.get() else 'false',
            'max_file_size': self.max_size_var.get(),
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

    def on_close(self):
        """Сохранение конфига при закрытии окна."""
        try:
            if self.config_dirty:
                self.update_config(reload_after_save=False)
        finally:
            self.root.destroy()

    def save_results(self):
        """Сохранение результатов в файл"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
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
    app = SearchApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()