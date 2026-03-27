import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
import threading
import os
import logging
import time  # Добавляем импорт модуля time
import subprocess
import sys
import traceback
from config_loader import (
    load_config,
    create_default_config,
    save_admin_password_hash,
    make_role_password_hash,
    verify_password_for_role,
    encode_user_permissions,
)
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
CRASH_LOG_FILE = "crash_log.txt"
ADMIN_ROLE = "admin"
USER_ROLE = "user"


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


def resolve_access_role(password, config):
    """Определяет роль пользователя по введенному паролю."""
    admin_hash = str(config.get("admin_password_hash", "")).strip()
    user_hash = str(config.get("user_password_hash", "")).strip()

    if verify_password_for_role(password, admin_hash, ADMIN_ROLE, conflicting_hash=user_hash):
        return ADMIN_ROLE
    if verify_password_for_role(password, user_hash, USER_ROLE, conflicting_hash=admin_hash):
        return USER_ROLE
    return None


def prompt_for_access_role(root):
    """Запрашивает пароль и возвращает роль пользователя."""
    if not os.path.exists("config.txt"):
        create_default_config()

    auth_config = load_config()
    while True:
        entered_password = simpledialog.askstring(
            "Вход в приложение",
            "Введите пароль:",
            show="*",
            parent=root
        )
        if entered_password is None:
            return None

        access_role = resolve_access_role(entered_password, auth_config)
        if access_role:
            return access_role

        messagebox.showerror("Ошибка входа", "Неверный пароль.", parent=root)


class SearchApp:
    def __init__(self, root, access_role):
        self.root = root
        self.root.title("Поиск файлов по ключевым словам")
        self.root.geometry("1000x800")
        self.root.minsize(900, 700)
        self.access_role = access_role
        self.is_admin = access_role == ADMIN_ROLE

        # Переменные для хранения состояний
        self.extension_options = ['*.txt', '*.pdf', '*.docx', '*.xlsx', '*.jpg', '*.png', '*.zip', '*.rar', '*.7z']
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
        self.progress_value = tk.DoubleVar(value=0.0)
        self.current_file = tk.StringVar(value="")
        self.theme_var = tk.StringVar(value="Светлая")
        self.total_files = 0
        self.processed_files = 0
        self.search_start_time = None  # Время начала поиска
        self.search_end_time = None  # Время окончания поиска

        # Загружаем конфигурацию ДО создания интерфейса
        self.config = self.load_configuration()
        self.user_permission_vars = {
            'extensions': tk.BooleanVar(value=self.config['config'].get('allow_user_change_extensions', True)),
            'keywords': tk.BooleanVar(value=self.config['config'].get('allow_user_change_keywords', True)),
            'threads': tk.BooleanVar(value=self.config['config'].get('allow_user_change_threads', True)),
            'max_file_size': tk.BooleanVar(value=self.config['config'].get('allow_user_change_max_file_size', True)),
            'search_images': tk.BooleanVar(value=self.config['config'].get('allow_user_change_search_images', True)),
            'max_pdf_pages': tk.BooleanVar(value=self.config['config'].get('allow_user_change_max_pdf_pages', True)),
            'theme': tk.BooleanVar(value=self.config['config'].get('allow_user_change_theme', True)),
            'results_context_menu': tk.BooleanVar(value=self.config['config'].get('allow_user_results_context_menu', True)),
        }
        self.extension_checkbuttons = []

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

    def mark_permissions_dirty(self):
        """Помечает изменение пользовательских прав как несохраненное."""
        self.config_dirty = True

    def can_user_change(self, permission_key):
        """Возвращает, разрешено ли текущему пользователю менять конкретную настройку."""
        if self.is_admin:
            return True
        return bool(self.config['config'].get(permission_key, True))

    def add_permission_checkbox(self, parent, row, permission_var, pady=3):
        """Добавляет админскую галочку, определяющую доступность настройки для пользователя."""
        if not self.is_admin:
            return
        ttk.Checkbutton(
            parent,
            text="Пользователь может менять",
            variable=permission_var,
            command=self.mark_permissions_dirty
        ).grid(row=row, column=2, sticky=tk.W, padx=(12, 0), pady=pady)

    def apply_user_permissions(self):
        """Применяет ограничения интерфейса для пользователя по настройкам администратора."""
        if self.can_user_change('allow_user_change_extensions'):
            for checkbox in self.extension_checkbuttons:
                checkbox.state(["!disabled"])
        else:
            for checkbox in self.extension_checkbuttons:
                checkbox.state(["disabled"])

        self.keywords_text.configure(
            state=tk.NORMAL if self.can_user_change('allow_user_change_keywords') else tk.DISABLED
        )
        self.threads_spin.configure(
            state=tk.NORMAL if self.can_user_change('allow_user_change_threads') else tk.DISABLED
        )
        self.max_size_spin.configure(
            state=tk.NORMAL if self.can_user_change('allow_user_change_max_file_size') else tk.DISABLED
        )
        self.search_images_checkbox.state(
            ["!disabled"] if self.can_user_change('allow_user_change_search_images') else ["disabled"]
        )
        self.max_pdf_pages_spin.configure(
            state=tk.NORMAL if self.can_user_change('allow_user_change_max_pdf_pages') else tk.DISABLED
        )
        self.theme_toggle_button.state(
            ["!disabled"] if self.can_user_change('allow_user_change_theme') else ["disabled"]
        )

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
        self.results_table.column("keywords", width=150, minwidth=100, stretch=True, anchor=tk.W)
        self.results_table.column("file", width=850, minwidth=300, stretch=True, anchor=tk.W)
        results_scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_table.yview)
        self.results_table.configure(yscrollcommand=results_scrollbar.set)
        self.results_table.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        results_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
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
        settings_tab.columnconfigure(2, weight=0)
        settings_tab.grid_anchor("nw")

        # Поля настроек идут по порядку
        ttk.Label(settings_tab, text="Расширения файлов:").grid(row=0, column=0, sticky=tk.NW, pady=3)
        extensions_frame = ttk.Frame(settings_tab)
        extensions_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=3)
        columns = 5
        for index, ext in enumerate(self.extension_options):
            row = index // columns
            col = index % columns
            checkbox = ttk.Checkbutton(extensions_frame, text=ext, variable=self.extension_vars[ext])
            checkbox.grid(row=row, column=col, sticky=tk.W, padx=(0, 12), pady=2)
            self.extension_checkbuttons.append(checkbox)
        self.add_permission_checkbox(settings_tab, 0, self.user_permission_vars['extensions'])

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
        self.add_permission_checkbox(settings_tab, 1, self.user_permission_vars['keywords'])

        ttk.Label(settings_tab, text="Потоки:").grid(row=2, column=0, sticky=tk.W, pady=3)
        configured_threads = self.config['config'].get('threads', self.get_auto_threads_count())
        self.threads_var = tk.StringVar(value=str(configured_threads))
        self.threads_var.trace_add("write", self.on_threads_var_change)
        self.threads_spin = ttk.Spinbox(
            settings_tab, from_=1, to=self.get_max_threads_count(), textvariable=self.threads_var, width=8
        )
        self.threads_spin.grid(row=2, column=1, sticky=tk.W, pady=3)
        self.add_permission_checkbox(settings_tab, 2, self.user_permission_vars['threads'])

        ttk.Label(settings_tab, text="Макс. размер файла (МБ):").grid(row=3, column=0, sticky=tk.W, pady=3)
        self.max_size_var = tk.StringVar(value=str(self.config['config'].get('max_file_size', 50)))
        self.max_size_spin = ttk.Spinbox(settings_tab, from_=1, to=1000, textvariable=self.max_size_var, width=8)
        self.max_size_spin.grid(row=3, column=1, sticky=tk.W, pady=3)
        self.add_permission_checkbox(settings_tab, 3, self.user_permission_vars['max_file_size'])

        self.search_images_var = tk.BooleanVar(value=self.config['config'].get('search_images', False))
        ttk.Label(settings_tab, text="Поиск по изображениям (OCR):").grid(row=4, column=0, sticky=tk.W, pady=3)
        self.search_images_checkbox = ttk.Checkbutton(settings_tab, text="Включить OCR", variable=self.search_images_var)
        self.search_images_checkbox.grid(
            row=4, column=1, sticky=tk.W, pady=3
        )
        self.add_permission_checkbox(settings_tab, 4, self.user_permission_vars['search_images'])

        ttk.Label(settings_tab, text="Макс. страниц PDF (0=без лимита):").grid(row=5, column=0, sticky=tk.W, pady=3)
        self.max_pdf_pages_var = tk.StringVar(value=str(self.config['config'].get('max_pdf_pages', 0)))
        self.max_pdf_pages_spin = ttk.Spinbox(
            settings_tab, from_=0, to=100000, textvariable=self.max_pdf_pages_var, width=10
        )
        self.max_pdf_pages_spin.grid(row=5, column=1, sticky=tk.W, pady=3)
        self.add_permission_checkbox(settings_tab, 5, self.user_permission_vars['max_pdf_pages'])

        ttk.Label(settings_tab, text="Тема интерфейса:").grid(row=6, column=0, sticky=tk.W, pady=(8, 3))
        self.theme_toggle_button = ttk.Button(settings_tab, text="", command=self.toggle_theme, width=14)
        self.theme_toggle_button.grid(row=6, column=1, sticky=tk.W, pady=(8, 3))
        self.add_permission_checkbox(settings_tab, 6, self.user_permission_vars['theme'], pady=(8, 3))

        if self.is_admin:
            ttk.Label(settings_tab, text="ПКМ по таблице результатов:").grid(row=7, column=0, sticky=tk.W, pady=(8, 3))
            ttk.Label(settings_tab, text="Открывать контекстное меню").grid(row=7, column=1, sticky=tk.W, pady=(8, 3))
            self.add_permission_checkbox(settings_tab, 7, self.user_permission_vars['results_context_menu'], pady=(8, 3))

        if self.is_admin:
            ttk.Button(
                settings_tab,
                text="Смена пароля для администратора",
                command=self.change_admin_password
            ).grid(row=8, column=0, columnspan=2, sticky=tk.W, pady=(12, 3))
            ttk.Button(
                settings_tab,
                text="Смена пароля для пользователя",
                command=self.change_user_password
            ).grid(row=9, column=0, columnspan=2, sticky=tk.W, pady=(3, 3))

        self.setup_logging()
        self.apply_theme(self.theme_var.get())
        self.apply_user_permissions()

    def toggle_theme(self):
        if self.theme_var.get() == "Светлая":
            self.theme_var.set("Темная")
        else:
            self.theme_var.set("Светлая")
        self.apply_theme(self.theme_var.get())
        self.config_dirty = True
        self.update_config(reload_after_save=False)

    def change_user_password(self):
        """Позволяет администратору изменить пароль пользователя."""
        if not self.is_admin:
            return

        new_password = simpledialog.askstring(
            "Смена пароля пользователя",
            "Введите новый пароль для пользователя:",
            show="*",
            parent=self.root
        )
        if new_password is None:
            return

        new_password = new_password.strip()
        if not new_password:
            messagebox.showwarning("Предупреждение", "Пароль пользователя не может быть пустым.", parent=self.root)
            return

        confirm_password = simpledialog.askstring(
            "Смена пароля пользователя",
            "Повторите новый пароль:",
            show="*",
            parent=self.root
        )
        if confirm_password is None:
            return

        if new_password != confirm_password:
            messagebox.showerror("Ошибка", "Пароли не совпадают.", parent=self.root)
            return

        new_hash = make_role_password_hash(new_password, USER_ROLE)
        admin_hash = self.config['config'].get('admin_password_hash', '')
        if verify_password_for_role(new_password, admin_hash, ADMIN_ROLE):
            messagebox.showerror(
                "Ошибка",
                "Пароль пользователя должен отличаться от пароля администратора.",
                parent=self.root
            )
            return

        try:
            self.config['config']['user_password_hash'] = new_hash
            self.config_dirty = True
            self.update_config(reload_after_save=True)
            messagebox.showinfo("Успех", "Пароль пользователя успешно изменен.", parent=self.root)
        except Exception as e:
            self.report_runtime_error("Не удалось изменить пароль пользователя", e, show_dialog=True)

    def change_admin_password(self):
        """Позволяет администратору изменить пароль администратора."""
        if not self.is_admin:
            return

        current_password = simpledialog.askstring(
            "Смена пароля администратора",
            "Введите текущий пароль администратора:",
            show="*",
            parent=self.root
        )
        if current_password is None:
            return

        admin_hash = self.config['config'].get('admin_password_hash', '')
        if admin_hash and not verify_password_for_role(
            current_password,
            admin_hash,
            ADMIN_ROLE,
            conflicting_hash=self.config['config'].get('user_password_hash', '')
        ):
            messagebox.showerror("Ошибка", "Текущий пароль администратора введен неверно.", parent=self.root)
            return

        new_password = simpledialog.askstring(
            "Смена пароля администратора",
            "Введите новый пароль администратора:",
            show="*",
            parent=self.root
        )
        if new_password is None:
            return

        new_password = new_password.strip()
        if not new_password:
            messagebox.showwarning("Предупреждение", "Пароль администратора не может быть пустым.", parent=self.root)
            return

        confirm_password = simpledialog.askstring(
            "Смена пароля администратора",
            "Повторите новый пароль администратора:",
            show="*",
            parent=self.root
        )
        if confirm_password is None:
            return

        if new_password != confirm_password:
            messagebox.showerror("Ошибка", "Пароли не совпадают.", parent=self.root)
            return

        new_hash = make_role_password_hash(new_password, ADMIN_ROLE)
        user_hash = self.config['config'].get('user_password_hash', '')
        if verify_password_for_role(new_password, user_hash, USER_ROLE):
            messagebox.showerror(
                "Ошибка",
                "Пароль администратора должен отличаться от пароля пользователя.",
                parent=self.root
            )
            return

        try:
            self.config['config']['admin_password_hash'] = new_hash
            save_admin_password_hash(new_hash)
            messagebox.showinfo("Успех", "Пароль администратора успешно изменен.", parent=self.root)
        except Exception as e:
            self.report_runtime_error("Не удалось изменить пароль администратора", e, show_dialog=True)

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
        self.safe_after(0, self.results_table.insert, '', tk.END, values=(keywords_str, file_path))

    def on_results_table_resize(self, event):
        """Поддерживает пропорцию колонок 15% / 85%."""
        total_width = max(1, event.width - 4)
        keywords_width = max(100, int(total_width * 0.15))
        file_width = max(300, total_width - keywords_width)
        self.results_table.column("keywords", width=keywords_width)
        self.results_table.column("file", width=file_width)

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

    def _open_file_by_path(self, file_path):
        """Открывает файл в системе."""
        if not file_path or not os.path.exists(file_path):
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
            if not file_path or not os.path.exists(file_path):
                messagebox.showwarning("Файл не найден", f"Файл не существует:\n{file_path}")
                return

            if sys.platform.startswith("win"):
                subprocess.Popen(["explorer", "/select,", os.path.normpath(file_path)])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", file_path])
            else:
                subprocess.Popen(["xdg-open", os.path.dirname(file_path) or "."])
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть расположение файла:\n{e}")

    def show_results_context_menu(self, event):
        """Показывает контекстное меню для строки таблицы результатов."""
        if not self.can_user_change('allow_user_results_context_menu'):
            return

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

        # Нормализуем текущее значение потоков, заданное в настройках
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
            args=(extensions, self.update_progress_callback),  # Передаем callback
            name="search-worker"
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

        except Exception as e:
            self.report_runtime_error("Критическая ошибка во время поиска", e, show_dialog=True)

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
            self.safe_after(0, self.on_search_finished)

    def update_progress_callback(self, file_name, processed_count):
        """Callback для обновления прогресса из search_engine"""
        # Обновляем в основном потоке через after
        self.safe_after(0, self._update_progress_in_main_thread, file_name, processed_count)

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
        permission_values = {
            'allow_user_change_extensions': self.user_permission_vars['extensions'].get(),
            'allow_user_change_keywords': self.user_permission_vars['keywords'].get(),
            'allow_user_change_threads': self.user_permission_vars['threads'].get(),
            'allow_user_change_max_file_size': self.user_permission_vars['max_file_size'].get(),
            'allow_user_change_search_images': self.user_permission_vars['search_images'].get(),
            'allow_user_change_max_pdf_pages': self.user_permission_vars['max_pdf_pages'].get(),
            'allow_user_change_theme': self.user_permission_vars['theme'].get(),
            'allow_user_results_context_menu': self.user_permission_vars['results_context_menu'].get(),
        }
        permissions_token = encode_user_permissions(permission_values, current_cfg.get('admin_password_hash', ''))

        # Обновляем конфиг
        config['Settings'] = {
            'extensions': ', '.join(extensions),
            'keywords_file': current_cfg.get('keywords_file', 'keywords.txt'),
            'directories': '; '.join(self.directories_list) if self.directories_list else '.',
            'directory': self.directories_list[0] if self.directories_list else current_cfg.get('directory', '.'),
            'theme': self.theme_var.get(),
            'threads': str(threads_count),
            'output_file': current_cfg.get('output_file', 'search_results.txt'),
            'search_images': 'true' if self.search_images_var.get() else 'false',
            'max_file_size': self.max_size_var.get(),
            'log_file': current_cfg.get('log_file', 'search_log.txt'),
            'tesseract_languages': current_cfg.get('tesseract_languages', 'rus'),
            'tesseract_config': current_cfg.get('tesseract_config', '--oem 3 --psm 6'),
            # Новые параметры для тонкой настройки поиска
            'max_pdf_pages': self.max_pdf_pages_var.get(),
            'user_password_hash': current_cfg.get('user_password_hash', ''),
            'user_permissions_token': permissions_token,
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
    access_role = prompt_for_access_role(root)
    if access_role is None:
        root.destroy()
        return

    app = SearchApp(root, access_role)

    def handle_unhandled_exception(error_title, exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return

        log_path = write_crash_report(error_title, exc_value, exc_traceback)
        logging.error("%s: %s", error_title, exc_value)
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
    root.mainloop()


if __name__ == "__main__":
    main()