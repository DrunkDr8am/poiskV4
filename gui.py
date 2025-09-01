import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import sys
import configparser
from PIL import Image, ImageTk
import logging

from main import check_dependencies

# Добавляем путь к текущей директории для импорта модулей
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Импорты из вашей программы
from config_loader import load_config, create_default_config
from tesseract_setup import setup_tesseract
from logging_setup import setup_logging
from file_processing import load_keywords
from search_engine import search_files

# Глобальные флаги для доступности функций (как в main.py)
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
        self.root.geometry("900x700")
        self.root.minsize(800, 600)

        # Переменные для хранения состояния
        self.selected_extensions = tk.StringVar(value="*.txt, *.pdf, *.docx, *.xlsx, *.jpg, *.png, *.zip, *.rar, *.7z")
        self.keywords_text = tk.StringVar()
        self.directories = []
        self.is_searching = False
        self.search_thread = None

        # Загружаем конфигурацию
        self.config = self.load_configuration()

        # Проверяем зависимости
        self.check_dependencies()

        # Создаем интерфейс
        self.create_widgets()

        # Центрируем окно
        self.center_window()

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def check_dependencies(self):
        """Проверка доступности опциональных зависимостей (аналогично main.py)"""
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

        # Выводим информацию о доступных функциях
        if not HAS_PDF:
            logging.warning("Обработка PDF недоступна")
        if not HAS_DOCX:
            logging.warning("Обработка DOCX недоступна")
        if not HAS_EXCEL:
            logging.warning("Обработка Excel недоступна")
        if not HAS_OCR and self.config.get('search_images', False):
            logging.warning("Поиск по изображениям недоступен (Tesseract не установлен)")
        if not HAS_7Z:
            logging.warning("Обработка 7z архивов недоступна")
        if not HAS_RAR:
            logging.warning("Обработка RAR архивов недоступна")

    def load_configuration(self):
        """Загрузка конфигурации с обработкой ошибок"""
        try:
            return load_config()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить конфигурацию: {e}")
            return {
                'extensions': ['*.txt', '*.pdf', '*.docx', '*.xlsx', '*.jpg', '*.png', '*.zip', '*.rar', '*.7z'],
                'keywords_file': 'keywords.txt',
                'directory': '.',
                'threads': 4,
                'output_file': 'search_results.txt',
                'search_images': False,
                'max_file_size': 50,
                'log_file': 'search_log.txt'
            }

    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # Row 0: Extensions selection
        ttk.Label(main_frame, text="Расширения файлов:").grid(row=0, column=0, sticky=tk.W, pady=5)
        extensions_frame = ttk.Frame(main_frame)
        extensions_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)

        extensions_entry = ttk.Entry(extensions_frame, textvariable=self.selected_extensions, width=50)
        extensions_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))
        extensions_frame.columnconfigure(0, weight=1)

        # Row 1: Keywords
        ttk.Label(main_frame, text="Ключевые слова:").grid(row=1, column=0, sticky=tk.NW, pady=5)
        keywords_text = scrolledtext.ScrolledText(main_frame, width=50, height=5)
        keywords_text.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        self.keywords_text_widget = keywords_text

        # Row 2: Directories
        ttk.Label(main_frame, text="Директории для поиска:").grid(row=2, column=0, sticky=tk.W, pady=5)

        dir_frame = ttk.Frame(main_frame)
        dir_frame.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)

        self.dirs_listbox = tk.Listbox(dir_frame, height=4)
        self.dirs_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        scrollbar = ttk.Scrollbar(dir_frame, orient=tk.VERTICAL, command=self.dirs_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.dirs_listbox.configure(yscrollcommand=scrollbar.set)

        dir_btn_frame = ttk.Frame(dir_frame)
        dir_btn_frame.grid(row=0, column=2, sticky=(tk.N, tk.S), padx=(5, 0))

        ttk.Button(dir_btn_frame, text="Добавить", command=self.add_directory).grid(row=0, column=0, pady=2)
        ttk.Button(dir_btn_frame, text="Удалить", command=self.remove_directory).grid(row=1, column=0, pady=2)

        dir_frame.columnconfigure(0, weight=1)
        dir_frame.rowconfigure(0, weight=1)

        # Row 3: Options
        options_frame = ttk.Frame(main_frame)
        options_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        ttk.Label(options_frame, text="Потоки:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.threads_var = tk.IntVar(value=self.config.get('threads', 4))
        threads_spin = ttk.Spinbox(options_frame, from_=1, to=16, textvariable=self.threads_var, width=5)
        threads_spin.grid(row=0, column=1, sticky=tk.W, padx=(0, 20))

        ttk.Label(options_frame, text="Макс. размер файла (МБ):").grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        self.max_size_var = tk.IntVar(value=self.config.get('max_file_size', 50))
        max_size_spin = ttk.Spinbox(options_frame, from_=1, to=1000, textvariable=self.max_size_var, width=5)
        max_size_spin.grid(row=0, column=3, sticky=tk.W, padx=(0, 20))

        self.search_images_var = tk.BooleanVar(value=self.config.get('search_images', False))
        ttk.Checkbutton(options_frame, text="Поиск по изображениям (OCR)", variable=self.search_images_var).grid(row=0,
                                                                                                                 column=4,
                                                                                                                 sticky=tk.W)

        # Row 4: Progress
        ttk.Label(main_frame, text="Прогресс:").grid(row=4, column=0, sticky=tk.W, pady=5)

        self.progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        progress_bar.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=5)

        # Row 5: Log
        ttk.Label(main_frame, text="Лог выполнения:").grid(row=5, column=0, sticky=tk.NW, pady=5)

        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=5, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, width=60, height=15)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        # Row 6: Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=10)

        self.search_button = ttk.Button(button_frame, text="Начать поиск", command=self.toggle_search)
        self.search_button.pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="Очистить лог", command=self.clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Сохранить лог", command=self.save_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Выход", command=self.root.quit).pack(side=tk.LEFT, padx=5)

        # Configure weights for main frame rows/columns
        main_frame.rowconfigure(5, weight=1)

        # Set up logging to text widget
        self.setup_logging()

    def setup_logging(self):
        """Настройка логирования в текстовый виджет"""

        class TextHandler(logging.Handler):
            def __init__(self, text_widget):
                super().__init__()
                self.text_widget = text_widget

            def emit(self, record):
                msg = self.format(record)
                self.text_widget.configure(state='normal')
                self.text_widget.insert(tk.END, msg + '\n')
                self.text_widget.see(tk.END)
                self.text_widget.configure(state='disabled')

        handler = TextHandler(self.log_text)
        handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logging.getLogger().addHandler(handler)
        logging.getLogger().setLevel(logging.INFO)

    def add_directory(self):
        directory = filedialog.askdirectory(title="Выберите директорию для поиска")
        if directory:
            self.directories.append(directory)
            self.update_dirs_listbox()

    def remove_directory(self):
        selected = self.dirs_listbox.curselection()
        if selected:
            index = selected[0]
            self.directories.pop(index)
            self.update_dirs_listbox()

    def update_dirs_listbox(self):
        self.dirs_listbox.delete(0, tk.END)
        for directory in self.directories:
            self.dirs_listbox.insert(tk.END, directory)

    def clear_log(self):
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')

    def save_log(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if filename:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(self.log_text.get(1.0, tk.END))

    def toggle_search(self):
        if self.is_searching:
            self.stop_search()
        else:
            self.start_search()

    def start_search(self):

        check_dependencies()

        # Проверяем введенные данные
        if not self.directories:
            messagebox.showerror("Ошибка", "Выберите хотя бы одну директорию для поиска")
            return

        keywords = self.keywords_text_widget.get(1.0, tk.END).strip()
        if not keywords:
            messagebox.showerror("Ошибка", "Введите ключевые слова для поиска")
            return

        extensions = [ext.strip() for ext in self.selected_extensions.get().split(',')]
        if not extensions:
            messagebox.showerror("Ошибка", "Выберите хотя бы одно расширение файла")
            return

        # Обновляем состояние UI
        self.is_searching = True
        self.search_button.config(text="Остановить поиск")

        # Запускаем поиск в отдельном потоке
        self.search_thread = threading.Thread(target=self.run_search, args=(keywords, extensions))
        self.search_thread.daemon = True
        self.search_thread.start()

    def stop_search(self):
        self.is_searching = False
        self.search_button.config(text="Начать поиск")
        logging.info("Поиск остановлен пользователем")

    def run_search(self, keywords, extensions):
        try:
            # Сохраняем ключевые слова во временный файл
            keywords_file = "temp_keywords.txt"
            with open(keywords_file, 'w', encoding='utf-8') as f:
                f.write(keywords)

            # Создаем конфигурацию для поиска
            config = {
                'extensions': extensions,
                'keywords_file': keywords_file,
                'directory': self.directories[0],  # используем первую директорию
                'threads': self.threads_var.get(),
                'output_file': 'search_results.txt',
                'search_images': self.search_images_var.get(),
                'max_file_size': self.max_size_var.get(),
                'log_file': 'search_log.txt'
            }

            # Если выбрано несколько директорий, обрабатываем их по очереди
            for directory in self.directories:
                if not self.is_searching:
                    break

                config['directory'] = directory
                logging.info(f"Начинаем поиск в директории: {directory}")

                # Загружаем ключевые слова
                try:
                    keywords_list = load_keywords(keywords_file)
                    if not keywords_list:
                        logging.error("Не найдено ключевых слов.")
                        continue
                except ValueError as e:
                    logging.error(e)
                    continue

                # Выполняем поиск
                results = search_files(
                    directory,
                    extensions,
                    config['threads'],
                    config['output_file'],
                    config['max_file_size'],
                    config
                )

                if results:
                    logging.info(f"Найдено совпадений в {len(results)} файлах")
                else:
                    logging.info("Ничего не найдено.")

            logging.info("Поиск завершен")

        except Exception as e:
            logging.error(f"Ошибка при выполнении поиска: {e}")
        finally:
            # Восстанавливаем состояние UI
            self.is_searching = False
            self.root.after(0, lambda: self.search_button.config(text="Начать поиск"))

            # Удаляем временный файл
            if os.path.exists(keywords_file):
                os.remove(keywords_file)


def main():
    root = tk.Tk()
    app = SearchApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()