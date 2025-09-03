import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import sys
import logging
from config_loader import load_config, create_default_config
from tesseract_setup import setup_tesseract
from logging_setup import setup_logging
from file_processing import load_keywords
from search_engine import search_files
from configparser import ConfigParser
from PIL import Image, ImageTk

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
        self.root.geometry("900x700")
        self.root.minsize(800, 600)

        # Переменные для хранения состояний
        self.selected_extensions = tk.StringVar()
        self.directories_list = []
        self.is_searching = False
        self.search_thread = None

        # Загружаем конфигурацию ДО создания интерфейса
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

    def load_configuration(self):
        """Загрузка конфигурации"""
        if not os.path.exists("config.txt"):
            create_default_config()

        config = load_config()

        # Устанавливаем значение для текстового поля расширений
        extensions_str = ', '.join(config['extensions'])
        if not hasattr(self, 'selected_extensions'):
            self.selected_extensions = tk.StringVar(value=extensions_str)
        else:
            self.selected_extensions.set(extensions_str)

        return {
            'config': config,
            'extensions_str': extensions_str
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
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Configure grid weights
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=2)

        # Row 0: Extensions selection
        ttk.Label(main_frame, text="Расширения файлов:").grid(row=0, column=0, sticky=tk.W, pady=5)
        extensions_frame = ttk.Frame(main_frame)
        extensions_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)

        self.extensions_entry = ttk.Entry(extensions_frame, textvariable=self.selected_extensions, width=50)
        self.extensions_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))
        extensions_frame.columnconfigure(0, weight=1)

        # Row 1: Keywords
        ttk.Label(main_frame, text="Ключевые слова:").grid(row=1, column=0, sticky=tk.NW, pady=5)

        keywords_frame = ttk.Frame(main_frame)
        keywords_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        self.keywords_text = scrolledtext.ScrolledText(keywords_frame, height=5)
        self.keywords_text.pack(fill=tk.BOTH, expand=True)

        # Загружаем ключевые слова из файла, если он существует
        if os.path.exists("keywords.txt"):
            try:
                with open("keywords.txt", "r", encoding="utf-8") as f:
                    keywords = f.read()
                    self.keywords_text.insert("1.0", keywords)
            except:
                pass

        # Row 2: Directories
        ttk.Label(main_frame, text="Директории для поиска:").grid(row=2, column=0, sticky=tk.NW, pady=5)

        dir_frame = ttk.Frame(main_frame)
        dir_frame.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)

        self.dirs_listbox = tk.Listbox(dir_frame, height=4)
        self.dirs_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(dir_frame, orient=tk.VERTICAL, command=self.dirs_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.dirs_listbox.configure(yscrollcommand=scrollbar.set)

        dir_btn_frame = ttk.Frame(dir_frame)
        dir_btn_frame.pack(side=tk.RIGHT, padx=(5, 0))

        ttk.Button(dir_btn_frame, text="Добавить", command=self.add_directory).pack(pady=2)
        ttk.Button(dir_btn_frame, text="Удалить", command=self.remove_directory).pack(pady=2)

        # Добавляем текущую директорию по умолчанию
        self.directories_list.append(".")
        self.dirs_listbox.insert(tk.END, ".")

        # Row 3: Options
        options_frame = ttk.Frame(main_frame)
        options_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        ttk.Label(options_frame, text="Потоки:").pack(side=tk.LEFT, padx=(0, 5))
        self.threads_var = tk.StringVar(value=str(self.config['config'].get('threads', 4)))
        threads_spin = ttk.Spinbox(options_frame, from_=1, to=16, textvariable=self.threads_var, width=5)
        threads_spin.pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(options_frame, text="Макс. размер файла (МБ):").pack(side=tk.LEFT, padx=(0, 5))
        self.max_size_var = tk.StringVar(value=str(self.config['config'].get('max_file_size', 50)))
        max_size_spin = ttk.Spinbox(options_frame, from_=1, to=1000, textvariable=self.max_size_var, width=5)
        max_size_spin.pack(side=tk.LEFT, padx=(0, 20))

        self.search_images_var = tk.BooleanVar(value=self.config['config'].get('search_images', False))
        ttk.Checkbutton(options_frame, text="Поиск по изображениям (OCR)",
                        variable=self.search_images_var).pack(side=tk.LEFT)

        # Row 4: Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)

        self.start_button = ttk.Button(button_frame, text="Начать поиск", command=self.start_search)
        self.start_button.pack(side=tk.LEFT, padx=5)

        self.stop_button = ttk.Button(button_frame, text="Остановить", command=self.stop_search, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="Сохранить результаты", command=self.save_results).pack(side=tk.LEFT, padx=5)

        # Row 5: Log
        ttk.Label(main_frame, text="Лог выполнения:").grid(row=5, column=0, sticky=tk.NW, pady=5)

        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=5, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=15)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Настраиваем логирование в текстовое поле
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
        """Добавление директории для поиска"""
        directory = filedialog.askdirectory(title="Выберите директорию для поиска")
        if directory:
            self.directories_list.append(directory)
            self.dirs_listbox.insert(tk.END, directory)

    def remove_directory(self):
        """Удаление выбранной директории"""
        selection = self.dirs_listbox.curselection()
        if selection:
            index = selection[0]
            self.dirs_listbox.delete(index)
            del self.directories_list[index]

    def start_search(self):
        """Запуск поиска в отдельном потоке"""
        if self.is_searching:
            return

        # Сохраняем текущие значения перед обновлением конфига
        current_extensions = self.selected_extensions.get()
        current_threads = self.threads_var.get()
        current_max_size = self.max_size_var.get()
        current_search_images = self.search_images_var.get()

        # Получаем выбранные расширения из текстового поля
        extensions_str = current_extensions
        extensions = [ext.strip() for ext in extensions_str.split(',') if ext.strip()]

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

        # Обновляем конфиг
        self.update_config()

        # Добавляем информацию о доступности модулей в конфиг
        self.config['config']['has_pdf'] = HAS_PDF
        self.config['config']['has_docx'] = HAS_DOCX
        self.config['config']['has_excel'] = HAS_EXCEL
        self.config['config']['has_7z'] = HAS_7Z
        self.config['config']['has_rar'] = HAS_RAR
        self.config['config']['has_ocr'] = HAS_OCR

        # Настраиваем логирование
        setup_logging(self.config['config'].get('log_file', 'search_log.txt'))

        # Загружаем ключевые слова
        try:
            load_keywords("keywords.txt")
        except ValueError as e:
            messagebox.showerror("Ошибка", str(e))
            return

        # Очищаем лог
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')

        # Меняем состояние кнопок
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.is_searching = True

        # Запускаем поиск в отдельном потоке
        self.search_thread = threading.Thread(target=self.run_search)
        self.search_thread.daemon = True
        self.search_thread.start()

    def stop_search(self):
        """Остановка поиска"""
        if self.is_searching:
            self.is_searching = False
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            logging.info("Поиск остановлен пользователем")

    def run_search(self):
        """Выполнение поиска"""
        try:
            # Получаем выбранные расширения из текстового поля
            extensions_str = self.selected_extensions.get()
            extensions = [ext.strip() for ext in extensions_str.split(',') if ext.strip()]

            # Выполняем поиск для каждой директории
            for directory in self.directories_list:
                if not self.is_searching:
                    break

                logging.info(f"Поиск в директории: {directory}")

                results = search_files(
                    directory,
                    extensions,
                    int(self.threads_var.get()),
                    "search_results.txt",
                    int(self.max_size_var.get()),
                    self.config['config']
                )

                if results:
                    logging.info(f"Найдено совпадений в {len(results)} файлах:")
                    for file_path, keywords in results.items():
                        logging.info(f"Файл: {file_path}")
                        logging.info(f"Ключевые слова: {', '.join(keywords)}")
                else:
                    logging.info("Ничего не найдено.")

            if self.is_searching:
                logging.info("Поиск завершен!")

        except Exception as e:
            logging.error(f"Ошибка при поиске: {e}")

        finally:
            self.is_searching = False
            self.root.after(0, self.on_search_finished)

    def on_search_finished(self):
        """Вызывается при завершении поиска"""
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

    def update_config(self):
        """Обновление конфигурации"""
        config = ConfigParser()

        # Получаем выбранные расширения из текстового поля
        extensions_str = self.selected_extensions.get()
        extensions = [ext.strip() for ext in extensions_str.split(',') if ext.strip()]

        # Обновляем конфиг
        config['Settings'] = {
            'extensions': ', '.join(extensions),
            'keywords_file': 'keywords.txt',
            'directory': self.directories_list[0] if self.directories_list else '.',
            'threads': self.threads_var.get(),
            'output_file': 'search_results.txt',
            'search_images': 'true' if self.search_images_var.get() else 'false',
            'max_file_size': self.max_size_var.get(),
            'log_file': 'search_log.txt',
            'tesseract_languages': self.config['config'].get('tesseract_languages', 'rus'),
            'tesseract_config': self.config['config'].get('tesseract_config', '--oem 3 --psm 6')
        }

        # Сохраняем конфиг
        with open('config.txt', 'w', encoding='utf-8') as configfile:
            config.write(configfile)

        # Обновляем self.config без сброса значений в интерфейсе
        new_config = load_config()
        self.config['config'] = new_config

    def save_results(self):
        """Сохранение результатов в файл"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")]
        )
        if filename:
            try:
                # Читаем результаты из файла поиска
                if os.path.exists("search_results.txt"):
                    with open("search_results.txt", 'r', encoding='utf-8') as f:
                        results = f.read()
                    with open(filename, 'w', encoding='utf-8') as f:
                        f.write(results)
                    messagebox.showinfo("Успех", "Результаты сохранены!")
                else:
                    messagebox.showwarning("Предупреждение", "Файл с результатами не найден.")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить результаты: {e}")


def main():
    """Основная функция"""
    root = tk.Tk()
    app = SearchApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()