import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import logging
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


class TextHandler(logging.Handler):
    """Кастомный обработчик для логирования в текстовое поле"""

    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        self.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

    def emit(self, record):
        msg = self.format(record)

        def append():
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.see(tk.END)
            self.text_widget.configure(state='disabled')

        # Вызываем в основном потоке
        self.text_widget.after(0, append)


class SearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Поиск файлов по ключевым словам")
        self.root.geometry("1000x800")
        self.root.minsize(900, 700)

        # Переменные для хранения состояний
        self.selected_extensions = tk.StringVar()
        self.directories_list = []
        self.is_searching = False
        self.search_thread = None
        self.progress_value = tk.DoubleVar(value=0.0)
        self.current_file = tk.StringVar(value="")
        self.total_files = 0
        self.processed_files = 0

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

        # Устанавливаем значение для текстового поле расширений
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
        main_frame.rowconfigure(6, weight=1)
        main_frame.rowconfigure(7, weight=1)

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

        # Добавляем директорию по умолчанию из config.txt
        default_directory = self.config['config'].get('directory', '.')
        self.directories_list.append(default_directory)
        self.dirs_listbox.insert(tk.END, default_directory)

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

        # Row 4: Progress
        ttk.Label(main_frame, text="Прогресс:").grid(row=4, column=0, sticky=tk.W, pady=5)

        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=5)

        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_value, maximum=100)
        self.progress_bar.pack(fill=tk.X, expand=True)

        ttk.Label(progress_frame, textvariable=self.current_file).pack(fill=tk.X)

        # Row 5: Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=10)

        self.start_button = ttk.Button(button_frame, text="Начать поиск", command=self.start_search)
        self.start_button.pack(side=tk.LEFT, padx=5)

        self.stop_button = ttk.Button(button_frame, text="Закончить поиск", command=self.stop_search, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="Сохранить результаты", command=self.save_results).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Очистить всё", command=self.clear_all).pack(side=tk.LEFT, padx=5)

        # Row 6: Results
        ttk.Label(main_frame, text="Результаты поиска:").grid(row=6, column=0, sticky=tk.NW, pady=5)

        results_frame = ttk.Frame(main_frame)
        results_frame.grid(row=6, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        self.results_text = scrolledtext.ScrolledText(results_frame, height=10)
        self.results_text.pack(fill=tk.BOTH, expand=True)
        self.results_text.configure(state='disabled')

        # Row 7: Log
        ttk.Label(main_frame, text="Лог выполнения:").grid(row=7, column=0, sticky=tk.NW, pady=5)

        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=7, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state='disabled')

        # Настраиваем логирование в текстовое поле
        self.setup_logging()

    def setup_logging(self):
        """Настройка логирования в текстовое поле"""
        # Очищаем существующие обработчики
        logger = logging.getLogger()
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)

        # Добавляем обработчик для текстового поля
        log_handler = TextHandler(self.log_text)
        log_handler.setLevel(logging.INFO)
        logger.addHandler(log_handler)
        logger.setLevel(logging.INFO)

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

    def clear_all(self):
        """Очистка всех полей"""
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')

        self.results_text.configure(state='normal')
        self.results_text.delete(1.0, tk.END)
        self.results_text.configure(state='disabled')

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
                elif file_name == "Поиск завершен":
                    progress_text = "Поиск завершен! Обработано всех файлов."
                else:
                    progress_text += f" | Текущий: {display_name}"

                self.current_file.set(progress_text)

        # Принудительно обновляем прогрессбар
        self.progress_bar.update_idletasks()

    def add_result(self, result_text):
        """Добавление результата в текстовое поле"""

        def append_result():
            self.results_text.configure(state='normal')
            self.results_text.insert(tk.END, result_text + '\n')
            self.results_text.see(tk.END)
            self.results_text.configure(state='disabled')

        self.root.after(0, append_result)

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

    def start_search(self):
        """Запуск поиска в отдельном потоке"""
        if self.is_searching:
            return

        # Сохраняем текущие значения перед обновлением конфига
        current_extensions = self.selected_extensions.get()

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
        self.stop_button.config(state=tk.NORMAL)
        self.is_searching = True
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
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            logging.info("Поиск остановлен пользователем")

            # Принудительно обновляем статус
            self.current_file.set("Поиск остановлен пользователем")

    def run_search(self, extensions, progress_callback):
        """Выполнение поиска"""
        try:
            # Сбрасываем только processed_files при начале нового поиска
            self.processed_files = 0
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
                    int(self.threads_var.get()),
                    "search_results.txt",
                    int(self.max_size_var.get()),
                    self.config['config'],
                    progress_callback,
                    self.processed_files  # Передаем текущее значение как offset
                )

                # Показываем результаты для текущей директории
                if results:
                    logging.info(f"Найдено совпадений в {len(results)} файлах в директории {directory}:")
                    for file_path, keywords in results.items():
                        result_text = f"Файл: {file_path}\nКлючевые слова: {', '.join(keywords)}\n"
                        logging.info(f"Файл: {file_path}")
                        logging.info(f"Ключевые слова: {', '.join(keywords)}")
                        self.add_result(result_text)
                else:
                    logging.info(f"В директории {directory} ничего не найдено.")

                # Обновляем счетчик обработанных файлов для этой директории
                files_in_dir = self.count_files_to_process(directory, extensions)
                logging.info(f"Обработано файлов в директории {directory}: {files_in_dir}")

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
            self.is_searching = False
            self.root.after(0, self.on_search_finished)

    def update_progress_callback(self, file_name, processed_count):
        """Callback для обновления прогресса из search_engine"""
        # Обновляем в основном потоке через after
        self.root.after(0, lambda: self._update_progress_in_main_thread(file_name, processed_count))

    def _update_progress_in_main_thread(self, file_name, processed_count):
        """Обновление прогресса в основном потоке"""
        self.processed_files = processed_count
        self.update_progress(file_name)

    def on_search_finished(self):
        """Вызывается при завершении поиска"""
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        # Не обновляем прогресс здесь, он уже обновлен в run_search

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
                # Сохраняем результаты из текстового поля
                self.results_text.configure(state='normal')
                results_content = self.results_text.get(1.0, tk.END)
                self.results_text.configure(state='disabled')

                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(results_content)
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