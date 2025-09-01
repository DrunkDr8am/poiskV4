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
        self.root.geometry("800x600")
        self.root.minsize(700, 500)

        # Переменные для хранения состояний
        self.extensions_vars = {}
        self.keywords_text = None
        self.directories_list = []
        self.is_searching = False
        self.search_thread = None

        # Загружаем конфигурацию ДО создания интерфейса
        self.config = self.load_configuration()

        # Создаем вкладки
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Вкладка настроек
        self.settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_frame, text="Настройки поиска")

        # Вкладка результатов
        self.results_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.results_frame, text="Результаты")

        self.setup_settings_tab()
        self.setup_results_tab()

        # Проверяем зависимости
        self.check_dependencies()

    def load_configuration(self):
        """Загрузка конфигурации"""
        if not os.path.exists("config.txt"):
            create_default_config()

        config = load_config()

        # Инициализируем переменные для расширений
        default_extensions = ['*.txt', '*.pdf', '*.docx', '*.xlsx', '*.jpg', '*.png', '*.zip', '*.rar', '*.7z']
        extensions_vars = {}
        for ext in default_extensions:
            extensions_vars[ext] = tk.BooleanVar(value=ext in config['extensions'])

        return {
            'config': config,
            'extensions_vars': extensions_vars
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

        # Обновляем состояние чекбоксов в зависимости от доступности функций
        if not HAS_PDF:
            self.extensions_vars['*.pdf'].set(False)
            self.pdf_check.config(state=tk.DISABLED)

        if not HAS_DOCX:
            self.extensions_vars['*.docx'].set(False)
            self.docx_check.config(state=tk.DISABLED)

        if not HAS_EXCEL:
            self.extensions_vars['*.xlsx'].set(False)
            self.xlsx_check.config(state=tk.DISABLED)

        if not HAS_7Z:
            self.extensions_vars['*.7z'].set(False)
            self.z7_check.config(state=tk.DISABLED)

        if not HAS_RAR:
            self.extensions_vars['*.rar'].set(False)
            self.rar_check.config(state=tk.DISABLED)

        if not HAS_OCR:
            self.extensions_vars['*.jpg'].set(False)
            self.jpg_check.config(state=tk.DISABLED)
            self.extensions_vars['*.png'].set(False)
            self.png_check.config(state=tk.DISABLED)

    def setup_settings_tab(self):
        """Настройка вкладки с параметрами поиска"""
        # Инициализируем переменные расширений
        self.extensions_vars = self.config['extensions_vars']
        config = self.config['config']

        # Фрейм для расширений файлов
        extensions_frame = ttk.LabelFrame(self.settings_frame, text="Расширения файлов для поиска")
        extensions_frame.pack(fill=tk.X, padx=10, pady=5)

        # Чекбоксы для расширений
        check_frame = ttk.Frame(extensions_frame)
        check_frame.pack(fill=tk.X, padx=5, pady=5)

        self.txt_check = ttk.Checkbutton(check_frame, text="TXT", variable=self.extensions_vars['*.txt'])
        self.txt_check.grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)

        self.pdf_check = ttk.Checkbutton(check_frame, text="PDF", variable=self.extensions_vars['*.pdf'])
        self.pdf_check.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)

        self.docx_check = ttk.Checkbutton(check_frame, text="DOCX", variable=self.extensions_vars['*.docx'])
        self.docx_check.grid(row=0, column=2, sticky=tk.W, padx=5, pady=2)

        self.xlsx_check = ttk.Checkbutton(check_frame, text="XLSX", variable=self.extensions_vars['*.xlsx'])
        self.xlsx_check.grid(row=0, column=3, sticky=tk.W, padx=5, pady=2)

        self.jpg_check = ttk.Checkbutton(check_frame, text="JPG", variable=self.extensions_vars['*.jpg'])
        self.jpg_check.grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)

        self.png_check = ttk.Checkbutton(check_frame, text="PNG", variable=self.extensions_vars['*.png'])
        self.png_check.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)

        self.zip_check = ttk.Checkbutton(check_frame, text="ZIP", variable=self.extensions_vars['*.zip'])
        self.zip_check.grid(row=1, column=2, sticky=tk.W, padx=5, pady=2)

        self.rar_check = ttk.Checkbutton(check_frame, text="RAR", variable=self.extensions_vars['*.rar'])
        self.rar_check.grid(row=1, column=3, sticky=tk.W, padx=5, pady=2)

        self.z7_check = ttk.Checkbutton(check_frame, text="7Z", variable=self.extensions_vars['*.7z'])
        self.z7_check.grid(row=1, column=4, sticky=tk.W, padx=5, pady=2)

        # Фрейм для ключевых слов
        keywords_frame = ttk.LabelFrame(self.settings_frame, text="Ключевые слова для поиска")
        keywords_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.keywords_text = scrolledtext.ScrolledText(keywords_frame, height=6)
        self.keywords_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Загружаем ключевые слова из файла, если он существует
        if os.path.exists("keywords.txt"):
            try:
                with open("keywords.txt", "r", encoding="utf-8") as f:
                    keywords = f.read()
                    self.keywords_text.insert("1.0", keywords)
            except:
                pass

        # Фрейм для выбора директорий
        directories_frame = ttk.LabelFrame(self.settings_frame, text="Директории для поиска")
        directories_frame.pack(fill=tk.X, padx=10, pady=5)

        # Список директорий
        self.directories_listbox = tk.Listbox(directories_frame, height=4)
        self.directories_listbox.pack(fill=tk.X, padx=5, pady=5)

        # Кнопки для управления директориями
        dir_buttons_frame = ttk.Frame(directories_frame)
        dir_buttons_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(dir_buttons_frame, text="Добавить директорию", command=self.add_directory).pack(side=tk.LEFT, padx=5)
        ttk.Button(dir_buttons_frame, text="Удалить выбранную", command=self.remove_directory).pack(side=tk.LEFT,
                                                                                                    padx=5)

        # Добавляем текущую директорию по умолчанию
        self.directories_list.append(".")
        self.directories_listbox.insert(tk.END, ".")

        # Фрейм для дополнительных настроек
        settings_frame = ttk.LabelFrame(self.settings_frame, text="Дополнительные настройки")
        settings_frame.pack(fill=tk.X, padx=10, pady=5)

        # Потоки
        ttk.Label(settings_frame, text="Количество потоков:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.threads_var = tk.StringVar(value=str(config.get('threads', 4)))
        threads_spin = ttk.Spinbox(settings_frame, from_=1, to=16, textvariable=self.threads_var, width=5)
        threads_spin.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)

        # Максимальный размер файла
        ttk.Label(settings_frame, text="Макс. размер файла (МБ):").grid(row=0, column=2, sticky=tk.W, padx=5, pady=2)
        self.max_size_var = tk.StringVar(value=str(config.get('max_file_size', 50)))
        max_size_spin = ttk.Spinbox(settings_frame, from_=1, to=1000, textvariable=self.max_size_var, width=5)
        max_size_spin.grid(row=0, column=3, sticky=tk.W, padx=5, pady=2)

        # Поиск по изображениям
        self.search_images_var = tk.BooleanVar(value=config.get('search_images', False))
        search_images_check = ttk.Checkbutton(settings_frame, text="Поиск по изображениям (OCR)",
                                              variable=self.search_images_var)
        search_images_check.grid(row=0, column=4, sticky=tk.W, padx=5, pady=2)

        # Кнопки запуска и остановки
        button_frame = ttk.Frame(self.settings_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=10)

        self.start_button = ttk.Button(button_frame, text="Начать поиск", command=self.start_search)
        self.start_button.pack(side=tk.LEFT, padx=5)

        self.stop_button = ttk.Button(button_frame, text="Остановить", command=self.stop_search, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)

    def setup_results_tab(self):
        """Настройка вкладки с результатами"""
        # Текстовое поле для результатов
        self.results_text = scrolledtext.ScrolledText(self.results_frame, wrap=tk.WORD)
        self.results_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Кнопка сохранения результатов
        save_button = ttk.Button(self.results_frame, text="Сохранить результаты", command=self.save_results)
        save_button.pack(pady=5)

    def add_directory(self):
        """Добавление директории для поиска"""
        directory = filedialog.askdirectory(title="Выберите директорию для поиска")
        if directory:
            self.directories_list.append(directory)
            self.directories_listbox.insert(tk.END, directory)

    def remove_directory(self):
        """Удаление выбранной директории"""
        selection = self.directories_listbox.curselection()
        if selection:
            index = selection[0]
            self.directories_listbox.delete(index)
            del self.directories_list[index]

    def start_search(self):
        """Запуск поиска в отдельном потоке"""
        if self.is_searching:
            return

        # Получаем выбранные расширения
        extensions = []
        for ext, var in self.extensions_vars.items():
            if var.get():
                extensions.append(ext)

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

        # Очищаем результаты
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, "Поиск начат...\n")

        # Меняем состояние кнопок
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.is_searching = True

        # Запускаем поиск в отдельном потоке
        self.search_thread = threading.Thread(target=self.run_search)
        self.search_thread.daemon = True
        self.search_thread.start()

        # Запускаем мониторинг прогресса
        self.monitor_progress()

    def stop_search(self):
        """Остановка поиска"""
        if self.is_searching:
            self.is_searching = False
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            self.results_text.insert(tk.END, "Поиск остановлен пользователем\n")

    def run_search(self):
        """Выполнение поиска"""
        try:
            # Получаем выбранные расширения
            extensions = []
            for ext, var in self.extensions_vars.items():
                if var.get():
                    extensions.append(ext)

            # Выполняем поиск для каждой директории
            for directory in self.directories_list:
                if not self.is_searching:
                    break

                logging.info(f"Поиск в директории: {directory}")
                self.results_text.insert(tk.END, f"Поиск в директории: {directory}\n")
                self.results_text.see(tk.END)

                results = search_files(
                    directory,
                    extensions,
                    int(self.threads_var.get()),
                    "search_results.txt",
                    int(self.max_size_var.get()),
                    self.config['config']
                )

                if results:
                    self.results_text.insert(tk.END, f"Найдено совпадений в {len(results)} файлах:\n")
                    for file_path, keywords in results.items():
                        self.results_text.insert(tk.END, f"Файл: {file_path}\n")
                        self.results_text.insert(tk.END, f"Ключевые слова: {', '.join(keywords)}\n\n")
                else:
                    self.results_text.insert(tk.END, "Ничего не найдено.\n")

                self.results_text.see(tk.END)

            if self.is_searching:
                self.results_text.insert(tk.END, "Поиск завершен!\n")

        except Exception as e:
            logging.error(f"Ошибка при поиске: {e}")
            self.results_text.insert(tk.END, f"Ошибка: {e}\n")

        finally:
            self.is_searching = False
            self.root.after(0, self.on_search_finished)

    def on_search_finished(self):
        """Вызывается при завершении поиска"""
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.results_text.see(tk.END)

    def monitor_progress(self):
        """Мониторинг прогресса поиска"""
        if self.is_searching:
            # Здесь можно добавить логику для отображения прогресса
            self.root.after(1000, self.monitor_progress)

    def update_config(self):
        """Обновление конфигурации"""
        config = ConfigParser()

        # Получаем выбранные расширения
        extensions = []
        for ext, var in self.extensions_vars.items():
            if var.get():
                extensions.append(ext)

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

        # Обновляем self.config
        self.config = self.load_configuration()

    def save_results(self):
        """Сохранение результатов в файл"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.results_text.get(1.0, tk.END))
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