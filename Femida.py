import sys
import os
import pandas as pd
import numpy as np
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout,
    QWidget, QFileDialog, QLabel, QLineEdit, QTabWidget,
    QTableView, QMessageBox, QStatusBar, QHBoxLayout,
    QRadioButton, QButtonGroup, QGroupBox, QHeaderView,
    QDialog, QVBoxLayout as QDialogVBoxLayout
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt, QSortFilterProxyModel, QPoint
from PyQt6.QtGui import QStandardItemModel, QStandardItem, QIcon, QColor, QBrush # Added QColor, QBrush
from datetime import datetime
import traceback
from functools import partial
import re


# Имя папки для кэша данных Excel
CACHE_DIR_NAME = "excel_cache_app"
# Имя папки для конфигурации приложения (в домашней папке пользователя)
CONFIG_DIR_NAME = ".excel_search_config"
# Имя файла для хранения пути к последней папки
LAST_FOLDER_FILE = "last_folder.txt"




# --- Рабочий поток для фоновых задач ---
class Worker(QThread):
    finished = pyqtSignal()
    progress = pyqtSignal(str)
    error = pyqtSignal(str)
    data_loaded_signal = pyqtSignal(dict)
    search_results_signal = pyqtSignal(dict)


    def __init__(self, task_type, folder_path=None, search_term=None, loaded_data=None, search_mode="normal"):
        super().__init__()
        self.task_type = task_type
        self.folder_path = folder_path
        self.search_term = search_term
        self.loaded_data = loaded_data
        self.search_mode = search_mode


    def run(self):
        try:
            if self.task_type == "load_data":
                if not self.folder_path:
                    self.error.emit("Папка не выбрана.")
                    self.finished.emit()
                    return


                cache_path_main = os.path.join("C:\\", CACHE_DIR_NAME)
                try:
                    os.makedirs(cache_path_main, exist_ok=True)
                except OSError as e:
                    self.error.emit(f"Не удалось создать папку кэша C:\\{CACHE_DIR_NAME}: {e}. [cite: 5]")
                    self.finished.emit()
                    return




                all_data = {}
                excel_files = [f for f in os.listdir(self.folder_path)
                               if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')
                               and not os.path.isdir(os.path.join(self.folder_path, f))]


                if not excel_files:
                    self.progress.emit("Excel файлы не найдены в указанной папке.")
                    self.data_loaded_signal.emit({})
                    self.finished.emit()
                    return


                for i, filename in enumerate(excel_files):
                    if not self.isRunning():
                        self.progress.emit("Загрузка прервана.")
                        return


                    self.progress.emit(f"Обработка файла: {filename} ({i+1}/{len(excel_files)})...")
                    file_path = os.path.join(self.folder_path, filename)


                    safe_filename_for_cache = "".join(c if c.isalnum() else "_" for c in os.path.splitext(filename)[0])
                    cache_file_path = os.path.join(cache_path_main, safe_filename_for_cache + ".pkl")


                    load_from_excel = True
                    if os.path.exists(cache_file_path):
                        try:
                            file_mod_time = os.path.getmtime(file_path)
                            cache_mod_time = os.path.getmtime(cache_file_path)
                            if file_mod_time <= cache_mod_time:
                                self.progress.emit(f"Загрузка {filename} из кэша...")
                                df = pd.read_pickle(cache_file_path)
                                all_data[filename] = df
                                load_from_excel = False
                        except Exception as e:
                            self.progress.emit(f"Ошибка чтения кэша для {filename}, перечитываем Excel: {e}")
                            if os.path.exists(cache_file_path):
                                try:
                                    os.remove(cache_file_path)
                                except OSError as oe:
                                    self.progress.emit(f"Не удалось удалить поврежденный кэш-файл {cache_file_path}: {oe}")
                            load_from_excel = True


                    if load_from_excel:
                        self.progress.emit(f"Чтение {filename} из Excel...")
                        try:
                            engine = 'openpyxl' if filename.endswith('.xlsx') else 'xlrd' if filename.endswith('.xls') else None
                            df = pd.read_excel(file_path, engine=engine, parse_dates=True)


                            if not df.empty:
                                df.columns = [str(col) for col in df.columns]


                                initial_rows, initial_cols = df.shape
                                cols_to_drop = [col for col in df.columns if not str(col).strip() or str(col).startswith('Unnamed:')]
                                cols_removed = 0
                                if cols_to_drop:
                                    actual_cols_to_drop = [col for col in cols_to_drop if col in df.columns]
                                    if actual_cols_to_drop:
                                        df.drop(columns=actual_cols_to_drop, inplace=True)
                                        cols_removed = len(actual_cols_to_drop)
                                        self.progress.emit(f"Удалено {cols_removed} пустых столбцов или столбцов без наименований ('Unnamed:') в {filename}.")


                                initial_rows_before_row_drop = df.shape[0]
                                df.dropna(axis=0, how='all', inplace=True)
                                rows_removed = initial_rows_before_row_drop - df.shape[0]
                                if rows_removed > 0:
                                    self.progress.emit(f"Удалено {rows_removed} полностью пустых строк в {filename}.")


                                all_data[filename] = df
                                try:
                                    df.to_pickle(cache_file_path)
                                except Exception as e_pickle:
                                    self.error.emit(f"Ошибка сохранения кэша для {filename} в {cache_file_path}: {e_pickle}")
                            else:
                                self.progress.emit(f"Файл {filename} пуст или не удалось прочитать первую строку как заголовки.")
                        except Exception as e:
                            self.error.emit(f"Ошибка чтения {filename}: {e}")
                            if os.path.exists(cache_file_path):
                                try:
                                    os.remove(cache_file_path)
                                except OSError as oe: # 
                                    self.progress.emit(f"Не удалось удалить кэш-файл {cache_file_path} после ошибки чтения: {oe}")


                self.progress.emit("Загрузка данных завершена.")
                self.data_loaded_signal.emit(all_data)


            elif self.task_type == "search_data":
                if not self.search_term and self.search_mode == "normal":
                     self.error.emit("Введите текст для обычного поиска.")
                     self.finished.emit()
                     return


                if self.loaded_data is None or not self.loaded_data:
                    self.error.emit("Данные не загружены. Сначала выберите и загрузите папку.")
                    self.finished.emit()
                    return


                results = {}
                search_term_lower = self.search_term.lower()
                num_files = len(self.loaded_data)
                count = 0
                for filename, df in self.loaded_data.items():
                    if not self.isRunning():
                        self.progress.emit("Поиск прерван.")
                        return
                    count += 1
                    self.progress.emit(f"Поиск в {filename} ({count}/{num_files})...")
                    if df is None or df.empty:
                        continue


                    try:
                        df_temp = df.copy()


                        for col in df_temp.columns:
                            if pd.api.types.is_datetime64_any_dtype(df_temp[col]):
                                df_temp.loc[:, col] = df_temp[col].dt.strftime('%d.%m.%Y')


                        if self.search_mode == "normal":
                            mask = df_temp.astype(str).apply(lambda x: x.str.lower().str.contains(search_term_lower, na=False))
                        elif self.search_mode == "strict":
                            df_filled = df_temp.fillna('')
                            mask = df_filled.astype(str).apply(lambda x: x.str.lower() == search_term_lower)


                        matching_rows = df[mask.any(axis=1)].copy()


                        if not matching_rows.empty:
                            results[filename] = matching_rows
                    except Exception as e:
                        self.error.emit(f"Ошибка при поиске в {filename}: {e}")


                self.progress.emit("Поиск завершен.")
                self.search_results_signal.emit(results)


        except Exception as e:
            self.error.emit(f"Произошла критическая ошибка в потоке Worker: {e}")
            traceback.print_exc()
        finally:
            self.finished.emit()


# --- Главное окно приложения ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Femida")
        self.setGeometry(100, 100, 1000, 700)


        try:
            icon_path = 'Femida.png'
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
            else:
                print(f"Файл иконки не был найден {icon_path}")
        except Exception as e:
            print(f"Ошибка установки иконки окна: {e}") # Modified from original print
            
        self.current_folder_path = ""
        self.loaded_data = {}
        self.current_search_mode = "normal"
        self.proxy_models = {} # Словарь для хранения proxy моделей для каждой вкладки: {filename: proxy_model}




        self.config_dir = os.path.join(os.path.expanduser("~"), CONFIG_DIR_NAME)
        self.last_folder_path_file = os.path.join(self.config_dir, LAST_FOLDER_FILE)


        self.folder_select_button = QPushButton("Выбрать/Обновить папку")
        self.folder_select_button.clicked.connect(self.select_folder_and_load_data)


        self.selected_folder_label = QLabel("Папка не выбрана")
        self.selected_folder_label.setWordWrap(True)


        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Введите текст для поиска...")
        self.search_input.returnPressed.connect(self.start_search)


        self.search_button = QPushButton("Поиск")
        self.search_button.clicked.connect(self.start_search)


        self.search_mode_group_box = QGroupBox("Тип поиска")
        self.search_mode_layout = QHBoxLayout()


        self.normal_search_radio = QRadioButton("Обычный (по вхождению)")
        self.strict_search_radio = QRadioButton("Строгий (полное совпадение)")


        self.search_mode_button_group = QButtonGroup(self)
        self.search_mode_button_group.addButton(self.normal_search_radio)
        self.search_mode_button_group.addButton(self.strict_search_radio)


        self.search_mode_layout.addWidget(self.normal_search_radio)
        self.search_mode_layout.addWidget(self.strict_search_radio)
        self.search_mode_layout.addStretch(1)


        self.search_mode_group_box.setLayout(self.search_mode_layout)


        self.normal_search_radio.setChecked(True)
        self.update_search_mode()
        self.normal_search_radio.toggled.connect(self.update_search_mode)
        self.strict_search_radio.toggled.connect(self.update_search_mode)


        self.results_tabs = QTabWidget()
        self.results_tabs.setTabsClosable(True)
        self.results_tabs.tabCloseRequested.connect(self.close_tab)
        self.results_tabs.tabCloseRequested.connect(self.remove_proxy_model_for_tab)




        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)


        main_layout = QVBoxLayout()


        folder_layout = QHBoxLayout()
        folder_layout.addWidget(self.folder_select_button)
        folder_layout.addWidget(self.selected_folder_label, 1)


        search_input_button_layout = QHBoxLayout()
        search_input_button_layout.addWidget(self.search_input, 1)
        search_input_button_layout.addWidget(self.search_button)


        main_layout.addLayout(folder_layout)
        main_layout.addWidget(self.search_mode_group_box)
        main_layout.addLayout(search_input_button_layout)
        main_layout.addWidget(self.results_tabs)


        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)


        self.worker_thread = None


        self.load_last_folder_on_startup()


    # --- Обновленный слот для обработки двойного клика по заголовку с дополнительными print() ---
    def on_header_double_clicked(self, logicalIndex):
        """Обрабатывает двойной клик по заголовку столбца для вызова фильтра.""" # 
        print(f"on_header_double_clicked called for column {logicalIndex}.") # Debug print
        try: # Add a try-except block around the entire method
            # Получаем текущую активную вкладку и ее widget
            print("Getting current tab widget...") # Debug print
            current_tab_widget = self.results_tabs.currentWidget()
            if not current_tab_widget:
                print("No active tab.")
                return
            print(f"Current tab widget: {current_tab_widget}") # Debug print


            # Находим QTableView внутри активной вкладки
            print("Finding QTableView in current tab...") # Debug print
            table_view = current_tab_widget.findChild(QTableView) # 
            if not table_view:
                print("No QTableView found in the current tab.")
                return
            print(f"Found QTableView: {table_view}") # Debug print


            # Получаем proxy model, связанную с этой таблицей
            print("Getting model from QTableView...") # Debug print
            proxy_model = table_view.model()
            if not isinstance(proxy_model, QSortFilterProxyModel):
                 print("Model is not a QSortFilterProxyModel.")
                 return
            print(f"Found Proxy Model: {proxy_model}") # Debug print




            # Получаем исходную модель, чтобы получить имена столбцов
            print("Getting source model from proxy model...") # Debug print
            source_model = proxy_model.sourceModel()
            if not source_model:
                print("Proxy model has no source model.")
                return
            print(f"Found Source Model: {source_model}") # Debug print




            col_name = source_model.horizontalHeaderItem(logicalIndex).text() if source_model.horizontalHeaderItem(logicalIndex) else f"Column {logicalIndex}"
            print(f"Column name for index {logicalIndex}: {col_name}") # Debug print


            # --- Добавляем еще один print() сразу после получения имени столбца ---
            print(">>> About to create filter dialog...") # Этот print должен либо появиться перед вылетом, либо сбой происходит на нем
            # --- Конец добавленного print() ---




            # Создаем всплывающее окно для фильтра
            print(">>> Starting filter dialog creation block...") # Debug print
            try: # Add specific try-except for dialog creation
                print("Attempting to create QDialog instance...") # Debug print
                filter_dialog = QDialog(self)
                print("QDialog instance created.") # Debug print


                print("Attempting to set dialog window title...") # Debug print
                filter_dialog.setWindowTitle(f"Фильтр для '{col_name}'")
                print("Dialog window title set.") # Debug print


                print("Attempting to create dialog layout...") # Debug print
                dialog_layout = QDialogVBoxLayout(filter_dialog)
                print("Dialog layout created.") # Debug print


                print("Attempting to create filter input LINEEDIT...") # Debug print
                filter_input = QLineEdit(filter_dialog)
                filter_input.setPlaceholderText("Введите текст фильтра...")
                print("Filter input LINEEDIT created.") # Debug print


                print("Attempting to add filter input to dialog layout...") # Debug print
                dialog_layout.addWidget(filter_input)
                print("Filter input added to dialog layout.") # Debug print


            except Exception as e:
                print(f"Error during filter dialog creation steps: {e}")
                traceback.print_exc()
                QMessageBox.critical(self, "Ошибка создания диалога", f"Произошла ошибка при создании окна фильтра: {e}")
                return # Stop processing if dialog creation fails




            # Подключаем сигнал textChanged поля ввода к слоту фильтрации
            print("Connecting filter input signal to filter_table_data...") # Debug print
            filter_input.textChanged.connect(partial(self.filter_table_data, col_index=logicalIndex, proxy_model=proxy_model))
            print("Filter input signal connected.") # Debug print




            # Связываем сигнал accepted (например, по нажатию Enter или кнопки OK) с закрытием диалога.
            print("Connecting returnPressed signal to dialog accept...") # Debug print
            filter_input.returnPressed.connect(filter_dialog.accept)
            print("ReturnPressed signal connected.") # Debug print




            # Позиционируем диалог рядом с заголовком столбца
            print("Positioning filter dialog...") # Debug print
            try: # Add specific try-except for positioning
                header = table_view.horizontalHeader()
                # Получаем глобальные координаты заголовка столбца
                local_x_pos = header.sectionViewportPosition(logicalIndex)
                point_in_header_coords = QPoint(local_x_pos, 0)
                header_pos = header.mapToGlobal(point_in_header_coords)


                # Корректируем позицию, чтобы диалог появился ниже заголовка
                dialog_pos = header_pos + QPoint(0, header.height())
                filter_dialog.move(dialog_pos)
                print(f"Filter dialog positioned at {dialog_pos.x()},{dialog_pos.y()}.") # Debug print
            except Exception as e:
                 print(f"Error calculating or setting dialog position: {e}")
                 traceback.print_exc()
                 print("Falling back to default position (100, 100).")
                 filter_dialog.move(100, 100) # Fallback




            # Запускаем диалог
            print("Executing filter dialog...") # Debug print
            try: # Add specific try-except for dialog execution
                filter_dialog.exec()
                print("Filter dialog finished execution.") # Debug print
            except Exception as e:
                print(f"Error during filter dialog execution: {e}")
                traceback.print_exc()
                QMessageBox.critical(self, "Ошибка выполнения диалога", f"Произошла ошибка во время работы окна фильтра: {e}")




            print("Filter dialog closed.") # Debug print
            # Когда диалог закрыт, фильтр остается примененным через textChanged


        except Exception as e: # Outer catch for any unexpected errors
            print(f"An unexpected error occurred in on_header_double_clicked: {e}")
            traceback.print_exc()
            QMessageBox.critical(self, "Неизвестная ошибка фильтрации", f"Произошла неизвестная ошибка при попытке фильтрации: {e}")




    # --- Новый слот для удаления proxy модели при закрытии вкладки ---
    def remove_proxy_model_for_tab(self, index):
        """Удаляет proxy модель из словаря при закрытии вкладки."""
        print(f"remove_proxy_model_for_tab called for index {index}.") # Debug print
        # Получаем виджет закрываемой вкладки
        tab_widget = self.results_tabs.widget(index) # 
        if tab_widget:
            print(f"Getting QTableView from tab widget: {tab_widget}") # Debug print
            # Находим QTableView в закрываемой вкладке
            table_view = tab_widget.findChild(QTableView)
            if table_view:
                print(f"Found QTableView: {table_view}. Getting its model...") # Debug print
                # Получаем модель (которая должна быть proxy_model)
                proxy_model = table_view.model()
                if isinstance(proxy_model, QSortFilterProxyModel): # 
                    print(f"Found Proxy Model: {proxy_model} (id: {id(proxy_model)}).") # Debug print
                    # Находим имя файла, связанное с этой proxy моделью в нашем словаре proxy_models
                    filename_to_remove = None
                    # Ищем по значению (экземпляру proxy_model) в словаре
                    # Итерируем по копии словаря на случай изменений во время итерации (хотя здесь вряд ли) # 
                    for fname, p_model in list(self.proxy_models.items()): # 
                        if p_model is proxy_model:
                            filename_to_remove = fname
                            break


                    if filename_to_remove in self.proxy_models: # 
                        print(f"Removing proxy model for {filename_to_remove} (id: {id(self.proxy_models[filename_to_remove])})...")
                        # Отключаем источник перед удалением proxy модели
                        try: # 
                             proxy_model.setSourceModel(None)
                             print(f"Source model disconnected for {filename_to_remove}.")
                        except Exception as e: # 
                             print(f"Error disconnecting source model for {filename_to_remove}: {e}")


                        # Удаляем из словаря
                        del self.proxy_models[filename_to_remove]
                        print(f"Proxy model for {filename_to_remove} removed from dictionary. [cite: 121]")
                        print(f"Remaining models: {len(self.proxy_models)}") # Debug print 


                    # Очистка самой proxy модели. deleteLater() безопаснее в Qt. [cite: 73]
                    try:
                        # Проверяем, что модель еще существует перед вызовом deleteLater
                        if proxy_model:
                             print(f"Calling deleteLater() on proxy model (id: {id(proxy_model)})...") # Debug print
                             proxy_model.deleteLater()
                             print(f"Proxy model deleteLater() called for {filename_to_remove or 'unknown file'}.")
                        else:
                            print("Proxy model is None, cannot call deleteLater().") # Debug print


                    except Exception as e:
                        print(f"Error calling deleteLater() on proxy model: {e}") # 
                        traceback.print_exc()


            else:
                 print("No QTableView found in the closing tab.") # Debug print
        else:
            print(f"No widget found for tab index {index}.") # Debug print




    def filter_table_data(self, text, col_index, proxy_model):
        """Применяет фильтр к указанному столбцу proxy_model."""
        print(f"Applying filter '{text}' to column {col_index}")
        try:
            if proxy_model and 0 <= col_index < proxy_model.columnCount(): # Добавлена проверка proxy_model
                proxy_model.setFilterKeyColumn(col_index)
                escaped_text = re.escape(text)
                proxy_model.setFilterRegularExpression(escaped_text)
                print(f"Filter applied. Rows after filtering: {proxy_model.rowCount()}")
            else:
                 print(f"Warning: Invalid column index {col_index} or proxy_model is None for filtering.")
        except Exception as e:
            print(f"Error in filter_table_data: {e}")
            traceback.print_exc()




    def update_search_mode(self):
        if self.normal_search_radio.isChecked():
            self.current_search_mode = "normal"
            self.search_input.setPlaceholderText("Введите текст для поиска (по вхождению)...")
        elif self.strict_search_radio.isChecked():
            self.current_search_mode = "strict"
            self.search_input.setPlaceholderText("Введите текст для поиска (полное совпадение, пустой запрос найдет пустые ячейки)...")


    def load_last_folder_on_startup(self):
        """Пытается загрузить последнюю папку при запуске приложения."""
        initial_folder = ""
        if os.path.exists(self.last_folder_path_file):
            try:
                with open(self.last_folder_path_file, 'r', encoding='utf-8') as f:
                    content = f.read().strip()
                    if os.path.isdir(content):
                        initial_folder = content
            except Exception as e:
                print(f"Ошибка чтения файла последней папки '{self.last_folder_path_file}': {e}")


        if initial_folder:
            self.current_folder_path = initial_folder
            self.selected_folder_label.setText(f"Выбранная папка: {self.current_folder_path} (автозагрузка)")
            self.status_bar.showMessage("Автоматическая загрузка данных из последней папки...")
            self.set_buttons_enabled(False)
            self.worker_thread = Worker(task_type="load_data", folder_path=self.current_folder_path)
            self.worker_thread.finished.connect(self.on_worker_finished)
            self.worker_thread.progress.connect(self.update_status_bar)
            self.worker_thread.error.connect(self.show_error_message)
            self.worker_thread.data_loaded_signal.connect(self.on_data_loaded)
            self.worker_thread.start()
        else:
            self.status_bar.showMessage("Выберите папку с Excel файлами.")




    def select_folder_and_load_data(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Выберите папку с Excel файлами")
        if folder_path:
            if os.path.normpath(self.current_folder_path) == os.path.normpath(folder_path) and self.loaded_data:
                 reply = QMessageBox.question(self, "Папка уже загружена",
                                              "Эта папка уже загружена.\nХотите обновить данные из этой папки?",
                                              QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) # 
                 if reply == QMessageBox.StandardButton.No:
                     return


            if self.worker_thread and self.worker_thread.isRunning():
                QMessageBox.warning(self, "Операция выполняется", "Дождитесь завершения текущей операции перед выбором новой папки.")
                return


            self.current_folder_path = folder_path
            self.selected_folder_label.setText(f"Выбранная папка: {self.current_folder_path}")
            self.status_bar.showMessage("Загрузка данных...")
            self.loaded_data = {}
            self.results_tabs.clear()
            self.proxy_models.clear() # Очищаем proxy модели при загрузке новой папки


            try:
                os.makedirs(self.config_dir, exist_ok=True)
                with open(self.last_folder_path_file, 'w', encoding='utf-8') as f:
                    f.write(self.current_folder_path)
            except Exception as e:
                print(f"Ошибка сохранения файла последней папки '{self.last_folder_path_file}': {e}")




            self.set_buttons_enabled(False)


            self.worker_thread = Worker(task_type="load_data", folder_path=self.current_folder_path)
            self.worker_thread.finished.connect(self.on_worker_finished)
            self.worker_thread.progress.connect(self.update_status_bar)
            self.worker_thread.error.connect(self.show_error_message)
            self.worker_thread.data_loaded_signal.connect(self.on_data_loaded)
            self.worker_thread.start()


    def on_data_loaded(self, data):
        self.loaded_data = data
        if not data:
            self.status_bar.showMessage("Excel файлы не найдены или не удалось загрузить данные.")
        else:
            self.status_bar.showMessage(f"Загружено {len(data)} файлов. Готово к поиску.")


    def start_search(self):
        search_term = self.search_input.text().strip()


        if not search_term and self.current_search_mode == "normal":
            QMessageBox.warning(self, "Пустой поиск", "Пожалуйста, введите текст для обычного поиска.")
            return


        if self.loaded_data is None or not self.loaded_data:
            QMessageBox.warning(self, "Данные не загружены", "Сначала выберите и загрузите папку.")
            return
        if self.worker_thread and self.worker_thread.isRunning():
            QMessageBox.warning(self, "Операция выполняется", "Дождитесь завершения текущей операции перед началом нового поиска.")
            return


        self.status_bar.showMessage(f"Поиск '{search_term}' ({'строгий' if self.current_search_mode == 'strict' else 'обычный'})...")
        self.results_tabs.clear()
        self.proxy_models.clear() # Очищаем proxy модели перед новым поиском


        self.set_buttons_enabled(False)


        self.worker_thread = Worker(task_type="search_data",
                                    search_term=search_term,
                                    loaded_data=self.loaded_data,
                                    search_mode=self.current_search_mode)
        self.worker_thread.finished.connect(self.on_worker_finished)
        self.worker_thread.progress.connect(self.update_status_bar)
        self.worker_thread.error.connect(self.show_error_message)
        self.worker_thread.search_results_signal.connect(self.display_search_results)
        self.worker_thread.start()


    def display_search_results(self, results):
        print("display_search_results called")
        if not results:
            self.status_bar.showMessage("Совпадений не найдено.")
            QMessageBox.information(self, "Результаты поиска", "Совпадений не найдено.")
            return

        self.status_bar.showMessage(f"Найдено совпадений в {len(results)} файлах.")
        
        # Get the search term and mode for highlighting
        search_term_for_highlight = self.search_input.text().strip().lower()
        current_search_mode_for_highlight = self.current_search_mode
        
        calm_green_brush = QBrush(QColor("#C8E6C9")) # Soft green

        for filename, df_matches in results.items():
            print(f"Processing results for file: {filename}")
            try:
                if df_matches.empty:
                    print(f"DataFrame for {filename} is empty.")
                    continue

                tab = QWidget()
                tab_layout = QVBoxLayout(tab)

                table_view = QTableView()
                print(f"Creating QStandardItemModel for {filename}...")
                model = QStandardItemModel(df_matches.shape[0], df_matches.shape[1])
                model.setHorizontalHeaderLabels(df_matches.columns.tolist())
                print(f"QStandardItemModel created for {filename}.")

                print(f"Creating QSortFilterProxyModel for {filename}...")
                proxy_model = QSortFilterProxyModel()
                print(f"Setting source model for proxy model for {filename}... (Model: {model})") 
                proxy_model.setSourceModel(model)
                print(f"Source model set. Setting filter case sensitivity for {filename}...")
                proxy_model.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
                print(f"Filter case sensitivity set for {filename}.")

                self.proxy_models[filename] = proxy_model
                print(f"Proxy model stored for {filename} at {id(proxy_model)}.") 

                header = table_view.horizontalHeader()
                header.sectionDoubleClicked.connect(self.on_header_double_clicked)
                print(f"Header double click signal connected for {filename}.")

                # --- Конец настройки Proxy Model ---

                print(f"Populating QStandardItemModel with data for {filename}...")
                try:
                    df_columns = df_matches.columns.tolist()
                    for j, col_name in enumerate(df_columns):
                        is_datetime_col = pd.api.types.is_datetime64_any_dtype(df_matches[col_name])

                        for i in range(df_matches.shape[0]):
                            cell_value = df_matches.iloc[i, j]
                            item_text = ""

                            if pd.isna(cell_value) or (isinstance(cell_value, str) and cell_value.strip() == ''):
                                item_text = ""
                            elif is_datetime_col and isinstance(cell_value, pd.Timestamp):
                                try:
                                    item_text = cell_value.strftime('%d.%m.%Y')
                                except ValueError:
                                    item_text = str(cell_value)
                            elif is_datetime_col and pd.isna(cell_value):
                                item_text = ""
                            elif isinstance(cell_value, (int, float, np.number)) and pd.notna(cell_value) and np.isfinite(cell_value):
                                if pd.api.types.is_integer_dtype(df_matches.dtypes[j]):
                                     item_text = str(int(cell_value))
                                elif isinstance(cell_value, float) and cell_value.is_integer():
                                    item_text = str(int(cell_value))
                                else:
                                    item_text = str(cell_value)
                            else:
                                item_text = str(cell_value)

                            item = QStandardItem(item_text)
                            item.setEditable(False)

                            # --- Apply highlighting ---
                            item_text_lower = item_text.lower()
                            should_highlight = False
                            if current_search_mode_for_highlight == "normal":
                                if search_term_for_highlight and search_term_for_highlight in item_text_lower:
                                    should_highlight = True
                            elif current_search_mode_for_highlight == "strict":
                                if search_term_for_highlight == "" and item_text_lower == "": 
                                    should_highlight = True
                                elif search_term_for_highlight != "" and item_text_lower == search_term_for_highlight:
                                    should_highlight = True
                            
                            if should_highlight:
                                item.setBackground(calm_green_brush)
                            # --- End highlighting ---

                            model.setItem(i, j, item)
                except Exception as e:
                    print(f"Error populating QStandardItemModel for {filename}: {e}")
                    traceback.print_exc()
                    continue

                print(f"QStandardItemModel populated for {filename}. Setting proxy model to QTableView...")
                try:
                    table_view.setModel(proxy_model)
                    print(f"Proxy model set for QTableView for {filename}. TableView model: {table_view.model()}.") 
                except Exception as e:
                    print(f"Error setting proxy model for QTableView for {filename}: {e}")
                    traceback.print_exc()
                    continue

                header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
                header.setStretchLastSection(True)

                table_view.setSortingEnabled(True)
                table_view.setAlternatingRowColors(True)

                tab_layout.addWidget(table_view)
                self.results_tabs.addTab(tab, f"{filename} ({df_matches.shape[0]})")
                print(f"Added tab for {filename}. Tab count: {self.results_tabs.count()}") 

            except Exception as e:
                print(f"An error occurred while processing results for file {filename}: {e}")
                traceback.print_exc()
                self.status_bar.showMessage(f"Ошибка при обработке {filename}. Проверьте консоль.")

        print("Finished display_search_results")


    def close_tab(self, index):
        """Слот для закрытия вкладки."""
        # Удаление proxy модели происходит в отдельном слоте remove_proxy_model_for_tab
        self.results_tabs.removeTab(index)




    def update_status_bar(self, message):
        self.status_bar.showMessage(message)


    def show_error_message(self, message):
        QMessageBox.critical(self, "Ошибка", message)
        self.status_bar.showMessage("Ошибка. Проверьте сообщения.")
        self.set_buttons_enabled(True)


    def set_buttons_enabled(self, enabled: bool):
        self.folder_select_button.setEnabled(enabled)
        self.search_button.setEnabled(enabled)
        self.search_input.setEnabled(enabled)
        self.normal_search_radio.setEnabled(enabled)
        self.strict_search_radio.setEnabled(enabled)




    def on_worker_finished(self, *args):
        print("on_worker_finished called")
        current_message = self.status_bar.currentMessage()
        if "Ошибка" not in current_message and "найдено" not in current_message:
            self.status_bar.showMessage("Операция завершена.")
        elif "Ошибка" in current_message:
            pass


        self.set_buttons_enabled(True)


        if self.worker_thread:
            print("Quitting and waiting for worker thread...")
            if self.worker_thread.isRunning():
                self.worker_thread.quit()
                if not self.worker_thread.wait(1000):
                     print("Worker thread did not quit gracefully, terminating...")
                     self.worker_thread.terminate()
                     self.worker_thread.wait(1000)
            self.worker_thread = None
            print("Worker thread finished.")




    def closeEvent(self, event):
        """Обработка события закрытия окна для корректной остановки потока."""
        if self.worker_thread and self.worker_thread.isRunning():
            reply = QMessageBox.question(self, 'Завершение работы',
                                         "Идет фоновая операция.\nВы уверены, что хотите выйти?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                         QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                self.status_bar.showMessage("Остановка фоновой задачи...")
                self.worker_thread.requestInterruption() # requestInterruption is preferred over terminate
                if not self.worker_thread.wait(3000):
                    self.status_bar.showMessage("Принудительное завершение фоновой задачи...")
                    self.worker_thread.terminate()
                    self.worker_thread.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    try:
        app.setStyle("Fusion")
    except Exception as e:
        print(f"Не удалось применить стиль Fusion: {e}")


    STYLE_SHEET = """
/* --- Глобальные настройки --- */
QMainWindow {
    background-color: #F4F6F8; /* Очень светло-серый/голубоватый фон */ /* */
}

QWidget {
    font-family: "Segoe UI", Arial, sans-serif; /* */
    font-size: 10pt; /* */
    color: #2C3E50; /* Темно-серо-синий для основного текста */ /* */
}

/* --- Заголовки и метки --- */
QLabel {
    color: #2C3E50; /* */ /* */
    padding: 2px; /* */
}
QLabel#selected_folder_label { 
    font-weight: bold;
    color: #34495E; /* Более темный серо-синий */ /* */
}

/* --- Кнопки --- */
QPushButton {
    background-color: #5DADE2; /* Спокойный синий */ /* */
    color: white;
    border: 1px solid #3498DB; /* Чуть ярче синий для рамки */
    padding: 7px 14px; /* Немного увеличенные отступы */
    border-radius: 5px; /* Более скругленные углы */
    min-height: 22px;  /* */
}
QPushButton:hover {
    background-color: #85C1E9; /* Светлее синий при наведении */ /* */
    border: 1px solid #5DADE2;
}
QPushButton:pressed {
    background-color: #3498DB; /* Ярче синий при нажатии */ /* */
}
QPushButton:disabled {
    background-color: #BDC3C7; /* Светло-серый для неактивных */ /* */
    color: #7F8C8D; /* Темно-серый текст для неактивных */
    border: 1px solid #AABBC3; /* */
}

/* --- Поля ввода --- */
QLineEdit {
    background-color: white;
    color: #2C3E50;
    border: 1px solid #AABBC3; /* Серая рамка */ /* */
    padding: 6px; /* Увеличен padding */
    border-radius: 5px;
    min-height: 22px;  /* */
}
QLineEdit:focus {
    border: 2px solid #5DADE2; /* Синяя рамка при фокусе */ /* */
}
QLineEdit:read-only {
    background-color: #ECF0F1; /* Очень светло-серый */ /* */
}

/* --- Вкладки --- */
QTabWidget::pane {
    border: 1px solid #AABBC3; /* Серая рамка */
    border-top: none;  /* */
    background-color: #FFFFFF; /* Белый фон для содержимого вкладок */ /* */
}
QTabBar::tab {
    background-color: #E4E7EB; /* Светло-серый для неактивных вкладок */ /* */
    color: #34495E; /* Темно-серо-синий текст */
    border: 1px solid #D0D3D4; /* Светло-серая рамка */
    border-bottom: none;
    padding: 9px 18px; /* Увеличены отступы */ /* */
    border-top-left-radius: 5px;
    border-top-right-radius: 5px;
    margin-right: 3px; 
}
QTabBar::tab:selected {
    background-color: #5DADE2; /* Спокойный синий для активной вкладки */ /* */
    color: white;
    border: 1px solid #3498DB;
    border-bottom: 1px solid #5DADE2;  /* */
}
QTabBar::tab:hover:!selected {
    background-color: #F0F3F4; /* Еще светлее серый при наведении */ /* */
}
QTabWidget::tab-bar {
    alignment: left; /* */ /* */
}
QTabBar::close-button {
    /* image: url(path_to_your_close_icon.png); */ 
    subcontrol-position: right; /* */
}
QTabBar::close-button:hover {
    background-color: #E74C3C; /* Красный при наведении на кнопку закрытия (можно сделать менее ярким) */
    border-radius: 2px;
}

/* --- Таблицы --- */
QTableView {
    border: 1px solid #AABBC3; /* Серая рамка */ /* */
    gridline-color: #000000; /* Черные линии сетки */ /* */
    background-color: white;
    selection-background-color: #AED6F1; /* Светло-голубой для выделения */ /* */
    selection-color: #2C3E50; /* Темный текст на выделении */
    alternate-background-color: #F8F9F9; /* Очень светлый для чередующихся строк */ /* */
}
QHeaderView::section {
    background-color: #738FAB; /* Серо-синий для заголовков таблицы */ /* */
    color: white;
    padding: 7px; /* Увеличен padding */
    border: 1px solid #5D6D7E; /* Более темная рамка для разделения */ /* */
    font-weight: bold;
    text-transform: none; /* Убираем uppercase, если не нужен строгий вид */ /* */
}
QHeaderView {
    background-color: #F4F6F8;  /* */
}

/* --- Группы --- */
QGroupBox {
    background-color: #FDFEFE; /* Почти белый, но чуть отличный от основного фона */ /* */
    border: 1px solid #AABBC3; /* Серая рамка */ /* */
    border-radius: 6px;
    margin-top: 2ex; /* */ /* */
    font-weight: bold;  /* */
    padding-top: 10px; /* Добавим padding сверху для содержимого */
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left; /* */ /* */
    padding: 0 6px 0 6px; 
    left: 12px;  /* */
    color: #34495E; /* Темно-серо-синий */
    background-color: #F4F6F8; /* Фон под заголовком, чтобы он "вырезался" из рамки */ /* */
    border-radius: 3px; /* */
}

/* --- Радио-кнопки --- */
QRadioButton {
    color: #2C3E50;
    spacing: 6px;  /* */
    padding: 3px;
}
QRadioButton::indicator {
    width: 15px;  /* */
    height: 15px;
    border: 2px solid #738FAB; /* Серо-синий */
    border-radius: 8px; 
}
QRadioButton::indicator:unchecked {
    background-color: white; /* */
}
QRadioButton::indicator:unchecked:hover {
    border: 2px solid #5DADE2; /* Синий при наведении */
}
QRadioButton::indicator:checked {
    background-color: qradialgradient(
        cx: 0.5, cy: 0.5, fx: 0.5, fy: 0.5, radius: 0.4,
        stop: 0 #FFFFFF, stop: 0.3 #FFFFFF, 
        stop: 0.35 #738FAB, stop: 1 #738FAB  /* Серо-синий */
    ); /* */
    border: 2px solid #738FAB;
}
QRadioButton::indicator:disabled {
    border: 2px solid #BDC3C7; /* Светло-серый */
    background-color: #ECF0F1; /* Очень светло-серый */ /* */
}

/* --- Статус-бар --- */
QStatusBar {
    background-color: #738FAB; /* Серо-синий */
    color: white;
    font-weight: bold; /* */ /* */
}
QStatusBar::item {
    border: none; 
}

/* --- Сообщения QMessageBox --- */
QMessageBox {
    background-color: #F4F6F8; /* */ /* */
}
QMessageBox QLabel { 
    color: #2C3E50; /* */ /* */
}
/* QMessageBox QPushButton можно стилизовать отдельно, если потребуется */ /* */

/* --- Полосы прокрутки --- */
QScrollBar:horizontal {
    border: 1px solid #D0D3D4;
    background: #F4F6F8;
    height: 15px;  /* */
    margin: 0px 18px 0 18px; /* Уменьшены отступы */
}
QScrollBar::handle:horizontal {
    background: #AABBC3; /* Серый ползунок */ /* */
    min-width: 18px;
    border-radius: 7px;
}
QScrollBar::handle:horizontal:hover {
    background: #95A5A6; /* Темнее серый */ /* */
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    border: 1px solid #D0D3D4;
    background: #E4E7EB; /* Светло-серый */
    width: 18px;
    subcontrol-origin: margin; /* */ /* */
}
QScrollBar::add-line:horizontal {
    subcontrol-position: right;
}
QScrollBar::sub-line:horizontal {
    subcontrol-position: left; /* */ /* */
}
QScrollBar::left-arrow:horizontal, QScrollBar::right-arrow:horizontal {
    width: 0px; /* Скрыть стрелки */ /* */ /* */
    height: 0px;
    background: none;
}

QScrollBar:vertical {
    border: 1px solid #D0D3D4; /* */ /* */
    background: #F4F6F8;
    width: 15px;
    margin: 18px 0 18px 0;
}
QScrollBar::handle:vertical {
    background: #AABBC3; /* Серый ползунок */ /* */
    min-height: 18px;
    border-radius: 7px;
}
QScrollBar::handle:vertical:hover {
    background: #95A5A6; /* Темнее серый */ /* */
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    border: 1px solid #D0D3D4;
    background: #E4E7EB;
    height: 18px;
    subcontrol-origin: margin; /* */ /* */
}
QScrollBar::add-line:vertical {
    subcontrol-position: bottom;
}
QScrollBar::sub-line:vertical {
    subcontrol-position: top; /* */ /* */
}
QScrollBar::up-arrow:vertical, QScrollBar::down-arrow:vertical {
    width: 0px; /* Скрыть стрелки */ /* */ /* */
    height: 0px;
    background: none;
}
""" 
    
    app.setStyleSheet(STYLE_SHEET)


    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec())
