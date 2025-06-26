import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QComboBox, QLineEdit, QPushButton, QTreeWidget,
                             QTreeWidgetItem, QMessageBox, QTextEdit, QDialog, QFrame, QMenu)
from PyQt6.QtCore import Qt
import webbrowser
import sheets_manager
from datetime import datetime
from collections import Counter

memo_ts = """
Памятка для оформления ТЗ.

Если хотите сделать заголовки для подзадач, укажите их в '@@' с обеих сторон, тогда разработчик сможет воспользоваться функцией для быстрой обрабокти и работы с ТЗ.

Если хотите выделить что-то отдельным цветом (если считаете деталь очень важной), укажите этот текст в знаках '!*', тогда у разработчика текст будет выделен другим цветом."""


class StartWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.welcome_label = QLabel("Добро пожаловать!")
        self.welcome_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.welcome_label)

        self.role_label = QLabel("Выберите роль:")
        layout.addWidget(self.role_label)

        self.role_combobox = QComboBox()
        self.role_combobox.addItems(["Разработчик", "Менеджер"])
        layout.addWidget(self.role_combobox)

        self.sheets_label = QLabel("Введите sheetsId:")
        layout.addWidget(self.sheets_label)

        self.sheets_input = QLineEdit()
        layout.addWidget(self.sheets_input)

        self.confirm_button = QPushButton("Подтвердить")
        self.confirm_button.clicked.connect(self.save_data)
        layout.addWidget(self.confirm_button)

        self.setLayout(layout)

    def save_data(self):
        role = self.role_combobox.currentText()
        sheet_id = self.sheets_input.text()

        with open("settings.txt", "w") as f:
            f.write(f"role: {role},\nspreadsheetId: {sheet_id}")

        self.parent.initialize_app()


class InfoWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setWindowTitle("Информация")
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        profile = self.parent.sheets_api.take_settings()

        self.title_label = QLabel("Информация об аккаунте")
        layout.addWidget(self.title_label)

        self.sheet_id_label = QLabel(f"spreadsheetId: {profile[-1][-1]}")
        layout.addWidget(self.sheet_id_label)

        self.role_label = QLabel(f"Роль: {profile[0][-1]}")
        layout.addWidget(self.role_label)

        self.sheet_button = QPushButton("Google sheet")
        self.sheet_button.clicked.connect(self.open_sheet)
        layout.addWidget(self.sheet_button)

        self.setLayout(layout)

    def open_sheet(self):
        profile = self.parent.sheets_api.take_settings()
        webbrowser.open(f"https://docs.google.com/spreadsheets/d/{profile[-1][-1]}", new=2, autoraise=True)


class EditRowWindow(QDialog):
    def __init__(self, row_data, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.row_data = row_data
        self.setWindowTitle("Строка")
        self.initUI()

    def initUI(self):
        main_layout = QVBoxLayout()

        # Project frame
        project_frame = QFrame()
        project_layout = QVBoxLayout(project_frame)

        # Table
        self.table = QTreeWidget()
        self.table.setColumnCount(6)
        self.table.setHeaderLabels(["Дата", "Название", "Сроки", "Сумма", "Статус", "Срочность"])
        item = QTreeWidgetItem(self.row_data[:6])
        self.table.addTopLevelItem(item)
        project_layout.addWidget(self.table)

        # Input fields
        fields_frame = QFrame()
        fields_layout = QHBoxLayout(fields_frame)

        self.date_edit = QLineEdit(self.row_data[0])
        self.name_edit = QLineEdit(self.row_data[1])
        self.range_edit = QLineEdit(self.row_data[2])
        self.sum_edit = QLineEdit(self.row_data[3])
        self.status_edit = QLineEdit(self.row_data[4])
        self.speed_edit = QLineEdit(self.row_data[5])

        fields_layout.addWidget(QLabel("Дата"))
        fields_layout.addWidget(self.date_edit)
        fields_layout.addWidget(QLabel("Название"))
        fields_layout.addWidget(self.name_edit)
        fields_layout.addWidget(QLabel("Сроки"))
        fields_layout.addWidget(self.range_edit)
        fields_layout.addWidget(QLabel("Сумма"))
        fields_layout.addWidget(self.sum_edit)
        fields_layout.addWidget(QLabel("Статус"))
        fields_layout.addWidget(self.status_edit)
        fields_layout.addWidget(QLabel("Срочность"))
        fields_layout.addWidget(self.speed_edit)

        project_layout.addWidget(fields_frame)
        main_layout.addWidget(project_frame)

        # Technical specification frame
        ts_frame = QFrame()
        ts_layout = QVBoxLayout(ts_frame)
        ts_layout.addWidget(QLabel("Техническое задание"))

        self.ts_text = QTextEdit()
        self.ts_text.setPlainText(self.row_data[6] if len(self.row_data) > 6 else "")
        ts_layout.addWidget(self.ts_text)

        ts_layout.addWidget(QLabel(memo_ts))
        main_layout.addWidget(ts_frame)

        # Buttons frame
        buttons_frame = QFrame()
        buttons_layout = QHBoxLayout(buttons_frame)

        self.save_button = QPushButton("Сохранить")
        self.save_button.clicked.connect(self.save_row)
        buttons_layout.addWidget(self.save_button)

        #self.delete_button = QPushButton("Удалить")
        #buttons_layout.addWidget(self.delete_button)

        main_layout.addWidget(buttons_frame)

        self.setLayout(main_layout)

    def save_row(self):
        new_values = [
            self.date_edit.text(),
            self.name_edit.text(),
            self.range_edit.text(),
            self.sum_edit.text(),
            self.status_edit.text(),
            self.speed_edit.text(),
            self.ts_text.toPlainText()
        ]

        # Here you would implement the logic to save the edited row
        # Similar to the original save_edited_row method
        self.parent.save_edited_row(new_values)
        self.accept()


class AnalyzeWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.sheets_api = parent.sheets_api if parent else None
        self.setWindowTitle("Анализ данных")
        self.setMinimumSize(800, 600)
        self.initUI()
        self.load_data()

    def initUI(self):
        layout = QVBoxLayout()

        # Заголовок
        title = QLabel("Анализ данных")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)

        # Фильтры
        filter_layout = QHBoxLayout()

        # Фильтр по годам
        self.year_filter = QComboBox()
        self.year_filter.addItem("Все годы")
        filter_layout.addWidget(QLabel("Год:"))
        filter_layout.addWidget(self.year_filter)

        # Фильтр по месяцам
        self.month_filter = QComboBox()
        self.month_filter.addItem("Все месяцы")
        self.month_filter.addItems([
            "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
            "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
        ])
        filter_layout.addWidget(QLabel("Месяц:"))
        filter_layout.addWidget(self.month_filter)

        # Кнопка применения фильтров
        self.apply_filters_btn = QPushButton("Применить фильтры")
        self.apply_filters_btn.clicked.connect(self.apply_filters)
        filter_layout.addWidget(self.apply_filters_btn)

        layout.addLayout(filter_layout)

        # Кнопка обновления данных
        self.refresh_btn = QPushButton("Обновить данные")
        self.refresh_btn.clicked.connect(self.load_data)
        layout.addWidget(self.refresh_btn)

        # Выбор типа анализа
        self.analysis_type = QComboBox()
        self.analysis_type.addItems([
            "Статистика по статусам",
            "Финансовая сводка",
            "Анализ срочности",
            "Распределение по срокам",
            "Распределение по месяцам"
        ])
        layout.addWidget(self.analysis_type)

        # Кнопка выполнения анализа
        self.run_analysis_btn = QPushButton("Выполнить анализ")
        self.run_analysis_btn.clicked.connect(self.run_analysis)
        layout.addWidget(self.run_analysis_btn)

        # Область для вывода результатов
        self.results_area = QTextEdit()
        self.results_area.setReadOnly(True)
        self.results_area.setStyleSheet("font-family: monospace;")
        layout.addWidget(self.results_area)

        # Статус загрузки данных
        self.status_label = QLabel("Данные не загружены")
        layout.addWidget(self.status_label)

        self.setLayout(layout)

    def load_data(self):
        try:
            if not self.sheets_api:
                raise Exception("Не подключен API для работы с Google Sheets")

            profile = self.sheets_api.take_settings()
            if not profile or len(profile) < 2:
                raise Exception("Не удалось получить настройки профиля")

            spreadsheet_id = profile[-1][-1]
            raw_data = self.sheets_api.take_datas(spreadsheet_id, "Основной расчёт",
                                                  "A2:F100")  # Убедитесь, что берем правильные столбцы

            # Обработка и фильтрация данных
            self.all_data = []
            years = set()

            for row in raw_data:
                if len(row) >= 1 and row[0]:  # Проверяем наличие даты
                    try:
                        # Парсим дату из формата дд.мм.гггг
                        date = datetime.strptime(row[0], "%d.%m.%Y")
                        year = date.year
                        month = date.month
                        years.add(year)

                        # Парсим сроки выполнения (берем только конечную дату)
                        timeline = row[2] if len(row) > 2 else ""
                        end_date = None
                        if timeline and "-" in timeline:
                            try:
                                end_date_str = timeline.split("-")[1].strip()
                                end_date = datetime.strptime(end_date_str, "%d.%m.%Y")
                            except:
                                end_date = None

                        # Рассчитываем длительность в днях
                        duration_days = (end_date - date).days + 1 if end_date else 0

                        # Сохраняем обработанные данные
                        processed_row = {
                            "date": date,
                            "year": year,
                            "month": month,
                            "name": row[1] if len(row) > 1 else "",
                            "timeline": timeline,
                            "duration_days": duration_days,
                            "amount": float(row[3]) if len(row) > 3 and str(row[3]).replace('.', '').isdigit() else 0,
                            "status": row[4] if len(row) > 4 else "",
                            "urgency": int(row[5]) if len(row) > 5 and str(row[5]).isdigit() else 0
                        }
                        self.all_data.append(processed_row)
                    except Exception as e:
                        print(f"Ошибка обработки строки {row}: {str(e)}")
                        continue

            # Обновляем фильтр годов
            self.year_filter.clear()
            self.year_filter.addItem("Все годы")
            for year in sorted(years):
                self.year_filter.addItem(str(year))

            self.filtered_data = self.all_data.copy()
            self.status_label.setText(f"Загружено {len(self.all_data)} строк")
            return True

        except Exception as e:
            self.status_label.setText(f"Ошибка загрузки: {str(e)}")
            self.results_area.setPlainText(f"Ошибка при загрузке данных:\n{str(e)}")
            return False

    def apply_filters(self):
        selected_year = self.year_filter.currentText()
        selected_month = self.month_filter.currentIndex()  # 0 = все месяцы

        self.filtered_data = []

        for item in self.all_data:
            year_match = (selected_year == "Все годы") or (str(item["year"]) == selected_year)
            month_match = (selected_month == 0) or (item["month"] == selected_month)

            if year_match and month_match:
                self.filtered_data.append(item)

        self.status_label.setText(f"Отфильтровано {len(self.filtered_data)} строк")

    def run_analysis(self):
        if not hasattr(self, 'filtered_data') or not self.filtered_data:
            self.results_area.setPlainText("Нет данных для анализа. Попробуйте обновить данные.")
            return

        analysis_type = self.analysis_type.currentText()
        results = ""

        try:
            if analysis_type == "Статистика по статусам":
                statuses = [item["status"] for item in self.filtered_data]
                counter = Counter(statuses)
                total = len(statuses)
                results = "Статистика по статусам:\n"
                for status, count in counter.most_common():
                    results += f"- {status}: {count} ({count / total:.1%})\n"

            elif analysis_type == "Финансовая сводка":
                amounts = [item["amount"] for item in self.filtered_data]
                if amounts:
                    results = "Финансовая сводка:\n"
                    results += f"- Общая сумма: {sum(amounts):,.2f} руб.\n"
                    results += f"- Средняя сумма: {sum(amounts) / len(amounts):,.2f} руб.\n"
                    results += f"- Максимальная сумма: {max(amounts):,.2f} руб.\n"
                    results += f"- Минимальная сумма: {min(amounts):,.2f} руб.\n"
                    results += f"- Кол-во проектов: {len(amounts)}\n"
                else:
                    results = "Нет данных для финансового анализа"

            elif analysis_type == "Анализ срочности":
                urgency = [item["urgency"] for item in self.filtered_data]
                counter = Counter(urgency)
                results = "Распределение по срочности:\n"
                for level, count in counter.most_common():
                    results += f"- {level}: {count}\n"


            elif analysis_type == "Распределение по срокам":
                durations = [item["duration_days"] for item in self.filtered_data if item["duration_days"] > 0]
                if durations:
                    results = "Анализ сроков выполнения:\n"
                    results += f"- Средний срок: {sum(durations) / len(durations):.1f} дней\n"
                    results += f"- Минимальный срок: {min(durations)} дней\n"
                    results += f"- Максимальный срок: {max(durations)} дней\n"

                    # Группировка по диапазонам
                    results += "\nГруппировка по срокам:\n"
                    ranges = {
                        "1 день": 0,
                        "2-3 дня": 0,
                        "4-7 дней": 0,
                        "1-2 недели": 0,
                        "более 2 недель": 0
                    }
                    for days in durations:
                        if days == 1:
                            ranges["1 день"] += 1
                        elif 2 <= days <= 3:
                            ranges["2-3 дня"] += 1
                        elif 4 <= days <= 7:
                            ranges["4-7 дней"] += 1
                        elif 8 <= days <= 14:
                            ranges["1-2 недели"] += 1
                        else:
                            ranges["более 2 недель"] += 1
                    total = len(durations)
                    for r, count in ranges.items():
                        results += f"- {r}: {count} ({count / total:.1%})\n"
                else:
                    results = "Нет данных о сроках выполнения"

            elif analysis_type == "Распределение по месяцам":
                month_names = [
                    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
                    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
                ]

                months = [item["month"] for item in self.filtered_data]
                counter = Counter(months)

                results = "Распределение по месяцам:\n"
                for month_num in range(1, 13):
                    count = counter.get(month_num, 0)
                    results += f"- {month_names[month_num - 1]}: {count}\n"

            self.results_area.setPlainText(results)

        except Exception as e:
            self.results_area.setPlainText(f"Ошибка при выполнении анализа:\n{str(e)}")

    @classmethod
    def show_analysis(cls, parent):
        analyzer = cls(parent)
        analyzer.exec()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.sheets_api = sheets_manager.sheetAPI()

        try:
            with open("settings.txt", "r") as f:
                settings_file = f.read()
        except FileNotFoundError:
            settings_file = ""

        if settings_file.strip() == "":
            self.start_window = StartWindow(self)
            self.setCentralWidget(self.start_window)
        else:
            self.initialize_app()

        self.setWindowTitle("BuisnessData")
        self.setMinimumSize(700, 500)

    def initialize_app(self):
        profile = self.sheets_api.take_settings()
        if profile[0][-1] == "Менеджер":
            self.show_manager_interface()
        elif profile[0][-1] == "Разработчик":
            self.show_developer_interface()

    def show_manager_interface(self):
        central_widget = QWidget()
        layout = QVBoxLayout(central_widget)

        # Data table
        self.data_table = QTreeWidget()
        self.data_table.setColumnCount(6)
        self.data_table.setHeaderLabels(["Дата", "Название", "Сроки", "Сумма", "Статус", "Срочность"])
        layout.addWidget(self.data_table)

        # Add data frame
        add_frame = QFrame()
        add_layout = QHBoxLayout(add_frame)

        self.date_input = QLineEdit()
        self.name_input = QLineEdit()
        self.range_input = QLineEdit()
        self.sum_input = QLineEdit()
        self.status_input = QLineEdit()
        self.speed_input = QLineEdit()

        add_layout.addWidget(self.date_input)
        add_layout.addWidget(self.name_input)
        add_layout.addWidget(self.range_input)
        add_layout.addWidget(self.sum_input)
        add_layout.addWidget(self.status_input)
        add_layout.addWidget(self.speed_input)

        self.data_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.data_table.customContextMenuRequested.connect(self.show_context_menu)

        self.add_button = QPushButton("Добавить")
        self.add_button.clicked.connect(self.add_data)
        add_layout.addWidget(self.add_button)

        layout.addWidget(add_frame)

        # Info button
        self.info_button = QPushButton("Инфо")
        self.info_button.clicked.connect(self.show_info)
        layout.addWidget(self.info_button)

        self.setCentralWidget(central_widget)
        self.load_data()

    def show_developer_interface(self):
        # Similar to manager interface for now
        self.show_manager_interface()

    def show_context_menu(self, position):
        try:
            menu = QMenu()
            selected_item = self.data_table.itemAt(position)

            edit_action = menu.addAction("Редактировать строку")
            #delete_action = menu.addAction("Удалить строку")
            menu.addSeparator()
            analyze_action = menu.addAction("Анализировать данные")

            if not selected_item:
                edit_action.setEnabled(False)
                #analyze_action.setEnabled(False)
                #delete_action.setEnabled(False)

            action = menu.exec(self.data_table.viewport().mapToGlobal(position))

            if action == edit_action and selected_item:
                self.edit_row(selected_item)
            elif action == analyze_action and selected_item:
                self.analyze_selected_row(selected_item)  # Убедитесь, что этот метод реализован
            #elif action == delete_action and selected_item:
                #self.delete_selected_row(selected_item)  # Убедитесь, что этот метод реализован

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка в контекстном меню: {str(e)}")

    def load_data(self):
        profile = self.sheets_api.take_settings()
        data = self.sheets_api.take_datas(profile[-1][-1], "Основной расчёт", "A2:G100")

        self.data_table.clear()
        for row in data:
            item = QTreeWidgetItem(row[:6])  # Only show first 6 columns in main table
            self.data_table.addTopLevelItem(item)

    def add_data(self):
        profile = self.sheets_api.take_settings()
        values = [
            self.date_input.text(),
            self.name_input.text(),
            self.range_input.text(),
            self.sum_input.text(),
            self.status_input.text(),
            self.speed_input.text()
        ]

        # Calculate the row number (similar to original logic)
        row_num = 2 + self.data_table.topLevelItemCount()

        # Insert data
        self.sheets_api.insert_datas(profile[-1][-1], "Основной расчёт", f"A{row_num}:F{row_num}", [values])

        QMessageBox.information(self, "Готово", "Строка добавлена!")
        self.load_data()

    def analyze_selected_row(self, item):
        try:
            if not item:
                QMessageBox.warning(self, "Ошибка", "Не выбрана строка для редактирования")
                return

            # Получаем данные
            profile = self.sheets_api.take_settings()
            all_data = self.sheets_api.take_datas(profile[-1][-1], "Основной расчёт", "A2:G100")

            # Проверяем индекс строки
            row_index = self.data_table.indexOfTopLevelItem(item)
            if row_index < 0 or row_index >= len(all_data):
                raise IndexError("Неверный индекс строки")

            row_data = all_data[row_index]

            # Создаем и показываем диалог редактирования ПЕРЕДАВАЯ ДАННЫЕ
            edit_dialog = AnalyzeWindow(self)  # Важно передать row_data!
            if edit_dialog.exec() == QDialog.DialogCode.Accepted:
                self.load_data()  # Обновляем данные после редактирования

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть строку для редактирования: {str(e)}")

    def edit_row(self, item):
        try:
            if not item:
                QMessageBox.warning(self, "Ошибка", "Не выбрана строка для редактирования")
                return

            # Получаем данные
            profile = self.sheets_api.take_settings()
            all_data = self.sheets_api.take_datas(profile[-1][-1], "Основной расчёт", "A2:G100")

            # Проверяем индекс строки
            row_index = self.data_table.indexOfTopLevelItem(item)
            if row_index < 0 or row_index >= len(all_data):
                raise IndexError("Неверный индекс строки")

            row_data = all_data[row_index]

            # Создаем и показываем диалог редактирования ПЕРЕДАВАЯ ДАННЫЕ
            edit_dialog = EditRowWindow(row_data, self)  # Важно передать row_data!
            if edit_dialog.exec() == QDialog.DialogCode.Accepted:
                self.load_data()  # Обновляем данные после редактирования

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть строку для редактирования: {str(e)}")

    def save_edited_row(self, new_values):
        profile = self.sheets_api.take_settings()
        all_data = self.sheets_api.take_datas(profile[-1][-1], "Основной расчёт", "A2:G100")

        # Find the row to update (similar to original logic)
        # This is a simplified version - you might need to adjust it
        for i, row in enumerate(all_data):
            if row[:6] == new_values[:6]:  # Compare first 6 columns
                # Update the data
                self.sheets_api.insert_datas(profile[-1][-1], "Основной расчёт", f"A{2 + i}:G{2 + i}", [new_values])
                break

        self.load_data()

    def show_info(self):
        info_window = InfoWindow(self)
        info_window.exec()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())