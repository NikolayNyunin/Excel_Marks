import sys
from time import time

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QPushButton, QLineEdit, QTextEdit, QLabel, QGridLayout
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont


def get_needed_mark(mark):
    if mark == 0:
        return 0
    elif mark < 2.5:
        return 2
    elif mark < 3.5:
        return 3
    elif mark < 4.5:
        return 4
    return 5


class ExcelMarksInterface(QWidget):
    def __init__(self):
        super().__init__()
        self.needed_file_description, self.select_file_button, self.selected_file_label, self.form_data_description,\
            self.form_input, self.start_analysing_button, self.output_console = [None] * 7
        self.init_ui()
        self.analyser = ExcelMarksAnalyser()
        self.filename = None
        self.show()

    def init_ui(self):
        self.setFixedSize(800, 600)
        self.setWindowTitle('Обработка оценок учеников')

        grid = QGridLayout()
        grid.setContentsMargins(40, 20, 40, 30)
        grid.setSpacing(20)
        self.setLayout(grid)

        self.needed_file_description = QLabel('Выберите файл с итоговыми оценками:', self)
        self.needed_file_description.setAlignment(Qt.AlignCenter)
        self.needed_file_description.setFont(QFont('Arial', 13))
        grid.addWidget(self.needed_file_description, 0, 0, 1, 3)

        self.select_file_button = QPushButton('Выбрать файл', self)
        self.select_file_button.setFont(QFont('Arial', 12))
        self.select_file_button.clicked.connect(self.select_file)
        self.select_file_button.setFixedSize(150, 40)
        grid.addWidget(self.select_file_button, 1, 0, alignment=Qt.AlignCenter)

        self.selected_file_label = QLabel('Файл не выбран.', self)
        self.selected_file_label.setFont(QFont('Arial', 12))
        grid.addWidget(self.selected_file_label, 1, 1, 1, 2)

        self.form_data_description = QLabel('Введите полное название класса (разделяя номер и букву дефисом):', self)
        self.form_data_description.setAlignment(Qt.AlignCenter)
        self.form_data_description.setFont(QFont('Arial', 13))
        self.form_data_description.setFixedWidth(570)
        grid.addWidget(self.form_data_description, 2, 0, 1, 2)

        self.form_input = QLineEdit('', self)
        self.form_input.setAlignment(Qt.AlignCenter)
        self.form_input.setFont(QFont('Arial', 14))
        self.form_input.setMaximumWidth(150)
        grid.addWidget(self.form_input, 2, 2, alignment=Qt.AlignRight)

        self.start_analysing_button = QPushButton('Начать', self)
        self.start_analysing_button.setFont(QFont('Arial', 13))
        self.start_analysing_button.clicked.connect(self.analyse)
        self.start_analysing_button.setAutoDefault(True)
        self.start_analysing_button.setFixedSize(200, 50)
        grid.addWidget(self.start_analysing_button, 4, 0, 1, 3, alignment=Qt.AlignCenter)

        self.output_console = QTextEdit('', self)
        self.output_console.setFont(QFont('Arial', 12))
        self.output_console.setReadOnly(True)
        self.output_console.setMaximumHeight(300)
        grid.addWidget(self.output_console, 5, 0, 1, 3)

    def select_file(self):  # метод для отображения окна выбора файла
        filename = QFileDialog.getOpenFileName(self, 'Выбор файла для обработки')[0]
        if filename == '':
            self.selected_file_label.setText('Файл не выбран.')
        else:
            self.filename = filename
            self.selected_file_label.setText('Выбранный файл: "{}".'.format(filename.split('/')[-1]))

    def analyse(self):  # метод для обработки файла
        if self.filename in (None, ''):  # проверка на то, что файл не выбран
            self.output_console.append('Ошибка: Файл не выбран.\n')
            return
        if not self.filename.endswith('.xlsx'):  # проверка на расширение файла
            self.output_console.append('Ошибка: Неподдерживаемое расширение файла: "{}".\n'.
                                       format(self.filename.split('.')[-1]))
            return

        form = self.form_input.text()
        if form == '':  # проверка на то, что класс не указан
            self.output_console.append('Ошибка: Класс не указан.\n')
            return
        if '-' not in form:  # проверка на отсутствие дефиса в названии класса
            self.output_console.append('Ошибка: Неправильный формат названия класса.\n')
            return

        try:
            start_time = time()

            new_filename = 'result_{}'.format(form)

            self.analyser.analyse_file(self.filename, form)
            self.analyser.create_resulting_file(new_filename)

            self.output_console.append('Успешно обработано: "{}" ({}).'.format(self.filename, form))
            self.output_console.append('Файл "{}" успешно создан.'.format(new_filename))
            self.output_console.append('Длительность выполнения: {} сек.\n'.format(str(round(time() - start_time, 2))))

        except Exception as e:
            self.output_console.append('Ошибка: {}.\n'.format(e))


class ExcelMarksAnalyser:
    def __init__(self):
        self.filename = None
        self.students = {}
        self.THIN = Side(border_style='thin', color='000000')
        self.THICK = Side(border_style='thick', color='000000')
        self.DOUBLE = Side(border_style='double', color='000000')

    def get_average_marks(self, path, filenames):  # метод для получения средних баллов из первого файла
        for file_num in range(len(filenames)):  # пробегаемся по файлам за каждый триместр
            workbook = load_workbook('{}{}'.format(path, filenames[file_num]), read_only=True)
            sheet = workbook.active

            subjects = list(map(lambda s: s.value, sheet[6][1:]))

            row_num = 7
            while True:  # пробегаемся по всем рядам таблицы с учениками
                if row_num > sheet.max_row:
                    break
                row = list(map(lambda c: c.value, sheet[row_num]))
                student = row[0]
                if student in ('', None):  # проверка на пустоту ячейки
                    break
                if student not in self.students.keys():
                    self.students[student] = {}

                for mark_index in range(len(row[1:])):  # пробегаемся по всем оценкам данного ученика
                    mark = float(row[1:][mark_index])
                    subject = subjects[mark_index]

                    if subject not in self.students[student].keys():
                        self.students[student][subject] = [[None] * 2, [None] * 2, [None] * 2]

                    self.students[student][subject][file_num][0] = mark

                row_num += 1

    def get_final_marks(self):  # метод для получения треместровых оценок
        workbook = load_workbook(self.filename, read_only=True)

        for sheet in workbook:  # пробегаемся по всем листам

            # получаем название предмета на данном листе
            subject = sheet['U41'].value.split(', ')[1]

            index = 1
            while sheet.cell(row=index, column=1).value:  # пробегаемся по всем таблицам

                # пробегаемся по всем строкам в этой таблице
                for row_num in range(index + 2, index + 50):

                    row = sheet[row_num][1:19]
                    student = ' '.join(row[0].value.split()[:2])  # получаем имя ученика
                    if not student:
                        break

                    for cell in row[1:]:  # пробегаемся по всем оценкам данного ученика в этой таблице
                        mark = cell.value
                        if mark:
                            if mark.isdigit() and cell.font.name == 'Arial Black':  # проверяем, триместровая ли оценка
                                for trimester in range(3):
                                    if self.students[student][subject][trimester][1] is None:
                                        if self.students[student][subject][trimester][0]:
                                            self.students[student][subject][trimester][1] = int(mark)
                                            break

                index += 50

    def analyse_file(self, filename, form):  # метод для обработки данных файлов
        self.filename = filename
        if len(filename.split('/')) > 1:
            path = '/'.join(filename.split('/')[:-1]) + '/'
        else:
            path = ''

        filenames = ['Отчёт по средним баллам 7-А класс. Iтр.xlsx',
                     'Отчёт по средним баллам 7-А класс. IIтр.xlsx',
                     'Отчёт по средним баллам 7-А класс. IIIтр.xlsx']

        self.get_average_marks(path, filenames)
        self.get_final_marks()

    def create_resulting_file(self, filename):  # метод для создания результирующего файла
        workbook = Workbook()
        sheet = workbook.active

        sheet['A1'] = 'Номер'
        sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
        sheet.merge_cells('A1:A3')
        sheet['A1'].border, sheet['A2'].border, sheet['A3'].border = \
            Border(right=self.THIN), Border(right=self.THIN), Border(right=self.THIN, bottom=self.THIN)
        sheet.column_dimensions['A'].width = 8

        sheet['B1'] = 'Имя ученика'
        sheet['B1'].alignment = Alignment(horizontal='center', vertical='center')
        sheet.merge_cells('B1:B3')
        sheet['B1'].border, sheet['B2'].border, sheet['B3'].border = \
            Border(right=self.THICK), Border(right=self.THICK), Border(right=self.THICK, bottom=self.THIN)
        sheet.column_dimensions['B'].width = 45

        student_index = 0  # порядковый номер ученика
        for student in sorted(self.students.keys()):  # пробегаемся по всем ученикам
            sheet.cell(row=student_index + 4, column=1).alignment = Alignment(horizontal='left')

            if student_index == len(self.students) - 1:
                sheet.cell(row=student_index + 4, column=2, value=student).border = Border(bottom=self.THIN,
                                                                                           right=self.THICK)
                sheet.cell(row=student_index + 4, column=1, value=student_index + 1).border = Border(bottom=self.THIN,
                                                                                                     right=self.THIN)
            else:
                sheet.cell(row=student_index + 4, column=1, value=student_index + 1).border = Border(right=self.THIN)
                sheet.cell(row=student_index + 4, column=2, value=student).border = Border(right=self.THICK)

            subject_index = 0
            for subject in self.students[student].keys():  # пробегаемся по всем предметам
                # возможно, стоит добавить сортировку

                subject_column = 3 + subject_index * 9
                if student_index == 0:  # если это список предметов первого ученика, заполняем шапку
                    sheet.cell(row=1, column=subject_column).value = subject
                    sheet.cell(row=1, column=subject_column).alignment = Alignment(horizontal='center')
                    sheet.merge_cells(start_row=1, start_column=subject_column,
                                      end_row=1, end_column=subject_column + 8)
                    sheet.cell(row=1, column=subject_column + 8).border = Border(right=self.THICK)

                for trimester in range(0, 3):  # пробегаемся по триместрам

                    column = subject_column + trimester * 3

                    if student_index == 0:  # если это список предметов первого ученика, заполняем шапку
                        sheet.cell(row=2, column=column).value = '{}-й триместр'.format(trimester + 1)
                        sheet.cell(row=2, column=column).alignment = Alignment(horizontal='center')
                        sheet.merge_cells(start_row=2, start_column=column,
                                          end_row=2, end_column=column + 2)
                        sheet.cell(row=2, column=column).border = Border(top=self.THIN, bottom=self.THIN)
                        sheet.cell(row=2, column=column + 1).border = Border(top=self.THIN, bottom=self.THIN)

                        sheet.cell(row=3, column=column).value = 'ср. б.'
                        sheet.cell(row=3, column=column + 1).value = 'рек.'
                        sheet.cell(row=3, column=column + 2).value = 'фактич.'

                        sheet.cell(row=3, column=column).alignment = Alignment(horizontal='center')
                        sheet.cell(row=3, column=column + 1).alignment = Alignment(horizontal='center')
                        sheet.cell(row=3, column=column + 2).alignment = Alignment(horizontal='center')

                        sheet.cell(row=3, column=column).border = Border(bottom=self.THIN, right=self.THIN)
                        sheet.cell(row=3, column=column + 1).border = Border(bottom=self.THIN, right=self.THIN)

                        if trimester == 2:
                            sheet.cell(row=2, column=column + 2).border = Border(bottom=self.THIN,
                                                                                 top=self.THIN, right=self.THICK)
                            sheet.cell(row=3, column=column + 2).border = Border(bottom=self.THIN, right=self.THICK)
                        else:
                            sheet.cell(row=2, column=column + 2).border = Border(bottom=self.THIN,
                                                                                 top=self.THIN, right=self.DOUBLE)
                            sheet.cell(row=3, column=column + 2).border = Border(bottom=self.THIN, right=self.DOUBLE)

                    marks = [0 if mark is None else mark for mark in self.students[student][subject][trimester]]

                    sheet.cell(row=student_index + 4, column=column).value = marks[0]
                    sheet.cell(row=student_index + 4, column=column + 2).value = marks[1]

                    sheet.cell(row=student_index + 4, column=column).alignment = Alignment(horizontal='center')
                    sheet.cell(row=student_index + 4, column=column + 2).alignment = Alignment(horizontal='center')

                    recommended = get_needed_mark(marks[0])
                    sheet.cell(row=student_index + 4, column=column + 1).value = recommended
                    sheet.cell(row=student_index + 4, column=column + 1).alignment = Alignment(horizontal='center')

                    if recommended != marks[1]:
                        sheet.cell(row=student_index + 4, column=column).font = Font(b=True)
                        sheet.cell(row=student_index + 4, column=column + 1).font = Font(b=True)
                        sheet.cell(row=student_index + 4, column=column + 2).font = Font(b=True)
                        sheet.cell(row=student_index + 4, column=column).fill = PatternFill(
                            start_color='FF4040',
                            end_color='FF4040',
                            fill_type='solid')
                        sheet.cell(row=student_index + 4, column=column + 1).fill = PatternFill(
                            start_color='FF4040',
                            end_color='FF4040',
                            fill_type='solid')
                        sheet.cell(row=student_index + 4, column=column + 2).fill = PatternFill(
                            start_color='FF4040',
                            end_color='FF4040',
                            fill_type='solid')

                    if student_index == len(self.students) - 1:  # если это последний ученик в списке,
                        # рисуем нижнюю границу
                        sheet.cell(row=student_index + 4, column=column).border = Border(bottom=self.THIN,
                                                                                         right=self.THIN)
                        sheet.cell(row=student_index + 4, column=column + 1).border = Border(bottom=self.THIN,
                                                                                             right=self.THIN)

                        if trimester == 2:
                            sheet.cell(row=student_index + 4, column=column + 2).border = Border(bottom=self.THIN,
                                                                                                 right=self.THICK)
                        else:
                            sheet.cell(row=student_index + 4, column=column + 2).border = Border(bottom=self.THIN,
                                                                                                 right=self.DOUBLE)
                    else:
                        sheet.cell(row=student_index + 4, column=column).border = Border(right=self.THIN)
                        sheet.cell(row=student_index + 4, column=column + 1).border = Border(right=self.THIN)

                        if trimester == 2:
                            sheet.cell(row=student_index + 4, column=column + 2).border = Border(right=self.THICK)
                        else:
                            sheet.cell(row=student_index + 4, column=column + 2).border = Border(right=self.DOUBLE)

                subject_index += 1

            student_index += 1

        workbook.save(filename)


def main():
    try:
        app = QApplication(sys.argv)
        gui = ExcelMarksInterface()
        sys.exit(app.exec())
    except Exception as e:
        print('Ошибка:', e)


if __name__ == '__main__':
    main()
