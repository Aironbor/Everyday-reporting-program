import os
import re
import sys
from PyQt5.QtWidgets import QDialog
from PyQt5.QtCore import *
from datetime import datetime, date
from PyQt5 import QtCore, QtWidgets, QtGui
import xlsxwriter
from PyQt5 import uic
from openpyxl import load_workbook

from config import db


class ExcelWork:
    def __init__(self, t_progect, n_date, date_for_plan):
        super().__init__()
        self.type_prog = t_progect
        self.date_of_report = n_date
        self.data_of_mounth_plan = date_for_plan
        if self.type_prog == 'КМД':
            self.report_excel_kmd()

        elif self.type_prog == 'СПУ':
            self.report_excel_spu()

        elif self.type_prog == 'ПВХ':
            self.report_excel_pvh()
        print(self.data_of_mounth_plan,date_for_plan )

    def report_excel_kmd(self):
        #  ОТЧЕТ ДЛЯ КМД
        workbook = xlsxwriter.Workbook(f'Отчеты/Отчет по цеху Металлоконструкций.xlsx')
        # Форматы format()
        percent_format = workbook.add_format(
            {'border': 1, 'num_format': '0.00%', 'align': 'left', 'valign': 'vcenter'})
        percent_format_for_plan = workbook.add_format(
            { 'num_format': '0.00%', 'align': 'left', 'valign': 'vcenter'})
        name_format = workbook.add_format(
            {'border': 1, 'bold': True, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        name_format_main = workbook.add_format(
            {'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        date_format = workbook.add_format(
            {'border': 1, 'text_wrap': True, 'num_format': 'dd MMM yy', 'align': 'center', 'valign': 'vcenter'})
        date_format_main = workbook.add_format(
            {'text_wrap': True, 'num_format': 'dd MMM yy', 'align': 'center', 'valign': 'vcenter'})
        special_numb = workbook.add_format(
            {'border': 1, 'num_format': '#0', 'align': 'center', 'valign': 'vcenter'})
        float_numb_w_board = workbook.add_format(
            {'border': 1, 'num_format': '#0.00', 'align': 'center', 'valign': 'vcenter'})
        numb_w_border = workbook.add_format(
            {'num_format': '#0.00', 'align': 'center', 'valign': 'vcenter'})
        # Форматы для объединнеых ячеек
        name_merge_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#0',
        })
        name_merge_format_main = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#0',
            'bold': True
        })
        name_merge_format_spec = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#0',
            'bold': True
        })
        name_merge_format_spec_2 = workbook.add_format({
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': '#0',
            'bold': True
        })
        name_merge_format_2 = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#0',
            'fg_color': '#DDEBF7'
        })
        workbook.set_properties({
            'title': f'Производственный отчет',
            'subject': 'With document properties',
            'author': 'Ivan Metliaev',
            'manager': '',
            'company': 'Тентовые конструкции',
            'category': 'КМД',
            'keywords': 'КМД, Ангары, Металл',
            'created': datetime.today(),
            'comments': 'Created with Python and Ivan Metliaev program'})
        # Создаваемые листы
        worksheet_0 = workbook.add_worksheet(f'Сводный отчет')
        # Размер колонок
        worksheet_0.set_column(0, 0, 14)
        worksheet_0.set_column(1, 1, 20)
        worksheet_0.set_column(2, 2, 14)
        worksheet_0.set_column(3, 3, 17)
        worksheet_0.set_column(4, 4, 19)
        worksheet_0.set_column(5, 5, 13)
        worksheet_0.set_column(6, 6, 14)
        # Записи Ход изготовления ангара
        worksheet_0.merge_range(0, 0, 0, 4, f'Cводный отчет по цеху Металлоконструкций',
                                name_merge_format_main)
        worksheet_0.write("F1", "Текущая дата:", name_format_main)
        worksheet_0.write("G1", f"{self.date_of_report}", date_format_main )
        worksheet_0.write("A4", f"План месяц, т", name_format_main)
        kmd_plan = 0.0
        try:
            for kmd in db.get_kmd_plan_for_mounth(self.data_of_mounth_plan):
                kmd_plan = kmd
        except:
            pass
        worksheet_0.write("B4", kmd_plan,  numb_w_border)
        worksheet_0.write("A5", f"Выполнение плана", name_format_main)
        worksheet_0.write_formula("B5", f"Выполнение плана", name_format_main)
        row_name = ['№', 'Проект', 'По проекту, т', 'Изготовлено на текущ. момент, т', 'Производственная готовность']
        # Заголовки первой таблицы
        curnt_numb_row = 7
        num = 0
        worksheet_0.write_row(6, 0, row_name, name_format)
        curnt_special_row = 7
        for info in db.get_massa_progect():
            key = info[1]
            for progect_data in db.get_info_about_today_kmd_report(key, self.date_of_report):
                # №
                num += 1
                worksheet_0.write(curnt_special_row, 0, num, special_numb)
                # Проект
                worksheet_0.write(curnt_special_row, 1, progect_data[0], name_format)
                # По проекту
                worksheet_0.write(curnt_special_row, 2, progect_data[1] / 1000, float_numb_w_board)
                # Готовность
                worksheet_0.write(curnt_special_row, 3, progect_data[2] / 1000, float_numb_w_board)
                # Процент
                worksheet_0.write(curnt_special_row, 4, progect_data[3]/100, percent_format)
                curnt_special_row +=1
            curnt_numb_row += 1
        worksheet_0.write_formula(curnt_special_row, 3 , f'=SUM(D8:D{curnt_special_row})', float_numb_w_board)
        worksheet_0.write_formula("B5", f'=D{curnt_special_row + 1}/B4', percent_format_for_plan)
        worksheet_0.merge_range(curnt_special_row, 0, curnt_special_row, 2, f'Итого:',
                                name_merge_format_spec)
        worksheet_0.conditional_format(7, 4, curnt_special_row, 4, {'type': 'data_bar'})
        worksheet_0.ignore_errors()
        for progect_in_work in db.get_info_progect_kmd():
            worksheet_1 = workbook.add_worksheet(f''.join(progect_in_work))
            # Размер колонок
            worksheet_1.set_column(3, 1, 10)
            worksheet_1.set_column(3, 2, 14)
            worksheet_1.set_column(3, 3, 14)
            worksheet_1.set_column(3, 4, 14)
            worksheet_1.set_column(5, 5, 18)
            # Записи Ход изготовления ангара
            worksheet_1.merge_range(0, 0, 0, 5, f'Производственный отчет нарастающим итогом на отчетную дату', name_merge_format_spec)
            worksheet_1.write(1, 0, "Проект:", name_format)
            worksheet_1.merge_range(1, 1, 1, 5, f' ,'.join(progect_in_work), name_merge_format)
            worksheet_1.merge_range(2, 0, 2, 5, f'Конструкции металлические и деталировка', name_merge_format_2)
            row_name = ['№', 'Дата', 'Заготовка', 'Сварка', 'Покраска', 'Общая готовность ангара']
            # Заголовки первой таблицы
            worksheet_1.write_row(3, 0, row_name, name_format)
            curnt_numb_row = 4
            num = 0
            for info in db.get_info_ab_day_report(''.join(progect_in_work)):
                num += 1
                # №
                worksheet_1.write(curnt_numb_row, 0, num, special_numb)
                # Дата
                worksheet_1.write(curnt_numb_row, 1, info[2], date_format)
                # % Заготовки
                worksheet_1.write(curnt_numb_row, 2, info[3] / 100, percent_format)
                # % Сварки
                worksheet_1.write(curnt_numb_row, 3, info[4] / 100, percent_format)
                # % Покраски
                worksheet_1.write(curnt_numb_row, 4, info[5] / 100, percent_format)
                # % Общий проц готовности
                worksheet_1.write(curnt_numb_row, 5, info[6] / 100, percent_format)
                curnt_numb_row += 1

            worksheet_1.conditional_format(4, 2, curnt_numb_row, 5, {'type': 'data_bar'})
        workbook.close()
        os.startfile(f'Отчеты\Отчет по цеху Металлоконструкций.xlsx')

        # ОТЧЕТ ДЛЯ СПУ
    def report_excel_spu(self):
        workbook = xlsxwriter.Workbook(f'Отчеты/Отчет по участку изготовления СПУ.xlsx')
        # Форматы format()
        percent_format = workbook.add_format(
            {'border': 1, 'num_format': '0.00%', 'align': 'left', 'valign': 'vcenter'})
        percent_format_for_plan = workbook.add_format(
            {'num_format': '0.00%', 'align': 'left', 'valign': 'vcenter'})
        name_format = workbook.add_format(
            {'border': 1, 'bold': True, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        name_format_main = workbook.add_format(
            {'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        date_format = workbook.add_format(
            {'border': 1, 'text_wrap': True, 'num_format': 'dd MMM yy', 'align': 'center', 'valign': 'vcenter'})
        date_format_main = workbook.add_format(
            {'text_wrap': True, 'num_format': 'dd MMM yy', 'align': 'center', 'valign': 'vcenter'})
        special_numb = workbook.add_format(
            {'border': 1, 'num_format': '#0', 'align': 'center', 'valign': 'vcenter'})
        float_numb_w_board = workbook.add_format(
            {'border': 1, 'num_format': '#0.00', 'align': 'center', 'valign': 'vcenter'})
        numb_w_border = workbook.add_format(
            {'num_format': '#0.00', 'align': 'center', 'valign': 'vcenter'})
        # Форматы для объединнеых ячеек
        name_merge_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#0',
        })
        name_merge_format_spec = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#0',
            'bold': True
        })
        name_merge_format_2 = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#0',
            'fg_color': '#DDEBF7'
        })
        # Создаваемые листы
        worksheet_0 = workbook.add_worksheet(f'Сводный отчет')
        # Размер колонок
        worksheet_0.set_column(0, 0, 14)
        worksheet_0.set_column(1, 1, 20)
        worksheet_0.set_column(2, 2, 14)
        worksheet_0.set_column(3, 3, 17)
        worksheet_0.set_column(4, 4, 19)
        worksheet_0.set_column(5, 5, 13)
        worksheet_0.set_column(6, 6, 14)
        # Записи Ход изготовления ангара
        worksheet_0.merge_range(0, 0, 0, 4, f'Cводный отчет по участку производства утеплителя',
                                name_merge_format)
        worksheet_0.write("F1", "Текущая дата:", name_format_main)
        worksheet_0.write("G1", f"{self.date_of_report}", date_format_main )
        worksheet_0.write("A4", f"План месяц, м2", name_format_main)
        tent_plan = 0.0
        try:
            for tent in db.get_spu_plan_for_mounth(self.data_of_mounth_plan):
                tent_plan = tent
        except:
            pass
        worksheet_0.write("B4", tent_plan,  numb_w_border)
        worksheet_0.write("A5", f"Выполнение плана", name_format_main)
        worksheet_0.write_formula("B5", f"Выполнение плана", name_format_main)
        row_name = ['№', 'Проект', 'По проекту, м2', 'Изготовлено на текущ. момент, м2',
                    'Производственная готовность', 'Готовность к отгрузке' ]
        # Заголовки первой таблицы
        curnt_numb_row = 7
        num = 0
        worksheet_0.write_row(6, 0, row_name, name_format)
        curnt_special_row = 7
        for info in db.get_squer_spu_progect():
            key = info[1]
            for progect_data in db.get_info_about_today_spu_report(key, self.date_of_report):
                # №
                num += 1
                worksheet_0.write(curnt_numb_row, 0, num, special_numb)
                # Проект
                worksheet_0.write(curnt_numb_row, 1, progect_data[0], name_format)
                # По проекту
                worksheet_0.write(curnt_numb_row, 2, progect_data[1] / 1000, float_numb_w_board)
                # Готовность
                worksheet_0.write(curnt_numb_row, 3, progect_data[2] / 1000, float_numb_w_board)
                # Процент готовности
                worksheet_0.write(curnt_numb_row, 4, progect_data[3]/100, percent_format)
                # Процент отгрузки
                worksheet_0.write(curnt_numb_row, 4, progect_data[4] / 100, percent_format)
                curnt_special_row +=1

            curnt_numb_row += 1
        worksheet_0.write_formula(curnt_special_row, 3 , f'=SUM(D8:D{curnt_special_row})', float_numb_w_board)
        worksheet_0.write_formula("B5", f'=D{curnt_special_row + 1}/B4', percent_format_for_plan)
        worksheet_0.merge_range(curnt_special_row, 0, curnt_special_row, 2, f'Итого:',
                                name_merge_format_spec)
        worksheet_0.conditional_format(7, 4, curnt_special_row, 5, {'type': 'data_bar'})

        for progect_in_work in db.get_info_progect_spu():
            # Создаваемые листы
            worksheet_1 = workbook.add_worksheet(f''.join(progect_in_work))
            # Размер колонок
            worksheet_1.set_column(3, 1, 10)
            worksheet_1.set_column(3, 2, 15)
            worksheet_1.set_column(3, 3, 15)
            worksheet_1.set_column(3, 4, 15)
            worksheet_1.set_column(3, 5, 15)
            worksheet_1.set_column(3, 6, 15)
            worksheet_1.set_column(3, 7, 15)
            worksheet_1.set_column(3, 8, 15)
            worksheet_1.set_column(4, 9, 18)
            # Записи Ход изготовления ангара
            worksheet_1.merge_range(0, 0, 0, 9, f'Производственный отчет нарастающим итогом на отчетную дату', name_merge_format_spec)
            worksheet_1.write(1, 0, "Проект:", name_format)
            worksheet_1.merge_range(1, 1, 1, 9, f' ,'.join(progect_in_work), name_merge_format)
            worksheet_1.merge_range(2, 0, 2, 9, f'Утеплитель', name_merge_format_2)
            row_name = ['№', 'Дата', 'Раскрой полипропилена',
                        'Пришить крючки и петли',
                        'Сшивка полипропилена',
                        'Наклейка синтепона',
                        'Пробивка люверс',
                        'Сборка',
                        'Упаковано',
                        'Общая готовность утеплителя']
            # Заголовки первой таблицы
            worksheet_1.write_row(3, 0, row_name, name_format)
            curnt_numb_row = 4
            num = 0
            for info in db.get_info_ab_day_report_spu(f''.join(progect_in_work)):
                num += 1
                # №
                worksheet_1.write(curnt_numb_row, 0, num, special_numb)
                # Дата
                worksheet_1.write(curnt_numb_row, 1, info[2], date_format)
                # % Раскрой полипропилена
                worksheet_1.write(curnt_numb_row, 2, info[3] / 100, percent_format)
                # % Пришить крючки и петли
                if info[4] is not None:
                    worksheet_1.write(curnt_numb_row, 3, info[4] / 100, percent_format)
                else:
                    worksheet_1.write(curnt_numb_row, 3, '-', percent_format)
                # % сшивка полипропилена
                worksheet_1.write(curnt_numb_row, 4, info[5] / 100, percent_format)
                # % Наклейка
                worksheet_1.write(curnt_numb_row, 5, info[6] / 100, percent_format)
                # % Сборка
                if info[7] is not None:
                    worksheet_1.write(curnt_numb_row, 6, info[7] / 100, percent_format)
                else:
                    worksheet_1.write(curnt_numb_row, 6, '-', percent_format)
                # % Пробивка люверс
                worksheet_1.write(curnt_numb_row, 7, info[8] / 100, percent_format)
                # % Упаковка
                worksheet_1.write(curnt_numb_row, 8, info[9] / 100, percent_format)
                # % готовности утеплителя
                worksheet_1.write(curnt_numb_row, 9, info[10] / 100, percent_format)
                curnt_numb_row += 1

            worksheet_1.conditional_format(4, 2, curnt_numb_row, 7, {'type': 'data_bar'})

        workbook.close()
        os.startfile(f'Отчеты\Отчет по участку изготовления СПУ.xlsx')

        # ОТЧЕТ ДЛЯ ПВХ
    def report_excel_pvh(self):
        workbook = xlsxwriter.Workbook(f'Отчеты/Отчет изготовления тентового полотна.xlsx')
        # Форматы format()
        percent_format = workbook.add_format(
            {'border': 1, 'num_format': '0.00%', 'align': 'left', 'valign': 'vcenter'})
        percent_format_for_plan = workbook.add_format(
            {'num_format': '0.00%', 'align': 'left', 'valign': 'vcenter'})
        name_format = workbook.add_format(
            {'border': 1, 'bold': True, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        name_format_main = workbook.add_format(
            {'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        date_format = workbook.add_format(
            {'border': 1, 'text_wrap': True, 'num_format': 'dd MMM yy', 'align': 'center', 'valign': 'vcenter'})
        date_format_main = workbook.add_format(
            {'text_wrap': True, 'num_format': 'dd MMM yy', 'align': 'center', 'valign': 'vcenter'})
        special_numb = workbook.add_format(
            {'border': 1, 'num_format': '#0', 'align': 'center', 'valign': 'vcenter'})
        float_numb_w_board = workbook.add_format(
            {'border': 1, 'num_format': '#0.00', 'align': 'center', 'valign': 'vcenter'})
        numb_w_border = workbook.add_format(
            {'num_format': '#0.00', 'align': 'center', 'valign': 'vcenter'})
        # Форматы для объединнеых ячеек
        name_merge_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#0',
        })
        name_merge_format_spec = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#0',
            'bold': True
        })
        name_merge_format_2 = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#0',
            'fg_color': '#DDEBF7'
        })
        # Создаваемые листы
        worksheet_0 = workbook.add_worksheet(f'Сводный отчет')
        # Размер колонок
        worksheet_0.set_column(0, 0, 14)
        worksheet_0.set_column(1, 1, 20)
        worksheet_0.set_column(2, 2, 14)
        worksheet_0.set_column(3, 3, 17)
        worksheet_0.set_column(4, 4, 19)
        worksheet_0.set_column(5, 5, 13)
        worksheet_0.set_column(6, 6, 14)
        # Записи Ход изготовления ангара
        worksheet_0.merge_range(0, 0, 0, 4, f'Cводный отчет по участку производства тентового полотна',
                                name_merge_format)
        worksheet_0.write("F1", "Текущая дата:", name_format_main)
        worksheet_0.write("G1", f"{self.date_of_report}", date_format_main)
        worksheet_0.write("A4", f"План месяц, м2", name_format_main)
        tent_plan = 0.0
        try:
            for tent in db.get_tent_plan_for_mounth(self.data_of_mounth_plan):
                tent_plan = tent
        except:
            pass
        worksheet_0.write("B4", tent_plan, numb_w_border)
        worksheet_0.write("A5", f"Выполнение плана", name_format_main)
        worksheet_0.write_formula("B5", f"Выполнение плана", name_format_main)
        row_name = ['№', 'Проект', 'По проекту, м2', 'Изготовлено на текущ. момент, м2', 'Производственная готовность',
                    'Готовность к отгрузке']
        # Заголовки первой таблицы
        curnt_numb_row = 7
        num = 0
        worksheet_0.write_row(6, 0, row_name, name_format)
        curnt_special_row = 7
        for info in db.get_squer_pvh_progect():
            key = info[1]
            for progect_data in db.get_info_about_today_pvh_report(key, self.date_of_report):
                # №
                num += 1
                worksheet_0.write(curnt_numb_row, 0, num, special_numb)
                # Проект
                worksheet_0.write(curnt_numb_row, 1, progect_data[0], name_format)
                # По проекту
                worksheet_0.write(curnt_numb_row, 2, progect_data[1] / 1000, float_numb_w_board)
                # Готовность
                worksheet_0.write(curnt_numb_row, 3, progect_data[2] / 1000, float_numb_w_board)
                # Процент готовности
                worksheet_0.write(curnt_numb_row, 4, progect_data[3] / 100, percent_format)
                # Процент отгрузки
                worksheet_0.write(curnt_numb_row, 4, progect_data[4] / 100, percent_format)
                curnt_special_row += 1

            curnt_numb_row += 1
        worksheet_0.write_formula(curnt_special_row, 3, f'=SUM(D8:D{curnt_special_row})', float_numb_w_board)
        worksheet_0.write_formula("B5", f'=D{curnt_special_row + 1}/B4', percent_format_for_plan)
        worksheet_0.merge_range(curnt_special_row, 0, curnt_special_row, 2, f'Итого:',
                                name_merge_format_spec)
        worksheet_0.conditional_format(7, 4, curnt_special_row, 5, {'type': 'data_bar'})

        for progect_in_work in db.get_info_progect_pvh():
            # Создаваемые листы
            worksheet_1 = workbook.add_worksheet(f''.join(progect_in_work))
            # Размер колонок
            worksheet_1.set_column(3, 1, 10)
            worksheet_1.set_column(3, 2, 15)
            worksheet_1.set_column(3, 3, 15)
            worksheet_1.set_column(3, 4, 15)
            worksheet_1.set_column(3, 5, 15)
            worksheet_1.set_column(3, 6, 15)
            worksheet_1.set_column(3, 7, 15)
            worksheet_1.set_column(3, 8, 15)
            worksheet_1.set_column(3, 9, 15)
            worksheet_1.set_column(10, 10, 18)
            # Записи Ход изготовления ангара
            worksheet_1.merge_range(0, 0, 0, 10, f'Производственный отчет за {date.today()}', name_merge_format_spec)
            worksheet_1.write(1, 0, "Проект:", name_format)
            worksheet_1.merge_range(1, 1, 1, 10, f' ,'.join(progect_in_work), name_merge_format)
            worksheet_1.merge_range(2, 0, 2, 10, f'ПВХ покрытие внешнее и внутренее', name_merge_format_2)
            row_name = ['№', 'Дата', 'Раскрой полотна', 'Раскрой карманов', 'Раскрой нащельников','Сварка карманов',
                        'Приварить карманы', 'Пришить второй слой', 'Приварить нащельники', 'Упаковка',
                        'Общая готовности ТП']
            # Заголовки первой таблицы
            worksheet_1.write_row(3, 0, row_name, name_format)
            curnt_numb_row = 4
            num = 0
            for info in db.get_info_ab_day_report_pvh(f' ,'.join(progect_in_work)):
                num += 1
                # №
                worksheet_1.write(curnt_numb_row, 0, num, special_numb)
                # Дата
                worksheet_1.write(curnt_numb_row, 1, info[2], date_format)
                # % Раскрой полотна
                worksheet_1.write(curnt_numb_row, 2, info[3] / 100, percent_format)
                # % Раскрой карманов
                worksheet_1.write(curnt_numb_row, 3, info[4] / 100, percent_format)
                # % Раскрой Нащельников
                worksheet_1.write(curnt_numb_row, 4, info[5] / 100, percent_format)
                # % Сварка карманов
                worksheet_1.write(curnt_numb_row, 5, info[6] / 100, percent_format)
                # % Приварить карманы
                worksheet_1.write(curnt_numb_row, 6, info[7] / 100, percent_format)
                # % Полоса второго слоя
                worksheet_1.write(curnt_numb_row, 7, info[8] / 100, percent_format)
                # % нащельники приварить
                worksheet_1.write(curnt_numb_row, 8, info[9] / 100, percent_format)
                # % упаковать
                worksheet_1.write(curnt_numb_row, 9, info[10] / 100, percent_format)
                # % готовности утеплителя
                worksheet_1.write(curnt_numb_row, 10, info[11] / 100, percent_format)
                curnt_numb_row += 1
            worksheet_1.conditional_format(4, 2, curnt_numb_row, 10, {'type': 'data_bar'})

        workbook.close()
        os.startfile(f'Отчеты\Отчет изготовления тентового полотна.xlsx')

class ExcelMontage():
    def __init__(self, n_date, date_for_plan):
        super().__init__()
        self.date_of_report = n_date
        self.data_of_mounth_plan = date_for_plan
        self.report_excel_montage()


    def report_excel_montage(self):
        #  ОТЧЕТ ДЛЯ КМД
        workbook = xlsxwriter.Workbook(f'Отчеты/Отчет по монтажным работам.xlsx')
        # Форматы format()
        percent_format = workbook.add_format(
            {'border': 1, 'num_format': '0.00%', 'align': 'left', 'valign': 'vcenter'})
        percent_format_for_plan = workbook.add_format(
            {'num_format': '0.00%', 'align': 'left', 'valign': 'vcenter'})
        name_format = workbook.add_format(
            {'border': 1, 'bold': True, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        name_format_no_bold = workbook.add_format(
            {'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        name_format_main = workbook.add_format(
            {'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        date_format = workbook.add_format(
            {'border': 1, 'text_wrap': True, 'num_format': 'dd MMM yy', 'align': 'center', 'valign': 'vcenter'})
        date_format_main = workbook.add_format(
            {'text_wrap': True, 'num_format': 'dd MMM yy', 'align': 'center', 'valign': 'vcenter'})
        special_numb = workbook.add_format(
            {'border': 1, 'num_format': '#0', 'align': 'center', 'valign': 'vcenter'})
        float_numb_w_board = workbook.add_format(
            {'border': 1, 'num_format': '#0.00', 'align': 'center', 'valign': 'vcenter'})
        numb_w_border = workbook.add_format(
            {'num_format': '#0.00', 'align': 'center', 'valign': 'vcenter'})
        # Форматы для объединнеых ячеек
        name_merge_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#0',
        })
        name_merge_format_main = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#0',
            'bold': True
        })
        name_merge_format_spec = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#0',
            'bold': True
        })
        name_merge_format_spec_2 = workbook.add_format({
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': '#0',
            'bold': True
        })
        name_merge_format_2 = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#0',
            'fg_color': '#DDEBF7'
        })
        workbook.set_properties({
            'title': f'Производственный отчет',
            'subject': 'With document properties',
            'author': 'Ivan Metliaev',
            'manager': '',
            'company': 'Тентовые конструкции',
            'category': 'КМД',
            'keywords': 'КМД, Ангары, Металл',
            'created': datetime.today(),
            'comments': 'Created with Python and Ivan Metliaev program'})
        # Создаваемые листы
        worksheet_0 = workbook.add_worksheet(f'Сводный отчет')
        # Размер колонок
        worksheet_0.set_column(0, 0, 14)
        worksheet_0.set_column(1, 1, 20)
        worksheet_0.set_column(2, 2, 14)
        worksheet_0.set_column(3, 3, 17)
        worksheet_0.set_column(4, 4, 19)
        worksheet_0.set_column(5, 5, 13)
        worksheet_0.set_column(6, 6, 14)
        # Записи Ход изготовления ангара
        worksheet_0.merge_range(0, 0, 0, 4, f'Cводный отчет по Монтажным работам',
                                name_merge_format_main)
        worksheet_0.write("F1", "Текущая дата:", name_format_main)
        worksheet_0.write("G1", f"{self.date_of_report}", date_format_main)
        row_name = ['№', 'Проект', 'Производственная готовность']
        # Заголовки первой таблицы
        curnt_numb_row = 7
        num = 0
        worksheet_0.write_row(6, 0, row_name, name_format)
        curnt_special_row = 7
        for info in db.get_all_active_montage_progect():
            key = info[1]
            for progect_data in db.get_info_about_today_montage_report(key, self.date_of_report):
                # №
                num += 1
                worksheet_0.write(curnt_special_row, 0, num, special_numb)
                # Проект
                worksheet_0.write(curnt_special_row, 1, progect_data[0], name_format)
                # По проекту
                # worksheet_0.write(curnt_special_row, 2, progect_data[1] / 1000, float_numb_w_board)
                # Готовность
                # worksheet_0.write(curnt_special_row, 3, progect_data[2] / 1000, float_numb_w_board)
                # Процент
                worksheet_0.write(curnt_special_row, 2, progect_data[1] / 100, percent_format)
                curnt_special_row += 1
            curnt_numb_row += 1
        # worksheet_0.write_formula(curnt_special_row, 3, f'=SUM(D8:D{curnt_special_row})', float_numb_w_board)
        # worksheet_0.write_formula("B5", f'=D{curnt_special_row + 1}/B4', percent_format_for_plan)
        # worksheet_0.merge_range(curnt_special_row, 0, curnt_special_row, 2, f'Итого:',name_merge_format_spec)
        worksheet_0.conditional_format(7, 4, curnt_special_row, 4, {'type': 'data_bar'})
        worksheet_0.ignore_errors()
        for progect_in_work in db.get_info_progect_montage():
            worksheet_1 = workbook.add_worksheet(f''.join(progect_in_work))
            # Размер колонок
            worksheet_1.set_column(3, 1, 10)
            worksheet_1.set_column(3, 2, 14)
            worksheet_1.set_column(3, 3, 18)
            worksheet_1.set_column(3, 4, 18)
            worksheet_1.set_column(5, 5, 18)
            worksheet_1.set_column(6, 6, 18)
            worksheet_1.set_column(7, 7, 22)
            worksheet_1.set_column(8, 8, 40)
            worksheet_1.set_column(9, 9, 45)
            # Записи Ход изготовления ангара
            worksheet_1.merge_range(0, 0, 0, 9, f'Производственный отчет нарастающим итогом на отчетную дату',
                                    name_merge_format_spec)
            worksheet_1.write(1, 0, "Проект:", name_format)
            worksheet_1.merge_range(1, 1, 1, 9, f' ,'.join(progect_in_work), name_merge_format)
            worksheet_1.merge_range(2, 0, 2, 9, f'Монтажные работы', name_merge_format_2)
            row_name = ['№', 'Дата', 'Организационные работы', 'Монтаж металлокаркаса', 'Монтаж ограждающих конструкций', 'Монтаж инженерных систем', 'Завершающие работы', 'Общая готовность ангара', 'Выявленые проблемы', 'Решение проблемы']
            # Заголовки первой таблицы
            worksheet_1.write_row(3, 0, row_name, name_format)
            curnt_numb_row = 4
            num = 0
            for info in db.get_info_evr_report_montage(f''.join(progect_in_work)):
                num += 1
                # №
                worksheet_1.write(curnt_numb_row, 0, num, special_numb)
                # Дата
                worksheet_1.write(curnt_numb_row, 1, info[2], date_format)
                # % Организационные работы
                worksheet_1.write(curnt_numb_row, 2, info[3] / 100, percent_format)
                # % Монтаж металлокаркаса
                worksheet_1.write(curnt_numb_row, 3, info[4] / 100, percent_format)
                # % Монтаж ограждающих конструкций
                worksheet_1.write(curnt_numb_row, 4, info[5] / 100, percent_format)
                # % Монтаж инженерных систем'
                worksheet_1.write(curnt_numb_row, 5, info[6] / 100, percent_format)
                # % Завершающие работы
                worksheet_1.write(curnt_numb_row, 6, info[7] / 100, percent_format)
                # % Общая готовность ангара
                worksheet_1.write(curnt_numb_row, 7, info[8] / 100, percent_format)
                # Проблемы
                if info[9] is not None:
                    worksheet_1.write(curnt_numb_row, 8, info[9], name_format_no_bold)
                else:
                    worksheet_1.write(curnt_numb_row, 8, '', name_format_no_bold)
                # Решение
                if info[10] is not None:
                    worksheet_1.write(curnt_numb_row, 9, info[10], name_format_no_bold)
                else:
                    worksheet_1.write(curnt_numb_row, 9, '', name_format_no_bold)

                curnt_numb_row += 1

            worksheet_1.conditional_format(4, 2, curnt_numb_row, 7, {'type': 'data_bar'})
        workbook.close()
        os.startfile(f'Отчеты\Отчет по монтажным работам.xlsx')

class ExcelMontageReportWrite():
    def __init__(self, filename, date, real_date):
        super().__init__()
        self.path_to_filename = "{}".format(filename)
        self.date = date
        self.real_date = real_date


