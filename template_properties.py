# -*- coding: utf-8 -*-
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import *
from openpyxl.worksheet.worksheet import *
from openpyxl.utils import get_column_letter
from BMS_template.utilities import DesignUtils

class Properties:


    def template_type(lang):

        sheet_names_types = {
            'sheet_1': ('Результати з бази даних', 'Результаты из базы данных', 'BvD Results'),
            'sheet_2': ('Стратегія пошуку', 'Стратегия поиска', 'Search Strategy'),
            'sheet_3': ('Кроки стратегії', 'Шаги стратегии', 'BvD Strategy Steps'),
            'sheet_4': ('Додаткові критерії вибору', 'Дополнительные критерии отбора', 'Manual steps'),
            'sheet_5': ('Спостереження', 'Наблюдения', 'Search Matrix'),
            'sheet_6': ('Перевірка діяльності', 'Проверка деятельности', 'Web review'),
            'sheet_7': ('Діапазон рентабельності','Диапазон рентабельности','Market range'),
            'sheet_8': ('Порівняння з попереднім роком','Сравнение с прошлым годом','Reconciliation matrix')
        }
        sheets = list(sheet_names_types.keys())
        sheet_names = {}
        for sheet_no in range(len(sheets)):
            type_item = sheet_names_types.get(sheets[sheet_no])
            if lang == 'UA': sheet_names[sheets[sheet_no]] = type_item[0]
            elif lang == 'RU': sheet_names[sheets[sheet_no]] = type_item[1]
            elif lang == 'EU': sheet_names[sheets[sheet_no]] = type_item[2]
            else: print('Wrong language')
        return sheet_names

    def lang_from_template_type(bms_template):
        sheet_names = bms_template.get_sheet_names()
        lang = 'Undetermined'
        if sheet_names == list(Properties.template_type('UA').values()):
            lang = 'UA'
        elif sheet_names == list(Properties.template_type('RU').values()):
            lang = 'RU'
        elif sheet_names == list(Properties.template_type('EU').values()):
            lang = 'EU'
        return lang

    def define_col_width(sheet,col_range,etalon_row):

        empty_col_width = 6
        max_col_width = 16
        for col in range(col_range[0],col_range[1]):
            cell_content = str(sheet.cell(column=col,row = etalon_row).value)
            if cell_content is None:
                sheet.column_dimensions[get_column_letter(col)].width = empty_col_width
            word_lengths = []
            word_count = 0
            for word in cell_content.split():
                word_count = word_count+1
                word_lengths.append(len(word))
            if max(word_lengths)<= empty_col_width and word_count == 1:
                sheet.column_dimensions[get_column_letter(col)].width = empty_col_width
            elif max(word_lengths)*1.5 > max_col_width and word_count <= 8:
                sheet.column_dimensions[get_column_letter(col)].width = max_col_width
            elif word_count <= 8 or not (word_count == 1 and word_lengths[0]*1.5 <= 3):
                sheet.column_dimensions[get_column_letter(col)].width = max(word_lengths)*1.5
            elif word_count>=8:
                sheet.column_dimensions[get_column_letter(col)].width = max_col_width*2
            else:
                sheet.column_dimensions[get_column_letter(col)].width = empty_col_width
        return sheet

    def number_to_format(sheet,column,row, format_str):
        cell = sheet.cell(column=column, row=row)
        try:
            n = float(cell.value)
        except ValueError:
            n = cell.value
        _ = sheet.cell(column=column,row=row, value = n)
        sheet.cell(column=column, row=row).number_format = format_str
        return sheet


class Styles:
    def basic_style(cell):

        cell.font = Font(name='Verdana', size=8)
        cell.border = Border(left=Side(border_style='thin',color=colors.BLACK),
                      right=Side(border_style='thin',color=colors.BLACK),
                      top=Side(border_style='thin',color=colors.BLACK),
                      bottom=Side(border_style='thin',color=colors.BLACK))
        cell.alignment = Alignment(horizontal='left',vertical='bottom',wrap_text=False)
        return cell


    def header_style_dark_blue(cell):

        cell.font = Font(name='Verdana', size=8, bold=True, color=colors.WHITE)
        cell.border = Border(left=Side(border_style='thin', color=colors.WHITE),
                             right=Side(border_style='thin', color=colors.WHITE),
                             top=Side(border_style='thin', color=colors.WHITE),
                             bottom=Side(border_style='thin', color=colors.WHITE))
        cell.fill = PatternFill(patternType='solid',fgColor=DesignUtils.color_palette()['BLUE6'])
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        return cell

    def header_style_green(cell):

        cell.font = Font(name='Verdana', size=8, bold=True, color=colors.WHITE)
        cell.border = Border(left=Side(border_style='thin', color=colors.WHITE),
                             right=Side(border_style='thin', color=colors.WHITE),
                             top=Side(border_style='thin', color=colors.WHITE),
                             bottom=Side(border_style='thin', color=colors.WHITE))
        cell.fill = PatternFill(patternType='solid',fgColor=DesignUtils.color_palette()['DELOITTEGREEN'])
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        return cell

    def header_style_light_blue(cell):

        cell.font = Font(name='Verdana', size=8, bold=True, color=colors.WHITE)
        cell.border = Border(left=Side(border_style='thin', color=colors.WHITE),
                             right=Side(border_style='thin', color=colors.WHITE),
                             top=Side(border_style='thin', color=colors.WHITE),
                             bottom=Side(border_style='thin', color=colors.WHITE))
        cell.fill = PatternFill(patternType='solid',fgColor=DesignUtils.color_palette()['BLUE2'])
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        return cell

    def search_strategy_style(cell):

        cell.font = Font(name='Verdana', size=8, bold=True, color=DesignUtils.color_palette()['BLUE6'])
        cell.border = Border(left=Side(border_style='thin', color=DesignUtils.color_palette()['BLUE6']),
                             right=Side(border_style='thin', color=DesignUtils.color_palette()['BLUE6']),
                             top=Side(border_style='thin', color=DesignUtils.color_palette()['BLUE6']),
                             bottom=Side(border_style='thin', color=DesignUtils.color_palette()['BLUE6']))
        cell.fill = PatternFill(patternType='solid', fgColor=DesignUtils.color_palette()['LIGHTBERLINBLUE'])
        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

        return cell

    def thick_border_white_cell_style(cell):

        cell.font = Font(name='Verdana', size=8, bold=True, color=DesignUtils.color_palette()['BLUE6'])
        cell.border = Border(left=Side(border_style='thick', color=DesignUtils.color_palette()['BLUE6']),
                             right=Side(border_style='thick', color=DesignUtils.color_palette()['BLUE6']),
                             top=Side(border_style='thick', color=DesignUtils.color_palette()['BLUE6']),
                             bottom=Side(border_style='thick', color=DesignUtils.color_palette()['BLUE6']))
        cell.alignment = Alignment(horizontal='center',vertical='center')

        return cell

    def thick_border_fill_cell_style(cell):

        cell.font = Font(name='Verdana', size=8, bold=True, color=DesignUtils.color_palette()['BLUE6'])
        cell.border = Border(left=Side(border_style='thick', color=DesignUtils.color_palette()['BLUE6']),
                             right=Side(border_style='thick', color=DesignUtils.color_palette()['BLUE6']),
                             top=Side(border_style='thick', color=DesignUtils.color_palette()['BLUE6']),
                             bottom=Side(border_style='thick', color=DesignUtils.color_palette()['BLUE6']))
        cell.fill = PatternFill(patternType='solid',fgColor=DesignUtils.color_palette()['LIGHTBERLINBLUE'])
        cell.alignment = Alignment(horizontal='center',vertical='center')

        return cell

    def thin_border_fill_cell_style(cell):

        cell.font = Font(name='Verdana', size=8, bold=True, color=DesignUtils.color_palette()['BLUE6'])
        cell.border = Border(left=Side(border_style='thin', color=DesignUtils.color_palette()['BLUE6']),
                             right=Side(border_style='thin', color=DesignUtils.color_palette()['BLUE6']),
                             top=Side(border_style='thin', color=DesignUtils.color_palette()['BLUE6']),
                             bottom=Side(border_style='thin', color=DesignUtils.color_palette()['BLUE6']))
        cell.fill = PatternFill(patternType='solid',fgColor=DesignUtils.color_palette()['LIGHTBERLINBLUE'])
        cell.alignment = Alignment(horizontal='left',vertical='center',wrap_text=True)

        return cell