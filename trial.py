from openpyxl import Workbook, load_workbook
from BMS_template.bms_template import Properties
from BMS_template.bms_template import Template, InputParsing
from BMS_template.sheets.search_matrix import SearchMatrix
from BMS_template.sheets.manual_steps import ManualStep

file_name = 'input_data\Ruslana_Export.xlsx'
file = Template.create_template('UA')
df = InputParsing.get_results_from_export(file_name)
ManualStep.steps_table(file).save('results_data\empty_book.xlsx')
#print(ManualStep.define_step_name_type('RU'))

"""
def input_file_check(file):
    correct_header_names = ['Компания', 'Юридическая форма', 'ЕДРПОУ/ИИН', 'Номер телефона', 'Город', 'Адрес',
                            'NACE Rev. 2.', 'Адрес в интернете', 'Перечень деятельности']
    input_file = load_workbook(file)
    sheet = input_file.active
    input_error_message = []
    header = []
    for cell in sheet[1]:
        header.append(cell.value)
    for cell in correct_header_names:
        if cell not in header:
            input_error_message.append('Error in field ' + cell)
    if len(input_error_message) == 0:
        return True
    else:
        return False, input_error_message

print(input_file_check(file))
"""