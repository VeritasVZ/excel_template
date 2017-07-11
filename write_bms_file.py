from template_properties import Properties
from BMS_template.bms_template import Template, InputParsing
from BMS_template.sheets.strategy import StrategySheet
from BMS_template.sheets.bvd_results import BvDResultsSheet
from BMS_template.sheets.search_matrix import SearchMatrix


#path = input('Enter file name - ')
#lang = input('Enter template language (UA, RU or EU(for English)) - ')
lang = 'RU'
file_name = 'input_data\Ruslana_Export.xlsx'
print('Progress status:')
doc = Template.create_template(lang)
data = InputParsing.get_results_from_export(file_name)
strategy = InputParsing.get_strategy_from_export(file_name)
"""build_template function takes fillowing arguments
    bms_template >>> template that is output of create_template function
    bvd_results_df >>> pandas dataframe with bvd results
    strategy >>> excel sheet with strategy
    lang >>> 'UA' 'RU' or 'EU'
"""
Template.build_template(bms_template=doc,bvd_results_df=data,strategy=strategy,lang=lang).save('results_data\BMS.xlsx')
