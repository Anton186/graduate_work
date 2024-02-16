from datetime import datetime
from DataManipulator import copy_column_names_to_excel, copy_rows_to_excel_by_value, get_search_date, get_filename, \
    sum_quantity


def excel_to_excel():
    # Получаем и присваиваем имя файла
    filename = get_filename()
    if not filename:
        return None
    result_date = get_search_date()
    start_date, final_date = None, None
    if isinstance(result_date, datetime):
        start_date = result_date
    elif result_date and len(result_date) == 2:
        start_date, final_date = result_date

    output_file = 'выборка.xlsx'
    summary = 'отчет.xlsx'

    copy_column_names_to_excel(filename, output_file)
    copy_column_names_to_excel(filename, summary)
    value = input('Введите искомое значение (или оставьте пустым): ')
    if not (value or result_date):
        print('Ничего не введено для поиска!')
        return None
    copy_rows_to_excel_by_value(filename, output_file, value, start_date, final_date)

    sum_quantity(output_file, summary)
