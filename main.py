# from DataManipulator import sum_quantity
from VirtualCheck import env_check, moduls_install
from ExcelProcessing import excel_to_excel

env_check()
moduls_install()
excel_to_excel()
# Сообщение перед ожиданием нажатия клавиши
print("\nФормирование отчета завершено!\nДля завершения скрипта нажмите Enter...")
input()
