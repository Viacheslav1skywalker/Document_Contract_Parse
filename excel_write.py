import openpyxl
import os

# Добавляем данные в ячейки


# Сохраняем файл


class ExcelWrite:
    def __init__(self,path):
        self.path = path

    def write(self,data:list):
        self.check_and_delete(self.path+'\\' + 'result_programm_working_excel_parsing.xlsx')
        # Создаем новый файл Excel
        wb = openpyxl.Workbook()
        # Выбираем активный лист
        sheet = wb.active
        sheet['A1'] = 'фамилия'
        sheet['B1'] = 'имя'
        sheet['C1'] = 'отчество'
        sheet['D1'] = 'снилс'
        sheet['E1'] = 'серия'
        sheet['F1'] = 'номер'
        sheet['G1'] = 'кем выдан'
        sheet['H1'] = 'дата выдачи'
        sheet['I1'] = 'код подразделения'
        sheet['J1'] = 'индекс'
        sheet['K1'] = 'телефон'
        sheet['L1'] = 'адрес прописки'
        sheet['M1'] = 'кадастровый номер работ'
        sheet['N1'] = 'номер договора'
        sheet['O1'] = 'вид работ'
        sheet['P1'] = 'дата договора'
        sheet['Q1'] = 'стоимость работ ООО'
        sheet['R1'] = 'стоимость работ ИП'
        sheet['S1'] = 'файл, из которого взяты данные'
        write_session = 0
        index_letter = 0
        letters = 'ABCDEFGHIJKLMNOPQRS'
        for dict_data in data:
            write_session += 1
            index_letter = 0
            for value in dict_data:
                print(f'{letters[index_letter]}{str(write_session+1)}')
                print(dict_data[value])
                sheet[f'{letters[index_letter]}{str(write_session+1)}'] = dict_data[value]
                index_letter += 1

        wb.save('result_programm_working_excel_parsing.xlsx')
        wb.close()

    def check_and_delete(self,file_path):
        # Путь к файлу, который нужно проверить и удалить
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
            except:
                print('Ошибка: Закройте файл result_programm_working_excel_parsing.xlsx и запустите программу заново')
                raise PermissionError('Ошибка: Закройте файл result_programm_working_excel_parsing.xlsx и запустите программу заново')
        else:
            print('файл не был найден')




