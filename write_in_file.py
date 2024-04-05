import openpyxl
import os


class MainWrite:
    def check_and_delete(self, file_path):
        # Путь к файлу, который нужно проверить и удалить
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
            except:
                print(f'Ошибка: Закройте файл result_programm_working_excel_parsing.xlsx и запустите программу заново')
                raise PermissionError(
                    file_path)
        else:
            print('файл не был найден')

class ExcelWrite(MainWrite):
    def write(self,data:list):
        self.check_and_delete('result_programm_working_excel_parsing.xlsx')
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

class ExcelWriteFromWordFiles(MainWrite):

    def write(self, data: list,path=""):
        self.check_and_delete('result_programm_working_word_parsing.xlsx')
        # Создаем новый файл Excel
        wb = openpyxl.Workbook()
        # Выбираем активный лист
        sheet = wb.active
        sheet['A1'] = 'название организации'
        sheet['B1'] = 'инн'
        sheet['C1'] = 'кпп'
        sheet['D1'] = 'огрн'
        sheet['E1'] = 'расчетный счет'
        sheet['F1'] = 'корреспондентский счет'
        sheet['G1'] = 'бик'
        sheet['H1'] = 'адрес'
        sheet['I1'] = 'телефон'
        sheet['J1'] = 'номер договора'
        sheet['K1'] = 'дата договора'
        sheet['L1'] = 'кадастровый номер работ'
        sheet['M1'] = 'вид работ'
        sheet['N1'] = 'стоимость работ'
        sheet['O1'] = 'файл, из которого взяты данные'
        write_session = 0
        index_letter = 0
        letters = 'ABCDEFGHIJKLMNO'
        for dict_data in data:
            write_session += 1
            index_letter = 0
            for value in dict_data:
                print(f'{letters[index_letter]}{str(write_session + 1)}')
                print(dict_data[value])
                sheet[f'{letters[index_letter]}{str(write_session + 1)}'] = dict_data[value]
                index_letter += 1
        wb.save('result_programm_working_word_parsing.xlsx')
        wb.close()




