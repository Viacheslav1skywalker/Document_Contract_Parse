from openpyxl import load_workbook

data_file = 'C:\\Users\Slava-Stat\Desktop\Проекты_Python\Excel_parsing\пример_файлов\\2023-0256.xlsm'

# Load the entire workbook.
wb = load_workbook(data_file)
ws = wb['ИД']
# List all the sheets in the file.

for i in list(ws.rows)[0]:
    print(i.value)