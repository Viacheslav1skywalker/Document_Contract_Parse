import parse_excel
import parse_files
import write_in_file

def main(path):
    test = parse_excel.ExcelParsing(path)
    list_values = test.main_call()
    write_in_file.ExcelWrite().write(list_values)
    obj = parse_files.ParsingContractDocument(r'C:\Users\Slava-Stat\Desktop\Проекты_Python\Excel_parsing\test_data\юр лица 2023')
    list_values = obj.main_parse()
    write_in_file.ExcelWriteFromWordFiles().write(list_values)
main(r'C:\Users\Slava-Stat\Desktop\Проекты_Python\Excel_parsing\test_data')