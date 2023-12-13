import xlsxwriter

workbook = xlsxwriter.Workbook('report.xlsx')
worksheet = workbook.add_worksheet()

worksheet.set_column(0, 6, 20)
bold = workbook.add_format({'bold': True})

row, col = 0, 0
worksheet.write_string(row, col, 'Сервер', bold)
worksheet.write_string(row, col+1, 'Процессор', bold)
worksheet.write_string(row, col+2, 'Материнская плата', bold)
worksheet.write_string(row, col+3, 'Память', bold)
worksheet.write_string(row, col+4, 'Диск', bold)
worksheet.write_string(row, col+5, 'Сетевая', bold)
worksheet.write_string(row, col+6, 'Шильдик', bold)

row += 1

file = input("Enter the file name: ")

f = open(file, 'r', encoding='utf-8')

line_number = 1

memory_line = ""
processor_line = ""

for line in f:
    if line_number == 5:
        print("\nMemory Info:")
    if line_number >= 5 and line_number <= 20:
        # print(str(line_number-4) + ": " + line[96:114])
        memory_line += line[96:114] + ", "

    if line_number == 20:
        print(memory_line)
        worksheet.write_string(row, col+3, memory_line)

    if line_number == 27:
        print("\nMotherboard Info:")
    if line_number == 31:
        print(line[60:72])
        worksheet.write_string(row, col+2, line[60:72])

    if line_number == 34:
        print("\nProcessor Info:")
    if line_number >= 38 and line_number <= 39:
        # print(str(line_number-37)+ ": " + line[60:72])
        processor_line += line[60:72] + ", "

    if line_number == 39:
        print(processor_line)
        worksheet.write_string(row, col+1, processor_line)
    
    if line_number == 42: #Hard Info
        print("\n" + line[:14])
    if line_number == 43:
        print(line)
        worksheet.write_string(row, col+4, line)
    
    line_number += 1

#for line in f:

f.close()

workbook.close()
