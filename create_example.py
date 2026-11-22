"""Создать таблицу только с вымышленными данными, не заменяя существующую."""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill, Alignment

path = Path(__file__).resolve().parent / 'Attachments' / 'certificates_example.xlsx'
path.parent.mkdir(exist_ok=True)
if path.exists():
    print(f'Готовый пример уже есть: {path}\nОткройте его и сохраните свою таблицу рядом с bot.py как certificates.xlsx.')
    raise SystemExit(0)
book = Workbook()
sheet = book.active
sheet.title = 'Подарки'
sheet.append(['user_login', 'full_name', 'certificate_text'])
sheet.append(['ivanov@example.org', 'Иванов Иван Петрович', 'Демонстрационный сертификат: EXAMPLE-001'])
sheet.append(['ivanov@example.org', 'Иванов Иван Петрович', 'Демонстрационный сертификат: EXAMPLE-002'])
sheet.append(['ivanov2@example.org', 'Иванов Иван Сергеевич', 'Демонстрационный сертификат: EXAMPLE-003'])
sheet.freeze_panes = 'A2'
sheet.auto_filter.ref = sheet.dimensions
for column, width in [('A', 32), ('B', 34), ('C', 68)]:
    sheet.column_dimensions[column].width = width
for cell in sheet[1]:
    cell.font = Font(bold=True, color='FFFFFF')
    cell.fill = PatternFill('solid', fgColor='1F4E78')
    cell.alignment = Alignment(wrap_text=True)
sheet['A1'].comment = Comment('Обязательно. Полный корпоративный адрес получателя. Замените пример своим адресом.', 'Инструкция')
sheet['B1'].comment = Comment('Необязательно. Фамилия, имя и отчество, если оно есть. Получатель определяется по user_login.', 'Инструкция')
sheet['C1'].comment = Comment('Обязательно. Полный текст сертификата. Замените демонстрационные значения своими.', 'Инструкция')
for row in sheet.iter_rows(min_row=2):
    for cell in row:
        cell.number_format = '@'
        cell.alignment = Alignment(wrap_text=True, vertical='top')
with path.open('xb') as file:
    book.save(file)
book.close()
print(f'Пример создан: {path}\nЗамените вымышленные данные своими и сохраните как certificates.xlsx.')
