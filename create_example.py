"""Создать таблицу только с вымышленными данными, не заменяя существующую."""
from pathlib import Path
from openpyxl import Workbook

path = Path(__file__).resolve().parent / 'certificates_example.xlsx'
book = Workbook()
sheet = book.active
sheet.title = 'Certificates'
sheet.append(['user_login', 'full_name', 'certificate_text'])
sheet.append(['ivanov@example.org', 'Иван Иванов', 'Демонстрационный сертификат: EXAMPLE-001'])
sheet.append(['ivanov@example.org', 'Иван Иванов', 'Демонстрационный сертификат: EXAMPLE-002'])
with path.open('xb') as file:
    book.save(file)
book.close()
print(f'Пример создан: {path}\nЗамените вымышленные данные своими и сохраните как certificates.xlsx.')
