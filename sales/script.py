import csv
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle


def set_column_h_to_date_format(file_path):
    workbook = load_workbook(filename=file_path)
    worksheet = workbook.active
    
    date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
    
    if "date_style" not in workbook.named_styles:
        workbook.add_named_style(date_style)
    
    for row in worksheet['H']:
        if row.value:
            row.style = date_style
    
    workbook.save(filename=file_path)


def append_rows_to_xlsx(file_path, rows):
    workbook = load_workbook(filename=file_path)
    worksheet = workbook.active
    
    for row in rows:
        fields = row.split(';')
        worksheet.append(fields)
    
    workbook.save(filename=file_path)


def get_current_date():
    return datetime.now().strftime("%d/%m/%Y")


def read_csv_into_dict(file_path):
    with open(file_path, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        return {row[reader.fieldnames[0]]: row for row in reader}


def find_rows_only_in_first_csv(file_path1, file_path2):
    csv1_dict = read_csv_into_dict(file_path1)
    csv2_dict = read_csv_into_dict(file_path2)

    only_in_csv1 = {id: csv1_dict[id] for id in csv1_dict if id not in csv2_dict}
    
    return only_in_csv1


previous = 'C:\\Users\\erixy\\OneDrive\\Work\\_Estoques\\anterior.csv'
current = 'C:\\Users\\erixy\\OneDrive\\Work\\_Estoques\\atual.csv'
sales = 'C:\\Users\\erixy\\OneDrive\\Work\\_Relatórios\\VENDAS.xlsx'

rows_only_in_first_file = find_rows_only_in_first_csv(previous, current)

rows_with_date = []
for id, row in rows_only_in_first_file.items():
    rows_with_date.append(id + str(";") + str(get_current_date()))

append_rows_to_xlsx(sales, rows_with_date)
set_column_h_to_date_format(sales)