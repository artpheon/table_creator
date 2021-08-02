from openpyxl import Workbook
from openpyxl.worksheet.table import Table


def get_records(delim, file):
    text_file = open(file)
    records = []
    text_file.seek(0)
    for record in text_file.readlines():
        records.append(record.rstrip('\n').split(delim))
    text_file.close()
    return records


def fill_workbook(records, size, result):
    workbook = Workbook()
    workbook.save(result)
    sheet = workbook['Sheet']
    for row in records:
        sheet.append(row)
    table = Table(displayName='Table', ref=size)
    sheet.add_table(table)
    workbook.save(result)
    workbook.close()


def create_document(delim=',', file='./random.csv', result='./result.xlsx', size='A1:A2'):
    records = get_records(delim=delim, file=file)
    fill_workbook(records, size, result)


if __name__ == '__main__':
    create_document(delim=',', file='./new_table.csv', result='./result.xlsx', size='A1:H8')
