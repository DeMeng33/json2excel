from collections import OrderedDict
import xlsxwriter
import json


keyColumns = []


def json_to_excel(ws, data, row=0):
    if isinstance(data, list):
        row -= 1
        for value in data:
            row = json_to_excel(ws, value, row + 1)
    elif isinstance(data, dict):
        for key, value in data.itjsoems():
            if isinstance(value, (dict,list)):
                json_to_excel(ws, value, row)
            else:
                ws.write(row, index(key), value)
    else:
        ws.write(row, index(data), data)

    return row


def index(key):
    try:
        return keyColumns.index(key)
    except ValueError:
        keyColumns.append(key)
        return len(keyColumns) - 1


if __name__ == '__main__':
    filePath = 'your_file_path/aaa.json'
    with open(filePath) as f:
        data = json.load(f, object_pairs_hook=OrderedDict)
    wb = xlsxwriter.Workbook("output1.xlsx")
    ws = wb.add_worksheet()
    json_to_excel(ws, data)
    wb.close()