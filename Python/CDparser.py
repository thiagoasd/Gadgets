import re
import sys
from os import listdir;
from os.path import isdir, join;
from openpyxl import Workbook;
from openpyxl.cell import get_column_letter;

path = str(sys.argv[1]);
book = Workbook();
sheet = book.active;

row_number = 1;
for f in listdir(path):
    column_number = 1;
    cpath = join(path, f);
    if isdir(cpath):
        m = re.search('(\[.*\])(.*)(\[.*\])', cpath);
        if m is None or len(m.groups()) < 3:
            print("Error " + cpath);
            continue;
        cell = sheet.cell(row=row_number, column=column_number);
        cell.value = m.group(1);
        print(cell.value);
        cell = sheet.cell(row=row_number, column=column_number + 1);
        cell.value = m.group(2);
        print(cell.value);
        cell = sheet.cell(row=row_number, column=column_number + 2);
        cell.value = m.group(3);
        print(cell.value);
        row_number += 1;
        sheet.column_dimensions[get_column_letter(row_number -1)].width = 50;

book.save("test.xlsx");

# TODO SEPARATE FLACS
