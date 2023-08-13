import openpyxl
import re
def column_finder():
    wb = openpyxl.load_workbook("Book1.xlsx")
    ws = wb["mark"]
    titles = ws.iter_cols(min_col=1,max_col=3,min_row=1,max_row=1)
    for title in titles:
        for cell in title:
            if cell.value.lower() == "address":
                return cell.coordinate[0]
numbers = []
def number_adder():
    global numbers
    wb = openpyxl.load_workbook("Book1.xlsx")
    ws = wb["mark"]
    maxrow = ws.max_row
    rows = ws[column_finder()+"2":f"{column_finder()}{maxrow}"]
    rowcount = 2
    for row in rows:
        for cell in row:
            try:
                coordinate10 = f"{chr(ord(column_finder()) + 1)}{rowcount}"
                cellval = cell.value
                pattern = r'\d{10,11}'
                matches = re.findall(pattern, cellval)
                for match in matches:
                    numbers.append(match)
                result_string = ', '.join(map(str, numbers))
                ws[coordinate10] = result_string
                wb.save("Book1.xlsx")
                rowcount += 1
                numbers = []
            except TypeError:
                pass
print(number_adder())
