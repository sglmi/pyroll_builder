from openpyxl import load_workbook


def sheet():
    wb = load_workbook(filename="sample.xlsx", data_only=True)
    sheet = wb.active
    return sheet


def headers(sheet):
    # if we want to skipe last cols if cell.value is not None
    header_cells = sheet[1]
    header = [cell.value for cell in header_cells]
    return header


def employees(sheet):
    # skip hedaer
    rows = list(sheet.iter_rows())[2:]
    # if value is not that mean we found an employee
    employees = [row for row in rows if row[0].value is not None]
    return employees[:-1]  # skipe last row


# cols: row, name, email
def items(employees, column_name="name"):
    columns = {"row": 0, "name": 1, "email": 20}
    index = columns.get(column_name)
    return [employee[index].value for employee in employees]


def main():
    wb = load_workbook(filename="sample.xlsx", data_only=True)
    sheet = wb.active

    emps = employees(sheet)
    print(len(emps))
    x = 0
    print(x)


if __name__ == "__main__":
    main()