import tkinter as tk
from tkinter import ttk

import code

checkbuttons = []
checkvars = []
emails_to_send = []
selected_rows = []


def state():
    global checkvars
    indeces = []
    items = list(map((lambda var: var.get()), checkvars))
    for index, item in enumerate(items):
        if item:
            indeces.append(index)
    return indeces

    # y = list(map((lambda var: var.get()), textvars))
    # # print(x)
    # # print(dir(chkvars[0]))
    # # print(chkvars[0].get())
    # # for item in chkvars:
    # #     if item.get() == 1:
    # #         print(chkvars.index(item.get()))
    # print(x, y)
    # global checkbuttons
    # print(dir(checkbuttons[0]))
    # global selected_rows
    # for checkbutton in checkbuttons:
    #     if checkbutton.instate(["selected"]):
    #         selected_rows.append(checkbutton["text"])

    # print(selected_rows)


def test(tree):
    global emails_to_send
    # Extract email based on selected checkbox
    indeces = state()
    # print(indeces)
    # filter based on first value of tree :: nums
    rows = tree.get_children()
    for row in rows:
        row_num = tree.item(row)["values"][0]
        if row_num in indeces:
            email = tree.item(row)["values"][1]
            emails_to_send.append(email)

    print(emails_to_send)


def checkbox(frame):
    global checkbuttons, checkvars
    # get active sheet
    sheet = code.sheet()
    # list of employees
    employees = code.employees(sheet)
    row_nums = code.items(employees, column_name="row")

    for row in row_nums:
        var = tk.BooleanVar()
        checkbutton = ttk.Checkbutton(frame, variable=var)
        checkbutton.pack()
        checkbuttons.append(checkbutton)
        checkvars.append(var)


def left_side(frame):
    global emails_to_send

    # employee tree
    tree = ttk.Treeview(frame, columns=("row", "name", "email"), show="headings")
    # tree heading
    tree.heading("row", text="Row")
    tree.heading("name", text="Name")
    tree.heading("email", text="email")
    # tree column conf
    tree.column("row", minwidth=10, width=30)

    # get active sheet
    sheet = code.sheet()
    # list of employees
    employees = code.employees(sheet)
    # insert rows number
    names = code.items(employees, column_name="name")
    emails = code.items(employees, column_name="email")
    nums = range(len(names))

    # populate row num, names and emails on mployee Tree
    for num, name, email in zip(nums, names, emails):
        tree.insert("", "end", values=(num, name, email))

    # grid tree and button
    tree.grid()
    send_button = ttk.Button(frame, text="Send", command=lambda: test(tree))
    send_button.grid(row=1)


def main():
    root = tk.Tk()
    root.title("Payroll")

    mainframe = ttk.Frame(root)
    left_frame = ttk.Frame(mainframe, padding=5, relief="solid")
    right_frame = ttk.Frame(mainframe, padding=5, relief="solid")
    check_frame = ttk.Frame(mainframe, padding=5, relief="solid")
    check_frame.grid(row=0, column=0, sticky=(tk.N, tk.S))
    left_frame.grid(row=0, column=1, sticky=(tk.E, tk.W, tk.N, tk.S))
    right_frame.grid(row=0, column=2, sticky=(tk.E, tk.W, tk.N, tk.S))
    label = ttk.Button(right_frame, text="Preview Payroll")
    label.grid()

    # display widgets
    checkbox(check_frame)
    left_side(left_frame)

    mainframe.pack(fill=tk.BOTH, expand=True)
    mainframe.columnconfigure(1, weight=1)
    mainframe.columnconfigure(2, weight=6)
    mainframe.rowconfigure(0, weight=1)

    for child in mainframe.winfo_children():
        child.grid_configure(padx=5, pady=5)
    root.mainloop()


if __name__ == "__main__":
    main()