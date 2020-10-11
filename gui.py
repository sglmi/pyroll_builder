import tkinter as tk
from tkinter import ttk

import code

checkbuttons = []
checkvars = []
emails_to_send = []
selected_rows = []
employees_id = []


def state():
    global checkvars
    indeces = []
    items = list(map((lambda var: var.get()), checkvars))
    for index, item in enumerate(items):
        if item:
            indeces.append(str(index))
    return indeces


# Highligh checked rows
def highlight(tree):
    items = state()
    # for item in items:
    #     tree.selection_add(item)
    selections = tree.selection()
    tree.selection_remove(selections)
    tree.selection_add(items)


# preview the employee payroll that has foucs on the tree
def preview(tree, label):
    # employee data
    employee_id = tree.focus()
    selected_employee_data = tree.item(employee_id)
    print(selected_employee_data)
    employee_name = selected_employee_data.get("values")[1]

    sh = code.sheet()
    employees = code.employees(sh)
    employee = code.employee(employees, employee_name)
    template = code.read_template()
    payroll_html = code.create_payroll_html(employee, template)
    payroll_pdf = code.html_to_pdf(payroll_html, "payroll.pdf")
    pix = code.pdf_to_image(payroll_pdf)

    # preview in tkinter
    imgdata = code.get_image_bytes(pix)
    tkimg = tk.PhotoImage(data=imgdata)
    label.img = tkimg
    label.config(image=label.img)


def test(tree):
    # global emails_to_send
    emails = []
    names = []
    # Extract email based on selected checkbox
    indeces = state()
    for item_id in indeces:
        name = tree.item(item_id).get("values")[1]
        email = tree.item(item_id).get("values")[2]
        names.append(name)
        emails.append(email)

    code.send_mail(names, emails)


def left_side(frame, righframe):
    global emails_to_send, employees_id

    check_frame = ttk.Frame(frame)
    tree_frame = ttk.Frame(frame)
    check_frame.grid(row=0, column=0, sticky="e")
    tree_frame.grid(row=0, column=1, sticky="sn")
    righframe.columnconfigure(1, weight=1)
    tree_frame.columnconfigure(0, weight=1)
    frame.columnconfigure(1, weight=1)
    # employee tree
    tree = ttk.Treeview(tree_frame, columns=("row", "name", "email"), show="headings")
    # tree heading
    tree.heading("row", text="Row")
    tree.heading("name", text="Name")
    tree.heading("email", text="email")
    # tree column conf
    tree.column("row", minwidth=10, width=30)

    ## Scroll bar
    scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    scrollbar.grid(row=0, column=1, sticky="esn")
    tree.configure(yscrollcommand=scrollbar.set)

    # get active sheet
    sheet = code.sheet()
    # list of employees
    employees = code.employees(sheet)
    # insert rows number
    rows_num = code.items(employees, column_name="row")
    names = code.items(employees, column_name="name")
    emails = code.items(employees, column_name="email")
    ids = range(len(names))

    # populate row num, names and emails on mployee Tree
    for i, row_num, name, email in zip(ids, rows_num, names, emails):
        employee_id = tree.insert("", "end", iid=i, values=(row_num, name, email))
        employees_id.append(employee_id)

    ## ===> Checkboxes <===
    global checkvars
    # get active sheet
    sheet = code.sheet()
    # list of employees
    employees = code.employees(sheet)
    row_nums = code.items(employees, column_name="row")

    for _ in row_nums:
        var = tk.BooleanVar()
        checkbutton = ttk.Checkbutton(check_frame, variable=var)
        checkbutton.config(command=lambda: highlight(tree))
        checkbutton.pack()

        checkvars.append(var)
    # tree bind
    # tree.bind("<<TreeviewSelect>>", onclick_tree_item)
    # grid tree and button
    tree.grid(row=0, column=0, sticky="ewsn")
    send_button = ttk.Button(frame, text="Send", command=lambda: test(tree))
    send_button.grid(row=1, column=0, sticky="s")
    # preview button
    preview_label = ttk.Label(righframe, text="Preview Payroll")
    preview_label.grid()
    preview_button = ttk.Button(
        frame, text="Preview", command=lambda: preview(tree, preview_label)
    )
    preview_button.grid(row=1, column=1, sticky="s")


def openfile():
    pass


# save selected employee payroll as pdf
def save_as():
    pass


def main():
    root = tk.Tk()
    root.title("Payroll")
    # root.geometry("800x400")

    # tkinter menubar
    menubar = tk.Menu(root)
    filemenu = tk.Menu(menubar, tearoff=0)
    filemenu.add_command(label="Open", command=openfile)
    filemenu.add_command(label="Save As", command=save_as)
    menubar.add_cascade(label="File", menu=filemenu)
    root.config(menu=menubar)

    mainframe = ttk.Frame(root)
    left_frame = ttk.Frame(mainframe, padding=5, relief="solid")
    right_frame = ttk.Frame(mainframe, padding=5, relief="groove")
    left_frame.grid(row=0, column=0, sticky=(tk.E, tk.W, tk.N, tk.S))
    right_frame.grid(row=0, column=1, sticky=(tk.E, tk.W, tk.N, tk.S))

    # display widgets
    # checkbox(check_frame)
    left_side(left_frame, right_frame)

    mainframe.pack(fill=tk.BOTH, expand=True)
    mainframe.columnconfigure(0, weight=1)
    mainframe.columnconfigure(1, weight=10)
    mainframe.rowconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)

    for child in mainframe.winfo_children():
        child.grid_configure(padx=5, pady=5)
    root.mainloop()


if __name__ == "__main__":
    main()