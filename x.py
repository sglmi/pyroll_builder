import code
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import time


class Navebar(ttk.Frame):
    def __init__(self, parent, mainframe, *arg, **kwargs):
        super().__init__(parent, *arg, *kwargs)
        self.parent = parent
        self.mainframe = mainframe
        self.filename = ""
        menubar = tk.Menu()
        # File
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Open", command=self.openfile)
        filemenu.add_command(label="Save As", command=self.save_as)
        menubar.add_cascade(label="File", menu=filemenu)
        # Help
        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="Help")
        helpmenu.add_command(label="About")
        menubar.add_cascade(label="Help", menu=helpmenu)
        self.parent.config(menu=menubar)

    def openfile(self):
        self.filename = filedialog.askopenfilename(
            filetypes=(
                ("Excel File", "*.xlsx"),
                ("Excel File (old)", "xls"),
                ("All files", "*.*"),
            )
        )
        if self.filename:
            self.mainframe.destroy()
            self.mainframe = MainForm(self.parent, self.filename)
            self.mainframe.grid(row=0, column=0, sticky=(tk.N, tk.E, tk.S, tk.W))
        else:
            self.filename = "small.xlsx"

    def save_as(self):

        data = self.mainframe.action.checked_employees()

        sheet = code.sheet(self.filename)
        employees = code.employees(sheet)
        template = code.read_template()
        for item in data:
            name = item.get("values")[1]
            row = item.get("values")[0]
            emp = code.employee(employees, name)
            html = code.create_payroll_html(emp, template)
            pdf = code.html_to_pdf(html, f"{row}.pdf")
            print(f"Pdf file created for {pdf}")
            # print(name, row, item)


class CheckFrame(tk.Frame):
    def __init__(self, parent, *arg, **kwargs):
        super().__init__(parent, *arg, **kwargs)
        self.checkvars = []

    def create_checkbuttons(self, number):
        r = 0
        for _ in range(number):
            var = tk.BooleanVar()
            checkbutton = ttk.Checkbutton(self, variable=var)
            self.checkvars.append(var)
            checkbutton.grid(row=r, column=0, sticky=tk.N)
            r += 1

    # return selected checkbuttons index
    def state(self):
        return list(map((lambda var: var.get()), self.checkvars))

    def checkeds(self):
        return [str(index) for index, item in enumerate(self.state()) if item]


class EmployeeFrame(tk.Frame):
    def __init__(self, parent, filename, *arg, **kwargs):
        super().__init__(parent, *arg, **kwargs)
        self.tree = ttk.Treeview(self)
        self.tree.pack(expand=True, fill=tk.Y)
        self.employees_id = []
        self._config()
        self._headings()
        # self._scrollbar(parent)
        self._populate_data(filename)

    def _config(self):
        self.tree.configure(
            columns=("row", "name", "email"),
            show="headings",
        )
        self.tree.column("row", minwidth=15, width=50)

    def _headings(self):
        self.tree.heading("row", text="Row")
        self.tree.heading("name", text="Name")
        self.tree.heading("email", text="email")

    # def _scrollbar(self, parent):
    #     scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self.yview)
    #     scrollbar.grid(row=0, column=1, sticky="esn")
    #     self.configure(yscrollcommand=scrollbar.set)

    def _populate_data(self, filename):
        sheet = code.sheet(filename)
        employees = code.employees(sheet)
        names = code.items(employees, column_name="name")
        emails = code.items(employees, column_name="email")
        ids = range(len(names))

        # populate row num, names and emails on mployee Tree
        for i, name, email in zip(ids, names, emails):
            employee_id = self.tree.insert(
                "", "end", iid=i, values=(i + 1, name, email)
            )
            self.employees_id.append(employee_id)

    def selected_employees(self):
        return [self.tree.item(item) for item in self.tree.selection()]

    def number_of_employees(self):
        return len(self.employees_id)


class MainForm(tk.Canvas):
    def __init__(self, parent, filename, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.filename = filename
        self.config(scrollregion=[0, 0, 500, 700])
        self.config(bg="#aaaaaa")
        self.parent = parent
        self._checkbuttons()
        self._employees()
        self._scrollbar()

    def _checkbuttons(self):
        checkframe = CheckFrame(self.parent)
        emps_num = 10
        checkframe.create_checkbuttons(emps_num)
        self.create_window(10, 10, anchor="nw", window=checkframe)

    def _employees(self):
        emp_frame = EmployeeFrame(self, self.filename)
        emp_frame.columnconfigure(0, weight=1)
        emp_frame.rowconfigure(0, weight=1)
        self.create_window(40, 10, anchor="nw", window=emp_frame)

    def _scrollbar(self):
        scrollbar = ttk.Scrollbar(self.parent, orient="vertical", command=self.yview)
        scrollbar.pack(side="right", fill="y")
        self.configure(yscrollcommand=scrollbar.set)


def main():
    root = tk.Tk()
    root.title("Payroll")
    root.geometry("700x500")
    # ++++++++++++++++++
    default_excel_filename = "sample.xlsx"
    mainframe = ttk.Frame(root, padding=10, relief="solid")
    maincanvas = MainForm(mainframe, default_excel_filename)
    navbar = Navebar(root, maincanvas)
    navbar.pack()
    maincanvas.pack(fill=tk.BOTH, expand=True)
    mainframe.pack(fill=tk.BOTH, expand=True)

    # columnconfigure
    maincanvas.columnconfigure(0, weight=1)
    maincanvas.rowconfigure(0, weight=1)
    mainframe.columnconfigure(0, weight=1)
    mainframe.rowconfigure(0, weight=1)
    # ++++++++++++++++++
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.mainloop()


if __name__ == "__main__":
    main()