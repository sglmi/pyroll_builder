import tkinter as tk
from tkinter import ttk
import code


class Preview(tk.Toplevel):
    pass


class Menubar(tk.Menu):
    def __init__(self, parent, *arg, **kwargs):
        super().__init__(parent, *arg, *kwargs)
        # File
        filemenu = tk.Menu(self, tearoff=0)
        filemenu.add_command(label="Open", command=self.openfile)
        filemenu.add_command(label="Save As", command=self.save_as)
        self.add_cascade(label="File", menu=filemenu)
        # Help
        helpmenu = tk.Menu(self, tearoff=0)
        helpmenu.add_command(label="Help")
        helpmenu.add_command(label="About")
        self.add_cascade(label="Help", menu=helpmenu)

    def openfile(self):
        pass

    def save_as(self):
        pass


class EmployeeTree(ttk.Treeview):
    def __init__(self, parent, *arg, **kwargs):
        super().__init__(parent, *arg, **kwargs)
        self.employees_id = []
        self._config()
        self._headings()
        self._scrollbar(parent)
        self._populate_data()

    def _config(self):
        self.configure(
            columns=("row", "name", "email"),
            show="headings",
        )
        self.column("row", minwidth=15, width=50)

    def _headings(self):
        self.heading("row", text="Row")
        self.heading("name", text="Name")
        self.heading("email", text="email")

    def _scrollbar(self, parent):
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self.yview)
        scrollbar.grid(row=0, column=1, sticky="esn")
        self.configure(yscrollcommand=scrollbar.set)

    def _populate_data(self):
        sheet = code.sheet()
        employees = code.employees(sheet)
        names = code.items(employees, column_name="name")
        emails = code.items(employees, column_name="email")
        ids = range(len(names))

        # populate row num, names and emails on mployee Tree
        for i, name, email in zip(ids, names, emails):
            employee_id = self.insert("", "end", iid=i, values=(i + 1, name, email))
            self.employees_id.append(employee_id)

    def selected_employees(self):
        return [self.item(item) for item in self.selection()]


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
    def __init__(self, parent, *arg, **kwargs):
        super().__init__(parent, *arg, **kwargs)
        self.employee_tree = EmployeeTree(self)
        self.employee_tree.grid(row=0, column=0)

    def number_of_employees(self):
        return len(self.employee_tree.employees_id)


class MainFrame(tk.Frame):
    def __init__(self, parent, *arg, **kwargs):
        super().__init__(parent, *arg, **kwargs)
        self.employee_frame = EmployeeFrame(self)
        self.employee_frame.grid(row=0, column=1)
        # get employee number

        self.checkframe = CheckFrame(self)
        self.checkframe.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E))
        self.checkframe.create_checkbuttons(self.employee_frame.number_of_employees())

        ttk.Button(self, text="Test", command=self.test).grid(row=1, column=0)

    def checked_employees(self):
        employees_data = []
        for index in self.checkframe.checkeds():
            employees_data.append(self.employee_frame.employee_tree.item(index))
        return employees_data

    def test(self):
        print(self.checked_employees())


def main():
    root = tk.Tk()
    root.title("Payroll")
    # ++++++++++++++++++
    menubar = Menubar(root)
    mainframe = MainFrame(root)
    mainframe.grid(row=0, column=0)
    root.config(menu=menubar)
    # ++++++++++++++++++
    root.mainloop()


if __name__ == "__main__":
    main()