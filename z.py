import code
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import time


class EmployeeFrame(ttk.Frame):
    def __init__(self, parent, filename, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.filename = filename
        self.parent = parent
        self.chbvars = []
        self.select_all = tk.BooleanVar()

        self.number = 0
        # display header

    def _config(self):
        for child in self.winfo_children():
            child.pack_configure(pady=20)

    def employees_data(self):
        data = []
        sheet = code.sheet(self.filename)
        employees = code.employees(sheet)
        names = code.items(employees, column_name="name")
        emails = code.items(employees, column_name="email")
        ids = range(len(names))
        for eid, name, email in zip(ids, names, emails):
            data.append({"id": eid, "name": name, "email": email})
        self.number = len(data)
        return data

    def create_record(self, employees):
        # header
        frame = ttk.Frame(self)
        frame.pack(fill=tk.X, expand=True)
        # header widgets
        ttk.Checkbutton(frame, variable=self.select_all).grid(
            row=0, column=0, sticky="nw"
        )
        ttk.Label(frame, text="Row").grid(row=0, column=1, sticky="nw")
        ttk.Label(frame, text="Name").grid(row=0, column=2, sticky="nw")
        ttk.Label(frame, text="Email").grid(row=0, column=3, sticky="nw")
        ttk.Label(frame, text="_" * 100).grid(row=1, column=0, columnspan=5, sticky="n")

        for employee in employees:
            eid, name, email = employee.values()

            # Variables
            var = tk.IntVar()
            ttk.Checkbutton(frame, variable=var).grid(
                row=eid + 2, column=0, sticky="nw"
            )
            ttk.Label(frame, text=eid + 1).grid(row=eid + 2, column=1, sticky="nw")
            ttk.Label(frame, text=name).grid(row=eid + 2, column=2, sticky="nw")
            ttk.Label(frame, text=email).grid(row=eid + 2, column=3, sticky="nw")
            self.chbvars.append(var)
        for child in frame.winfo_children():
            child.grid_configure(padx=(0, 10))

    def state(self):
        return list(map((lambda var: var.get()), self.chbvars))


class MainCanvas(tk.Canvas):
    def __init__(self, parent, filename, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.filename = filename
        self.parent = parent
        self._config()
        self._display_records()

    def _config(self):
        self.config(bg="#ababab")

    def _display_records(self):
        employee = EmployeeFrame(self, self.filename)
        employees = employee.employees_data()
        employee.create_record(employees)
        self.create_window(10, 10, anchor="nw", window=employee)


class MainFrame(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)


def main():
    # default excel to parse
    FILENAME = "small.xlsx"

    root = tk.Tk()
    root.title("Payroll")
    root.geometry("400x300")

    # Main frame
    mainframe = MainFrame(root)
    mainframe.pack(fill=tk.BOTH, expand=True, anchor="nw")

    # main canvas
    canvas = MainCanvas(mainframe, FILENAME)
    canvas.pack(fill=tk.BOTH, expand=True)

    # root config
    root.mainloop()


if __name__ == "__main__":
    main()