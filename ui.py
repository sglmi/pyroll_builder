import code
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import time


class Employee:
    def __init__(self, filename):
        self.filename = filename

    def employees(self):
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

    def extract_name_email(self, esid):
        emails = []
        names = []
        data = self.employees()
        for item in data:
            eid, name, email = item.values()
            if eid in esid:
                names.append(name)
                emails.append(email)

        return names, emails


class Navebar(ttk.Frame):
    def __init__(self, parent, mainframe, *arg, **kwargs):
        super().__init__(parent, *arg, *kwargs)
        self.parent = parent
        self.mainframe = mainframe
        # self.mainframe = mainframe
        self.filename = ""
        menubar = tk.Menu()
        # File
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Open", command=self.openfile)
        filemenu.add_command(label="Save As", command=self.saveas)
        menubar.add_cascade(label="File", menu=filemenu)
        # Help
        payrollmenu = tk.Menu(menubar, tearoff=0)
        payrollmenu.add_command(label="Send Email", command=self.sendmail)
        payrollmenu.add_command(label="Preview", command=self.preview)
        menubar.add_cascade(label="Payroll", menu=payrollmenu)
        # Help
        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="Guide")
        helpmenu.add_command(label="About")
        menubar.add_cascade(label="Help", menu=helpmenu)
        self.parent.config(menu=menubar)

    def openfile(self):
        pass

    def saveas(self):
        pass

    def sendmail(self):
        es_id = self.mainframe.employee_table.checkeds()
        names, emails = self.mainframe.employee.extract_name_email(es_id)
        conn = code.email_connection()
        filename = self.mainframe.filename
        for name, email in zip(names, emails):
            code.send_mail(conn, filename, name, email)
            print("email send to ", email, "successfuly.")

    def preview(self):
        pass


class Table(tk.Canvas):
    def __init__(self, parent, filename, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.filename = filename
        self.rowframe = ttk.Frame(self, relief="solid", padding=10)
        self.rowframe.pack()
        self.check_all_var = tk.BooleanVar()
        self.chbsvar = []

    def header(self, *args, **kwargs):
        ttk.Checkbutton(
            self.rowframe, variable=self.check_all_var, command=self.checkall
        ).grid(row=0, column=0)
        for index, title in enumerate(args, 1):
            ttk.Label(self.rowframe, text=title).grid(row=0, column=index, padx=30)

    def insert_row(self, eid, name, email):
        var = tk.BooleanVar()
        ttk.Checkbutton(self.rowframe, variable=var).grid(row=eid, column=0)
        ttk.Label(self.rowframe, text=eid).grid(row=eid, column=1)
        ttk.Label(self.rowframe, text=name).grid(row=eid, column=2)
        ttk.Label(self.rowframe, text=email).grid(row=eid, column=3)
        self.chbsvar.append(var)

    def populate(self, data):
        for item in data:
            eid, name, email = item.values()
            self.insert_row(eid + 1, name, email)

    def checkall(self):
        flag = self.check_all_var.get()
        list(map(lambda x: x.set(flag), self.chbsvar))

    def state(self):
        return list(map((lambda var: var.get()), self.chbsvar))

    def checkeds(self):
        return [index for index, item in enumerate(self.state()) if item]


class MainFrame(ttk.Frame):
    def __init__(self, parent, filename, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        # Mainmenu
        self.filename = filename
        self.employee = Employee(filename)
        self.employee_table = Table(parent, filename)
        self.employee_table.header("Row", "Name", "Email")
        self.data = self.employee.employees()
        self.employee_table.populate(self.data)
        self.employee_table.pack()
        Navebar(parent, self)


def main():
    # default excel to parse
    FILENAME = "small.xlsx"

    root = tk.Tk()
    root.title("Payroll")
    root.geometry("550x400")

    # Main frame
    mainframe = MainFrame(root, FILENAME)
    mainframe.pack(fill=tk.BOTH, expand=True, anchor="nw")

    # root config
    root.mainloop()


if __name__ == "__main__":
    main()