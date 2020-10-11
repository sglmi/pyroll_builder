import code
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import time


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


class ProgressbarToplevel(tk.Toplevel):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.pb_text = ttk.Label(self)
        self.pb_text.pack()
        self.progressbar_var = tk.IntVar()
        self.progressbar = ttk.Progressbar(
            self,
            orient=tk.HORIZONTAL,
            length=200,
            maximum=100,
            mode="determinate",
            variable=self.progressbar_var,
        )
        self.progressbar.pack()


class PreviewToplevel(tk.Toplevel):
    def __init__(self, *arg, **kwargs):
        super().__init__(*arg, **kwargs)
        sheet = code.sheet()
        self.employees = code.employees(sheet)
        self.label = ttk.Label(self)

    def payroll_image(self, name):
        employee = code.employee(self.employees, name)
        template = code.read_template()
        payroll_html = code.create_payroll_html(employee, template)
        payroll_pdf = code.html_to_pdf(payroll_html, "payroll.pdf")
        pix = code.pdf_to_image(payroll_pdf)
        return pix

    def preview(self, name):
        # preview in tkinter
        pix = self.payroll_image(name)
        imgdata = code.get_image_bytes(pix)
        tkimg = tk.PhotoImage(data=imgdata)
        self.label.img = tkimg
        self.label.config(image=self.label.img)
        self.label.pack()


class EmployeeTree(ttk.Treeview):
    def __init__(self, parent, *arg, **kwargs):
        super().__init__(parent, *arg, **kwargs)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
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
        # self.columnconfigure(0, weight=1)
        # self.rowconfigure(0, weight=1)

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
        # self.columnconfigure(0, weight=1)
        # self.rowconfigure(0, weight=1)

    def number_of_employees(self):
        return len(self.employee_tree.employees_id)


class MainFrame(tk.Frame):
    def __init__(self, parent, *arg, **kwargs):
        super().__init__(parent, *arg, **kwargs)
        self.parent = parent
        self["relief"] = "sunken"
        # Treeview
        self.employee_frame = EmployeeFrame(self)
        self.employee_frame.grid(row=0, column=1, sticky=(tk.W, tk.S, tk.N))

        # Checkbuttons
        self.checkframe = CheckFrame(self)
        self.checkframe.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.W))
        self.checkframe.create_checkbuttons(self.employee_frame.number_of_employees())

        # Actions
        self.select_all_var = tk.BooleanVar()
        ttk.Checkbutton(
            self,
            text="Select All",
            variable=self.select_all_var,
            command=self.select_all,
        ).grid(row=1)
        ttk.Button(self, text="Send Email", command=self.send_mail).grid(
            row=1, column=1
        )
        ttk.Button(self, text="Preview", command=self.preview).grid(row=1, column=2)

        self.columnconfigure(0, weight=0)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=10)

    def checked_employees(self):
        employees_data = []
        for index in self.checkframe.checkeds():
            employees_data.append(self.employee_frame.employee_tree.item(index))
        return employees_data

    def send_mail(self):
        employees = {}
        for employee in self.checked_employees():
            name = employee.get("values")[1]
            email = employee.get("values")[2]
            employees[name] = email

        if employees.values():
            number_of_emails = sum(1 for value in employees.values() if value != "None")
            number_of_sent = 0
            # Progressbar
            pb_toplevel = ProgressbarToplevel(self)

            conn = code.email_connection()
            for name, email in employees.items():
                number_of_sent += code.send_mail(conn, name, email)
                precent_of_sent = (number_of_sent / number_of_emails) * 100
                text = f"{number_of_sent} / {number_of_emails}\n Send Email To {name}"
                pb_toplevel.pb_text.config(text=text)
                pb_toplevel.progressbar_var.set(precent_of_sent)
                self.parent.update_idletasks()
            # destroy pb toplevel
            time.sleep(2)
            pb_toplevel.destroy()
            conn.quit()
        else:
            messagebox.showerror("Choice Employee", "Choice at least one employee.")

    def preview(self):
        emp_id = self.employee_frame.employee_tree.focus()
        if emp_id != "":
            toplevel = PreviewToplevel()
            emp_values = self.employee_frame.employee_tree.item(emp_id).get("values")
            emp_name = emp_values[1]
            print(emp_id, emp_name)
            toplevel.preview(emp_name)
        else:
            messagebox.showerror(
                "Not Selected an Employee",
                "Select one employee to preview the payroll.",
            )

    def select_all(self):
        print(self.select_all_var.get())
        if self.select_all_var.get():
            for var in self.checkframe.checkvars:
                var.set(True)
        else:
            for var in self.checkframe.checkvars:
                var.set(False)


def main():
    root = tk.Tk()
    root.title("Payroll")
    # ++++++++++++++++++
    menubar = Menubar(root)
    mainframe = MainFrame(root)
    mainframe.grid(row=0, column=0, sticky=(tk.N, tk.E, tk.S, tk.W))
    root.config(menu=menubar)
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    # ++++++++++++++++++
    root.mainloop()


if __name__ == "__main__":
    main()
