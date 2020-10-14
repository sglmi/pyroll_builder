import code
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from ttkwidgets import CheckboxTreeview


class Navebar(tk.Menu):
    def __init__(self, parent, mainframe, *arg, **kwargs):
        super().__init__(parent, *arg, *kwargs)
        self.parent = parent
        self.mainframe = mainframe
        # File
        filemenu = tk.Menu(self, tearoff=0)
        filemenu.add_command(label="Open", command=self.openfile)
        filemenu.add_command(label="Save As", command=self.saveas)
        self.add_cascade(label="File", menu=filemenu)
        # Help
        payrollmenu = tk.Menu(self, tearoff=0)
        payrollmenu.add_command(label="Send Email", command=self.sendmail)
        payrollmenu.add_command(label="Preview", command=self.preview)
        self.add_cascade(label="Payroll", menu=payrollmenu)
        # Help
        helpmenu = tk.Menu(self, tearoff=0)
        helpmenu.add_command(label="Guide")
        helpmenu.add_command(label="About")
        self.add_cascade(label="Help", menu=helpmenu)

    def openfile(self):
        filename = filedialog.askopenfilename(
            filetypes=(
                ("Excel File", "*.xlsx"),
                ("Excel File (old)", "xls"),
                ("All files", "*.*"),
            )
        )
        if filename:
            self.mainframe.destroy()
            self.mainframe = MainFrame(self.parent, filename)
            self.mainframe.pack(fill=tk.BOTH, expand=True, anchor="nw")

    def saveas(self):
        directory = filedialog.askdirectory()
        if directory == ():  # if user click on cancel
            return

        tree = self.mainframe.table.tree
        if tree.get_checked() == []:  # if any employee not selected.
            messagebox.showinfo(
                "Employee Not Selected", "Select at least one employee to make pdf"
            )
            return
        employess = self.mainframe.table.employees
        names = []
        for employee in employess:
            eid, name, _ = employee.values()
            if eid in tree.get_checked():
                names.append(name)
        sh = code.sheet(self.mainframe.filename)
        emps = code.employees(sh)
        tmp = code.read_template()
        for name in names:
            emp = code.employee(emps, name)
            html = code.template_to_html(emp, tmp)
            pdf = code.html_to_pdf(html, f"{directory}/{name}.pdf")
            print(f"pdf created successfuly {pdf}")

    def sendmail(self):
        focus = self.mainframe.table.tree.focus()
        print(focus)

        # es_id = self.mainframe.employee_table.checkeds()
        # names, emails = self.mainframe.employee.extract_name_email(es_id)
        # conn = code.email_connection()
        # filename = self.mainframe.filename
        # for name, email in zip(names, emails):
        #     code.send_mail(conn, filename, name, email)
        #     print("email send to ", email, "successfuly.")
        pass

    def preview(self, eid=1):
        # window = tk.Toplevel(self.parent)
        # window.title("Payroll")
        # self.label = ttk.Label(window)
        # name = self.mainframe.employee.extract_name(eid)
        # sheet = code.sheet(self.mainframe.filename)
        # emps = code.employees(sheet)
        # emp = code.employee(emps, name)
        # tmp = code.read_template()
        # html = code.create_payroll_html(emp, tmp)
        # pdf = code.html_to_pdf(html, "payroll.pdf")
        # pix = code.pdf_to_image(pdf)
        # imgdata = code.get_image_bytes(pix)
        # tkimg = tk.PhotoImage(data=imgdata)
        # self.label.img = tkimg
        # self.label.config(image=self.label.img)
        # self.label.pack()
        pass


class EmployeeTable(ttk.Frame):
    def __init__(self, parent, filename, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.filename = filename
        # tree and pack
        self.tree = CheckboxTreeview(self)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # tree option config
        self._config()
        self._headings()
        self._scrollbar(parent)

        # store all employees' id
        self.employees = []

        # tree bind, change color when clicked on row
        self.tree.bind("<<TreeviewSelect>>", self.on_select_item)

        # last selected item
        self.last_selected_item = 0

    def _config(self):
        self.tree.configure(
            columns=("row", "name", "email"),
        )
        self.tree.column("#0", minwidth=20, width=20)
        self.tree.column("row", minwidth=10, width=10)
        self.tree.column("name", minwidth=40, width=100)

    def _headings(self):
        self.tree.heading("#0", text="Check")
        self.tree.heading("row", text="Row")
        self.tree.heading("name", text="Name")
        self.tree.heading("email", text="email")

    def _scrollbar(self, parent):
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, anchor="nw", fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)

    def on_select_item(self, event):
        # Remove highlight from last selected item
        self.tree.tag_del(self.last_selected_item, "highlight")
        item = self.tree.focus()
        # Add highlight tag to focused item
        self.tree.tag_add(item, "highlight")
        self.tree.tag_configure("highlight", background="orange")
        self.last_selected_item = item

    def populate_data(self, filename):
        sheet = code.sheet(filename)
        employees = code.employees(sheet)
        names = code.items(employees, column_name="name")
        emails = code.items(employees, column_name="email")
        ids = range(len(names))

        # populate row num, names and emails on mployee Tree
        for i, name, email in zip(ids, names, emails):
            self.tree.insert("", "end", iid=i, values=(i + 1, name, email))

            employee = {"id": str(i), "name": name, "email": email}
            self.employees.append(employee)


class MainFrame(ttk.Frame):
    def __init__(self, parent, filename, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.filename = filename
        self.table = EmployeeTable(self, self.filename)
        self.table.pack(fill=tk.BOTH, expand=True, anchor="nw")
        self.table.populate_data(filename)


def main():
    # default excel to parse
    FILENAME = "small.xlsx"

    root = tk.Tk()
    root.title("Payroll")
    root.geometry("550x400")

    # Main frame
    mainframe = MainFrame(root, FILENAME)
    mainframe.pack(fill=tk.BOTH, expand=True, anchor="nw")

    # self
    menubar = Navebar(root, mainframe)
    root.config(menu=menubar)

    # root config
    root.mainloop()


if __name__ == "__main__":
    main()