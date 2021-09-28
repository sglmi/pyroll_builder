import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from ttkwidgets import CheckboxTreeview
import utils


class Navebar(tk.Menu):
    def __init__(self, parent, mainframe, *arg, **kwargs):
        super().__init__(parent, *arg, *kwargs)
        self.parent = parent
        self.mainframe = mainframe
        self.checkall_var = tk.BooleanVar(False)
        # File
        filemenu = tk.Menu(self, tearoff=0)
        filemenu.add_command(label="Open", command=self.openfile)
        filemenu.add_command(label="Save As", command=self.saveas)
        filemenu.add_checkbutton(
            label="Check All", variable=self.checkall_var, command=self.checkall
        )
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
            self.mainframe.statusbar.destroy()
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
        sh = utils.sheet(self.mainframe.filename)
        emps = utils.employees(sh)
        tmp = utils.read_template()
        # Progress
        statusbar = self.mainframe.statusbar
        statusbar.display_progressbar()
        num_pdfs = len(names)
        for i, name in enumerate(names):
            emp = utils.employee(emps, name)
            html = utils.template_to_html(emp, tmp)
            pdf = utils.html_to_pdf(html, f"{directory}/{i}.pdf")
            prc_saved = (i / num_pdfs) * 100
            text = f"pdf saved with name {pdf}"
            # Progressbar
            statusbar.update_progress(prc_saved, text)
            self.parent.update_idletasks()
        # Statusbar update
        statusbar.text.set("All PDFs Saved!")
        statusbar.progressbar.forget()

    def checkall(self):
        eids = [employee.get("id") for employee in self.mainframe.table.employees]
        if self.checkall_var.get():
            for eid in eids:
                self.mainframe.table.tree.change_state(eid, "checked")
        else:
            for eid in eids:
                self.mainframe.table.tree.change_state(eid, "unchecked")

    def sendmail(self):
        if not self.mainframe.table.tree.get_checked():
            messagebox.showerror(
                "Not Selected Email", "Check one or more employee to send mail"
            )
            return
        names = []
        emails = []
        for employee in self.mainframe.table.employees:
            eid, name, email = employee.values()
            for item_id in self.mainframe.table.tree.get_checked():
                if item_id == eid:
                    names.append(name)
                    emails.append(email)

        num_emails = sum(1 for email in emails if email is not None)
        # all emails are None!
        if num_emails <= 0:
            messagebox.showinfo("No Valid Emails", "There are not any valid email!!")
            return
        num_sent = 0
        conn = utils.email_connection()
        filename = self.mainframe.filename
        statusbar = self.mainframe.statusbar
        statusbar.display_progressbar()  # pack progressbar
        sh = utils.sheet(filename)
        emps = utils.employees(sh)
        for name, email in zip(names, emails):
            emp = utils.employee(emps, name)
            extra = {
                "YEAR": emp.get("year"),
                "MONTH": emp.get("month"),
                "NAME": emp.get("name"),
            }
            num_sent += utils.send_mail(conn, filename, email, extra)
            prc_sent = (num_sent / num_emails) * 100
            msg = f"Sending Email To {name}"
            # Progressbar
            statusbar.update_progress(prc_sent, msg)
            self.parent.update_idletasks()
            # statusbar.display_label()
        statusbar.text.set("All emails sent successfuly.")
        statusbar.progressbar.forget()
        conn.quit()

    def preview(self):
        item = self.mainframe.table.tree.focus()
        if item == "":  # no item selected
            messagebox.showinfo(
                "Select an Employee", "Select an employee to show payroll."
            )
            return

        window = tk.Toplevel(self.parent)
        window.title("Payroll")
        self.label = ttk.Label(window)
        for employee in self.mainframe.table.employees:
            eid, name, _ = employee.values()
            if eid == str(item):
                break
        sheet = utils.sheet(self.mainframe.filename)
        emps = utils.employees(sheet)
        emp = utils.employee(emps, name)
        tmp = utils.read_template()
        html = utils.template_to_html(emp, tmp)
        pdf = utils.html_to_pdf(html, "payroll.pdf")
        pix = utils.pdf_to_image(pdf)
        imgdata = utils.get_image_bytes(pix)
        tkimg = tk.PhotoImage(data=imgdata)
        self.label.img = tkimg
        self.label.config(image=self.label.img)
        self.label.pack()


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
        sheet = utils.sheet(filename)
        employees = utils.employees(sheet)
        names = utils.items(employees, column_name="name")
        emails = utils.items(employees, column_name="email")
        ids = range(len(names))

        # populate row num, names and emails on mployee Tree
        for i, name, email in zip(ids, names, emails):
            self.tree.insert("", "end", iid=i, values=(i + 1, name, email))
            employee = {"id": str(i), "name": name, "email": email}
            self.employees.append(employee)


class Statusbar(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.text = tk.StringVar()
        self.text.set("Everything is OK.")
        self.pb_var = tk.IntVar()
        self.label = ttk.Label(self, textvariable=self.text)
        self.label.pack(side=tk.LEFT)
        self.progressbar = ttk.Progressbar(
            self,
            orient=tk.HORIZONTAL,
            length=100,
            maximum=100,
            mode="determinate",
            variable=self.pb_var,
        )

    def display_progressbar(self):
        self.progressbar.pack(
            side=tk.LEFT, anchor="nw", before=self.label, padx=(0, 10)
        )

    def update_progress(self, progress, text):
        self.pb_var.set(progress)
        self.text.set(text)


class MainFrame(ttk.Frame):
    def __init__(self, parent, filename, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.filename = filename
        self.table = EmployeeTable(self, self.filename)
        self.table.pack(fill=tk.BOTH, expand=True, anchor="nw")
        self.table.populate_data(filename)
        self.statusbar = Statusbar(parent)
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X, anchor="w")


def main():
    # default excel to parse
    FILENAME = "small.xlsx"

    root = tk.Tk()
    root.title("Payroll")
    root.geometry("550x400")
    # root.attributes("-topmost", True)
    root.deiconify()
    # root.focus_force()
    # root.update()
    # root.update_idletasks()

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
