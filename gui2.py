import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import fitz
from PIL import Image, ImageTk
import code2


checkbuttons = []
checkvars = []
emails_to_send = []
selected_rows = []
allitems = []


# def save_pdf():

#     filename = filedialog.asksaveasfilename(
#         defaultextension=".pdf", filetypes=(("pdf", "*.pdf"), ("All Files", "*.*"))
#     )
#     if filename != "":
#         dd = fitz.open(pdf_filename)

#         for i in dd:
#             i.insertImage(rect, filename=img_filename, rotate=rotate(direction_val))
#         dd.save(f"{filename}")
#         messagebox.showinfo("File", "file saved")
#     else:
#         messagebox.showerror("File", "choose a file!")


# pdf_filename = filedialog.askopenfilename()


def state():
    global checkvars
    indeces = []
    items = list(map((lambda var: var.get()), checkvars))
    for index, item in enumerate(items):
        if item:
            indeces.append(index)
    return indeces


def test(tree):
    global emails_to_send
    emails_to_send = []
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
    sheet = code2.sheet()
    # list of employees
    employees = code2.employees(sheet)
    row_nums = code2.items(employees, column_name="row")

    for row in row_nums:
        var = tk.BooleanVar()
        checkbutton = ttk.Checkbutton(frame, variable=var)
        checkbutton.grid(sticky="S")
        checkbuttons.append(checkbutton)
        checkvars.append(var)


def select_item(tree):
    global allitems
    itemss = tree.selection()
    allitems = []
    for i in itemss:
        allitems.append(tree.item(i)["values"])
    print(allitems)
    return allitems


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
    sheet = code2.sheet()
    # list of employees
    employees = code2.employees(sheet)
    # insert rows number
    names = code2.items(employees, column_name="name")
    emails = code2.items(employees, column_name="email")
    nums = range(len(names))
    print(emails)
    print(len(emails))
    tree.bind("<ButtonRelease-1>", lambda e: select_item(tree))
    # populate row num, names and emails on mployee Tree
    for num, name, email in zip(nums, names, emails):
        tree.insert("", "end", values=(num, name, email))

    # grid tree and button
    tree.grid()
    send_button = ttk.Button(frame, text="Send", command=lambda: test(tree))
    send_button.grid(row=1)


def pdf_preview(frame):
    global allitems
    print("allitems", allitems)
    sheet = code2.sheet()
    employee_dict = code2.find_right_row(allitems, sheet)

    label_img = ttk.Label(frame)
    pdfpaths = code2.create_pdf(employee_dict, allitems)
    pdfpath = pdfpaths[-1]
    doc_copy = fitz.open(pdfpath)
    page = doc_copy[0]  # first page of the pdf
    pix = page.getPixmap()
    mode = "RGBA" if pix.alpha else "RGB"
    img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
    # img.thumbnail((550, 400))
    tkimg = ImageTk.PhotoImage(img)
    print(label_img)
    print("hi")
    label_img.img = tkimg
    label_img.config(image=label_img.img)
    label_img.grid(row=1, column=0, sticky=(tk.E, tk.W, tk.N, tk.S))


def right_side(frame):
    preview_button = ttk.Button(
        frame, text="preview_button", command=lambda: pdf_preview(frame)
    )
    save_pdf_button = ttk.Button(
        frame, text="save_pdf_button", command=lambda: pdf_preview(frame)
    )
    preview_button.grid(row=0, column=0)
    save_pdf_button.grid(row=0, column=1)


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
    # frame# label.grid()

    # global label_img
    # label_img = ttk.Label(right_frame)

    # display widgets
    checkbox(check_frame)
    left_side(left_frame)
    right_side(right_frame)

    mainframe.pack(fill=tk.BOTH, expand=True)
    mainframe.columnconfigure(1, weight=1)
    mainframe.columnconfigure(2, weight=6)
    mainframe.rowconfigure(0, weight=1)

    for child in mainframe.winfo_children():
        child.grid_configure(padx=5, pady=5)
    root.mainloop()


if __name__ == "__main__":
    main()
