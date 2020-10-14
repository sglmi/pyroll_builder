# -*- coding: utf-8 -*-

# Copyright (c) Juliette Monsel 2017
# For license see LICENSE

from ttkwidgets import CheckboxTreeview
import tkinter as tk
from tkinter import ttk

root = tk.Tk()

tree = ttk.Treeview(root)
tree.pack()

tree.insert("", "end", "1", text="1")
tree.insert("1", "end", "11", text="11")
tree.insert("1", "end", "12", text="12")
tree.insert("11", "end", "111", text="111")
tree.insert("", "end", "2", text="2")

root.mainloop()
