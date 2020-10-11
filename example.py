import tkinter as tk
from tkinter import ttk
from time import sleep


root = tk.Tk()
pb = ttk.Progressbar(root, length=100, orient="horizontal", mode="determinate")
pb.pack(expand=True, fill=tk.BOTH, side=tk.TOP)
pb.config(value=10)
sleep(1)
pb.config(value=20)
sleep(1)
pb.config(value=30)
sleep(1)
root.mainloop()