import tkinter as tk
from tkinter import scrolledtext
from src.excel import Excel

window = tk.Tk()

window_width = 500
window_height = 500

exec_height = 50

ws = window.winfo_screenwidth()
hs = window.winfo_screenheight()
x = (ws/2) - (window_width/2)
y = (hs/2) - (window_height/2)

window.geometry("+%d+%d" % (x, y))
window.minsize(window_width, window_height)
window.title("SRDM to SDTM")
window.resizable(0, 0)

inputs_frame = tk.Frame(window, width=window_width, height=window_width-exec_height)
inputs_frame.pack(expand=True, fill=tk.X)

files_textbox = tk.Text(inputs_frame, height=2, state="disabled")
files_textbox.grid(row=0, column=0, padx=5, pady=5)

files_browse = tk.Button(inputs_frame, height=1, text="Browse Files", command=None)
files_browse.grid(row=0, column=1, padx=(0, 5), pady=5)

message_frame = tk.Frame(window, width=window_width, height=window_width-exec_height-30)
message_frame.pack(expand=True, fill=tk.X)

message_box = scrolledtext.ScrolledText(message_frame, height=30, state="disabled")
message_box.pack(expand=True, fill=tk.X, padx=(5))

button_group_frame = tk.Frame(master=window, width=window_width, height=exec_height)
button_group_frame.pack(side=tk.BOTTOM, expand=True, fill=tk.X)

button1 = tk.Button(button_group_frame, text="Execute", command=None)
button1.pack(expand=True, fill=tk.X, ipadx=10, ipady=10, padx=(5))

window.mainloop()
