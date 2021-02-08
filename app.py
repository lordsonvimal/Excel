import tkinter as tk
from tkinter import scrolledtext
from tkinter import filedialog
from tkinter import messagebox
import os, re
import threading

from src.specification.spec_process import Spec

window_width = 500
window_height = 500

exec_height = 50

class App:
    def __init__(self, title):
        self.window = tk.Tk()
        self.window.title(title)
        self.set_window_size()
        self.set_icon()

    def set_window_size(self):
        ws = self.window.winfo_screenwidth()
        hs = self.window.winfo_screenheight()
        x = (ws/2) - (window_width/2)
        y = (hs/2) - (window_height/2)

        self.window.geometry("+%d+%d" % (x, y))
        self.window.resizable(0, 0)

    def set_icon(self):
        dirname = os.path.dirname(__file__)
        self.window.iconbitmap(os.path.join(dirname, "icon.ico"))

    def run(self):
        self.window.mainloop()


class UI:
    def __init__(self, app):
        self.window = app.window
        self.filename = ""
        self.input_file = None
        self.input_spec = None
        self.input_spec_str = tk.StringVar()
        self.input_spec_gen = ""
        self.output_message = None
        self.output_message_lines = 0
        self.create()

    def create(self):
        self.create_browser()
        self.create_specification()
        self.create_message()
        self.create_execute()

    def create_browser(self):
        frame = tk.Frame(self.window, width=window_width, height=window_width-exec_height)
        frame.pack(expand=True, fill=tk.X)

        self.input_file = tk.Label(frame, borderwidth=2, text="Select a template file", fg="#aaaaaa", font=("Calibri", 12))
        self.input_file.grid(sticky=tk.N+tk.S+tk.W, row=0, column=0, padx=5, pady=5)

        file_browse = tk.Button(frame, text="Select Template", command=self.browse, width=35)
        file_browse.grid(row=0, column=1, padx=(0, 5), pady=5, ipady=2)

        tk.Grid.columnconfigure(frame, 0, weight=1)

    def create_specification(self):
        frame = tk.Frame(self.window, width=window_width, height=window_width-exec_height)
        frame.pack(expand=True, fill=tk.X)

        label = tk.Label(frame, borderwidth=2, justify=tk.LEFT, text="Enter specifications to create", fg="#aaaaaa", font=("Calibri", 12))
        label.grid(sticky=tk.N+tk.S+tk.W, row=0, column=0, padx=5, pady=(0, 5), ipady=2)

        self.input_spec = tk.Entry(frame, textvariable=self.input_spec_str, width=38)
        self.input_spec.grid(row=0, column=1, pady=(0, 5), padx=(0, 5))

        tk.Grid.columnconfigure(frame, 0, weight=1)

    def create_message(self):
        self.output_message = scrolledtext.ScrolledText(self.window, state="disabled", font=("Calibri", 12))
        self.output_message.pack(expand=True, fill=tk.BOTH, padx=(5))

    def create_execute(self):
        exec_btn = tk.Button(self.window, text="Execute", command=self.execute)
        exec_btn.pack(expand=True, fill=tk.X, ipadx=10, ipady=10, padx=(4), pady=4)

    def browse(self):
        self.filename = filedialog.askopenfilename(initialdir = "/", title = "Select a Template", filetypes = (("Excel Files", "*.xlsx"),))
        if len(self.filename) > 0:
            wrapped = self.filename[0:35]+"..." if len(self.filename) > 40 else self.filename
            self.input_file.config(text=wrapped)
            self.append_message("Selected File: " + self.filename)
        else:
            self.input_file.config(text="Select a template file")

    def append_message(self, message, lines=1):
        self.output_message_lines += lines
        empty_lines = lines * "\n"
#         line_no = " " * 5 + "[" + str(self.output_message_lines) + "]"
#         out_message = message + line_no + empty_lines if len(message) > 0 else empty_lines
        self.output_message.config(state="normal")
        self.output_message.insert("end", "[INFO] " + message + empty_lines)
        self.output_message.config(state="disabled")
        self.output_message.see(tk.END)

    def validate(self):
        return (len(self.input_spec_str.get()) > 0 and os.path.exists(self.filename))

    def get_validation_message(self):
        if (len(self.input_spec_str.get()) == 0 and not os.path.exists(self.filename)):
            return "Select a valid template in xlsx format and enter specifications to start processing"
        elif not os.path.exists(self.filename):
            return "Select a valid template in xlsx format to start processing"
        return "Enter specification to start processing"

    def popup_execute(self):
        spec_gen = re.sub(r"\s+", "", self.input_spec_str.get().strip().upper(), flags=re.UNICODE)
        message_text = "Do you want to create specifications: " + spec_gen + "?"
        res=messagebox.askquestion("Specifications to be created", message_text)
        if res == "yes":
            self.append_message("Specifications Entered: " + spec_gen)
            self.append_message("Starting Process: ")
            spec = Spec(self.filename, self.input_spec_str.get(), self.append_message)
            threading.Thread(target=spec.process).start()


    def execute(self):
        if self.validate():
            self.popup_execute()
        else:
            self.append_message(self.get_validation_message())


if __name__ == "__main__":
    app = App("SRDM to SDTM Specification")
    ui = UI(app)
    app.run()
