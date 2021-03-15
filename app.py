import tkinter as tk
from tkinter import ttk
from tkinter import scrolledtext
from tkinter import filedialog
from tkinter import messagebox
import os, re
import threading
from string import digits
from platform import system

from src.specification.spec_process import Spec

window_width = 500
window_height = 500

exec_height = 50

sys_platform = system()

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
        logo = "icon.ico"
        dir = os.path.dirname(__file__)
#         img = tk.Image("photo", file="icon.gif")
        if sys_platform == "Darwin":
            logo = "icon.icns"
#         self.window.iconphoto(True, img)
#         self.window.tk.call('wm','iconphoto', self.window._w, img)
        self.window.iconbitmap(logo)

    def run(self):
        self.window.mainloop()

class UI:
    def __init__(self, app):
        self.window = app.window
        self.filename = ""
        self.input_file = None
        self.input_spec = None
        self.input_spec_str = tk.StringVar()
        self.input_spec_gen_dir = ""
        self.input_spec_gen = []
        self.output_message = None
        self.output_message_lines = 0
        self.create()

    def create(self):
        self.create_browser()
        self.create_specification_browser()
        self.create_user_defined_domain()
        self.create_message()
        self.create_execute()

    def create_browser(self):
        frame = tk.Frame(self.window, width=window_width, height=window_width-exec_height)
        frame.pack(expand=True, fill=tk.X)

        self.input_file = tk.Label(frame, borderwidth=2, text="Select a template file*", fg="#aaaaaa", font=("Calibri", 12))
        self.input_file.grid(sticky=tk.N+tk.S+tk.W, row=0, column=0, padx=5, pady=5)

        self.browse_template = ttk.Button(frame, text="Select template file", command=self.browse, width=36)
        self.browse_template.grid(row=0, column=1, padx=(0, 5), pady=5, ipady=2)

        tk.Grid.columnconfigure(frame, 0, weight=1)

    def create_specification_browser(self):
        frame = tk.Frame(self.window, width=window_width, height=window_width-exec_height)
        frame.pack(expand=True, fill=tk.X)

        label = tk.Label(frame, borderwidth=2, justify=tk.LEFT, text="Select specifications files*", fg="#aaaaaa", font=("Calibri", 12))
        label.grid(sticky=tk.N+tk.S+tk.W, row=0, column=0, padx=5, pady=(0, 5), ipady=2)

        self.browse_spec = ttk.Button(frame, text="Select specification file(s)", command=self.browse_spec, width=36)
        self.browse_spec.grid(row=0, column=1, padx=(0, 5), pady=5, ipady=2)

        tk.Grid.columnconfigure(frame, 0, weight=1)

    def create_user_defined_domain(self):
        frame = tk.Frame(self.window, width=window_width, height=window_width-exec_height)
        frame.pack(expand=True, fill=tk.X)

        label = tk.Label(frame, borderwidth=2, justify=tk.LEFT, text="Enter additional domain (Optional)", fg="#aaaaaa", font=("Calibri", 12))
        label.grid(sticky=tk.N+tk.S+tk.W, row=0, column=0, padx=5, pady=(0, 5), ipady=2)

        self.input_domain = ttk.Entry(frame, width=38)
        self.input_domain.grid(row=0, column=1, padx=(0, 5), pady=5, ipady=2, ipadx=2)

        tk.Grid.columnconfigure(frame, 0, weight=1)

    def browse_spec(self):
        files = filedialog.askopenfilenames(title="Choose specification files", filetypes=(("Excel Files", "*.xlsx"),))
        print(files)
        if len(files) > 0:
            wrapped = ",".join(files)[0:35]+"..." if len(",".join(files)) > 40 else ",".join(files)
            self.browse_spec.config(text=wrapped)
            self.input_spec_gen_dir = os.path.dirname(files[0])
            file_names = [os.path.basename(f).split(".xlsx")[0] for f in files]
            self.input_spec_gen = [f.split("_")[0].rstrip(digits) for f in file_names]
            self.append_message("Selected Specifications")
            for spec in self.input_spec_gen:
                self.append_message(spec)
        else:
            self.browse_spec.config(text="Select specification file(s)")
            self.input_spec_gen_dir = ""
            self.input_spec_gen = []

    def create_message(self):
        self.output_message = scrolledtext.ScrolledText(self.window, state="disabled", font=("Calibri", 12), relief="solid", borderwidth=1)
        self.output_message.pack(expand=True, fill=tk.BOTH, padx=(5))

    def create_execute(self):
        exec_btn = ttk.Button(self.window, text="Execute", command=self.execute)
        exec_btn.pack(expand=True, fill=tk.X, ipadx=10, ipady=10, padx=(4), pady=4)

    def browse(self):
        self.filename = filedialog.askopenfilename(initialdir="/", title="Select a Template", filetypes=(("Excel Files", "*.xlsx"),))
        if len(self.filename) > 0:
            wrapped = self.filename[0:35]+"..." if len(self.filename) > 40 else self.filename
            self.browse_template.config(text=wrapped)
            self.append_message("Selected File: " + self.filename)
        else:
            self.browse_template.config(text="Select Template")

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
        return len(self.input_spec_gen) > 0 and os.path.exists(self.filename)

    def get_validation_message(self):
        if (len(self.input_spec_gen) == 0 and not os.path.exists(self.filename)):
            return "Select a valid template in xlsx format and enter specifications to start processing"
        elif not os.path.exists(self.filename):
            return "Select a valid template in xlsx format to start processing"
        return "Select a specification file to start processing"

    def popup_execute(self):
#         spec_gen = re.sub(r"\s+", "", self.input_spec_str.get().strip().upper(), flags=re.UNICODE)
        message_text = "Do you want to create specifications: " + ",".join(self.input_spec_gen) + "?"
        res=messagebox.askquestion("Specifications to be created", message_text)
        if res == "yes":
            self.append_message("Starting Process: ")
            domains = self.input_domain.get()
            additional_domain = domains.split(",") if len(domains) > 0 else []
            self.append_message(str(additional_domain))
            app = AppThread(self.filename, self.input_spec_gen, self.input_spec_gen_dir, additional_domain, self.append_message)
            threading.Thread(target=app.process).start()

    def execute(self):
        if self.validate():
            self.popup_execute()
        else:
            self.append_message(self.get_validation_message())

class AppThread:
    def __init__(self, file_name, specifications, spec_gen_dir, additional_domain, append_message):
        self.template_name = file_name
        self.specifications = specifications
        self.spec_gen_dir = spec_gen_dir
        self.additional_domain = additional_domain
        self.append_message = append_message

    def process(self):
        threads = []
        for spec_gen in self.specifications:
            spec = Spec(self.template_name, spec_gen, self.spec_gen_dir, self.additional_domain, self.append_message)
            thread = threading.Thread(target=spec.process)
            thread.start()
            threads.append(thread)
        for t in threads:
            t.join()


if __name__ == "__main__":
    app = App("SRDM to SDTM Specification")
    ui = UI(app)
    app.run()
