import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from threading import Thread
import Custom as Custom
from variables import lattest_report_path, new_file_name, financial_new_file_name
import os
import subprocess
from variables import log_file_path

class ModernSpinner:
    def __init__(self, canvas):
        self.canvas = canvas
        self.size = 30
        self.color = "#87CEEB"
        self.spinner_id = self.canvas.create_arc(
            10, 10, self.size + 10, self.size + 10,
            start=0, extent=60, outline=self.color, width=3, style=tk.ARC
        )
        self.angle = 0
        self.animate()

    def animate(self):
        self.angle = (self.angle + 15) % 360
        self.canvas.delete(self.spinner_id)
        self.spinner_id = self.canvas.create_arc(
            10, 10, self.size + 10, self.size + 10,
            start=self.angle, extent=60, outline=self.color, width=3, style=tk.ARC
        )
        self.canvas.after(40, self.animate)

class ExcelUploaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Report Generator")

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        window_width = int(screen_width * 0.5)
        window_height = int(screen_height * 0.3)

        x_position = int((screen_width - window_width) / 2)
        y_position = int((screen_height - window_height) / 2)

        self.root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        self.root.configure(bg='white')

        self.style = ttk.Style()
        self.style.configure('TButton', background='white', padding=10)
        self.style.configure('TLabel', background='white', foreground='black', font=("Arial", 12))
        self.style.configure('TCheckbutton', background='white', padding=5)
        self.style.map('TCheckbutton', background=[('selected', '#87CEEB'), ('!selected', 'white')])

        self.start_screen_frame = tk.Frame(root, bg='white')
        self.start_screen_frame.pack(expand=True, fill=tk.BOTH)

        self.start_content_frame = tk.Frame(self.start_screen_frame, bg='white')
        self.start_content_frame.pack(expand=True, pady=(20, 10))

        self.start_message_label = ttk.Label(self.start_content_frame, text="Report Generator", style='TLabel', wraplength=window_width * 0.8)
        self.start_message_label.pack(pady=(0, 10), padx=10)

        self.checkbox_frame = tk.Frame(self.start_content_frame, bg='white')
        self.checkbox_frame.pack(pady=(10, 20))

        self.do_financial_var = tk.BooleanVar()
        self.do_dashboard_var = tk.BooleanVar()
        self.analyze_descriptions_only_var = tk.BooleanVar()

        self.financial_checkbox = ttk.Checkbutton(self.checkbox_frame, text="Financial", variable=self.do_financial_var, style='TCheckbutton', command=self.toggle_checkboxes)
        self.dashboard_checkbox = ttk.Checkbutton(self.checkbox_frame, text="Dashboard", variable=self.do_dashboard_var, style='TCheckbutton', command=self.toggle_checkboxes)
        self.analyze_checkbox = ttk.Checkbutton(self.checkbox_frame, text="Analyze Descriptions Only", variable=self.analyze_descriptions_only_var, style='TCheckbutton', command=self.toggle_checkboxes)

        self.financial_checkbox.pack(side=tk.LEFT, padx=10)
        self.dashboard_checkbox.pack(side=tk.LEFT, padx=10)
        self.analyze_checkbox.pack(side=tk.LEFT, padx=10)

        self.upload_button = ttk.Button(self.start_content_frame, text="Upload Files", command=self.upload_files, style='TButton')
        self.upload_button.pack(pady=20)

        self.loading_frame = tk.Frame(root, bg='white')
        self.loading_frame.pack_forget()
        self.loading_label = ttk.Label(self.loading_frame, text="Uploading files...", style='TLabel')
        self.loading_label.pack()

        self.loading_canvas = tk.Canvas(self.loading_frame, width=50, height=50, bg="white", bd=0, highlightthickness=0)
        self.loading_canvas.pack()

        self.spinner = ModernSpinner(self.loading_canvas)

        self.end_screen_frame = tk.Frame(root, bg='white')
        self.end_screen_frame.pack_forget()

        self.end_content_frame = tk.Frame(self.end_screen_frame, bg='white')
        self.end_content_frame.pack(expand=True, pady=(20, 10))

        self.end_message_label = ttk.Label(self.end_content_frame, text="Report Generated Successfully!", style='TLabel', wraplength=window_width * 0.8)
        self.end_message_label.pack(pady=(0, 10), padx=10)

        self.buttons_frame = tk.Frame(self.end_content_frame, bg='white')
        self.buttons_frame.pack(pady=20)

        self.new_report_button = ttk.Button(self.buttons_frame, text="Generate a new report", command=self.go_to_upload_screen, style='TButton')
        self.quit_button = ttk.Button(self.buttons_frame, text="Quit", command=self.quit_application, style='TButton')
        self.open_result_log_button = ttk.Button(self.buttons_frame, text="Open Analysis", command=self.open_result_log, style='TButton')
        self.open_new_file_button = ttk.Button(self.buttons_frame, text="Open Dashboard", command=self.open_new_file, style='TButton')
        self.open_financial_button = ttk.Button(self.buttons_frame, text="Open Financial", command=self.open_financial_file, style='TButton')

        self.new_report_button.grid(row=0, column=0, padx=5)
        self.open_result_log_button.grid(row=0, column=1, padx=5)
        self.open_new_file_button.grid(row=0, column=2, padx=5)
        self.open_financial_button.grid(row=0, column=3, padx=5)
        self.quit_button.grid(row=0, column=4, padx=5)

        self.root.protocol("WM_DELETE_WINDOW", self.quit_application)

    def toggle_checkboxes(self):
        if self.analyze_descriptions_only_var.get():
            # If Analyze Descriptions Only is checked, disable and uncheck other checkboxes
            self.do_financial_var.set(False)
            self.do_dashboard_var.set(False)
            self.financial_checkbox.state(['disabled'])
            self.dashboard_checkbox.state(['disabled'])
        else:
            # If Analyze Descriptions Only is unchecked, enable other checkboxes
            self.financial_checkbox.state(['!disabled'])
            self.dashboard_checkbox.state(['!disabled'])

        if self.do_financial_var.get() or self.do_dashboard_var.get():
            # If either Financial or Dashboard is checked, disable and uncheck Analyze Descriptions Only
            self.analyze_descriptions_only_var.set(False)
            self.analyze_checkbox.state(['disabled'])
        else:
            # Enable Analyze Descriptions Only if neither Financial nor Dashboard is checked
            self.analyze_checkbox.state(['!disabled'])

    def upload_files(self):
        self.start_screen_frame.pack_forget()
        self.loading_frame.pack(expand=True)

        thread = Thread(target=self.process_files)
        thread.start()

    def process_files(self):
        file_paths = filedialog.askopenfilenames(
            title="Select Excel Files",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )

        if file_paths:
            self.root.after(0, lambda: self.loading_label.config(text="Generating the report..."))
            self.root.update_idletasks()

            do_financial = self.do_financial_var.get()
            do_dashboard = self.do_dashboard_var.get()
            analyzed_descriptions_only = self.analyze_descriptions_only_var.get()

            try:
                Custom.excecute(file_paths, (do_financial, do_dashboard, analyzed_descriptions_only))
                self.root.after(0, self.show_end_screen)
            except Exception as e:
                print(f"Error storing the files as csv: {e}")
                messagebox.showerror("Error", f"An error occurred: {e}")
                self.root.after(0, self.reset_ui)
            finally:
                self.root.after(0, self.reset_ui(show_upload=False))
        else:
            self.root.after(0, self.reset_ui)

    def reset_ui(self, show_upload=True):
        self.loading_frame.pack_forget()
        if show_upload:
            self.start_screen_frame.pack(expand=True, fill=tk.BOTH)
        self.root.update_idletasks()

    def show_end_screen(self):
        self.start_screen_frame.pack_forget()
        self.root.update_idletasks()
        self.end_screen_frame.pack(expand=True, fill=tk.BOTH)

    def go_to_upload_screen(self):
        self.end_screen_frame.pack_forget()
        self.start_screen_frame.pack(expand=True, fill=tk.BOTH)

    def quit_application(self):
        self.root.destroy()

    def open_result_log(self):
        log_path = os.path.abspath(log_file_path).replace('\\', '/')
        if os.path.exists(log_path):
            if os.name == 'nt':  # For Windows
                os.startfile(log_path)
            else:  # For Unix-based systems
                subprocess.call(['open', log_path])
        else:
            messagebox.showerror("Error", "Log file not found.")

    def open_new_file(self):
        directory = os.path.abspath(lattest_report_path + "/Dashboard").replace('\\', '/')
        prefix = new_file_name

        found_file = None
        for file_name in os.listdir(directory):
            if file_name.startswith(prefix) and file_name.endswith('.xlsx'):
                found_file = os.path.join(directory, file_name)
                break

        if found_file and os.path.exists(found_file):
            if os.name == 'nt':  # For Windows
                os.startfile(found_file)
            else:  # For Unix-based systems
                subprocess.call(['open', found_file])
        else:
            messagebox.showerror("Error", "Dashboard report not found.")

    def open_financial_file(self):
        directory = os.path.abspath(lattest_report_path + "/Financial").replace('\\', '/')
        prefix = financial_new_file_name
        
        found_file = None
        for file_name in os.listdir(directory):
            if file_name.startswith(prefix) and file_name.endswith('.xlsx'):
                found_file = os.path.join(directory, file_name)
                break

        if found_file and os.path.exists(found_file):
            if os.name == 'nt':  # For Windows
                os.startfile(found_file)
            else:  # For Unix-based systems
                subprocess.call(['open', found_file])
        else:
            messagebox.showerror("Error", "Financial report not found.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelUploaderApp(root)
    root.mainloop()
