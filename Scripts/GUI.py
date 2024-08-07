import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from threading import Thread
import Custom as Custom
from variables import lattest_report_path, new_file_name
import os
import subprocess
from variables import log_file_path

class ModernSpinner:
    def __init__(self, canvas):
        self.canvas = canvas
        # Size of the spinner
        self.size = 30
        # Light blue color
        self.color = "#87CEEB"
        # Create the spinner arc
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
        self.canvas.after(40, self.animate)  # Faster animation

class ExcelUploaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Report Generator")

        # Set the window size to resemble a phone (height much greater than width)
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Define the window width as a smaller fixed size and a larger height
        window_width = int(screen_width * 0.4)  # Narrower width
        window_height = int(screen_height * 0.3)  # Taller height

        # Center the window
        x_position = int((screen_width - window_width) / 2)
        y_position = int((screen_height - window_height) / 2)

        self.root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

        # Set the background color of the root window to white
        self.root.configure(bg='white')

        # Create a style for the buttons and labels
        self.style = ttk.Style()
        self.style.configure('TButton', background='white')
        self.style.configure('TLabel', background='white', foreground='black', font=("Arial", 12))

        # Create a frame for the start screen
        self.start_screen_frame = tk.Frame(root, bg='white')
        self.start_screen_frame.pack(expand=True, fill=tk.BOTH)

        # Create a frame to hold the message and button
        self.start_content_frame = tk.Frame(self.start_screen_frame, bg='white')
        self.start_content_frame.pack(expand=True, pady=(20, 10))

        # Create and place the "Report Generator" label with the common style
        self.start_message_label = ttk.Label(self.start_content_frame, text="Report Generator", style='TLabel', wraplength=window_width * 0.8)
        self.start_message_label.pack(pady=(0, 10), padx=10)

        # Create and place the "Upload Files" button with white background
        self.upload_button = ttk.Button(self.start_content_frame, text="Upload Files", command=self.upload_files, style='TButton')
        self.upload_button.pack(pady=20)

        # Create the loading message and wheel (initially hidden)
        self.loading_frame = tk.Frame(root, bg='white')
        self.loading_frame.pack_forget()  # Initially hidden
        self.loading_label = ttk.Label(self.loading_frame, text="Uploading files...", style='TLabel')
        self.loading_label.pack()

        self.loading_canvas = tk.Canvas(self.loading_frame, width=50, height=50, bg="white", bd=0, highlightthickness=0)
        self.loading_canvas.pack()

        self.spinner = ModernSpinner(self.loading_canvas)

        # Create the end screen frame (initially hidden)
        self.end_screen_frame = tk.Frame(root, bg='white')
        self.end_screen_frame.pack_forget()  # Initially hidden

        # Create a frame for the end message and buttons
        self.end_content_frame = tk.Frame(self.end_screen_frame, bg='white')
        self.end_content_frame.pack(expand=True, pady=(20, 10))

        # Create and place the end message label with the common style
        self.end_message_label = ttk.Label(self.end_content_frame, text="Report Generated Successfully!", style='TLabel', wraplength=window_width * 0.8)
        self.end_message_label.pack(pady=(0, 10), padx=10)

        # Create a frame for buttons to be placed in the center
        self.buttons_frame = tk.Frame(self.end_content_frame, bg='white')
        self.buttons_frame.pack(pady=20)

        # Create and place buttons in a row within the buttons frame with white background
        self.new_report_button = ttk.Button(self.buttons_frame, text="Generate a new report", command=self.go_to_upload_screen, style='TButton')
        self.quit_button = ttk.Button(self.buttons_frame, text="Quit", command=self.quit_application, style='TButton')
        self.open_result_log_button = ttk.Button(self.buttons_frame, text="Open result.log", command=self.open_result_log, style='TButton')
        self.open_new_file_button = ttk.Button(self.buttons_frame, text="Open Report", command=self.open_new_file, style='TButton')

        # Arrange buttons in a row
        self.new_report_button.grid(row=0, column=0, padx=5)
        self.open_result_log_button.grid(row=0, column=1, padx=5)
        self.open_new_file_button.grid(row=0, column=2, padx=5)
        self.quit_button.grid(row=0, column=3, padx=5)

        # Set the protocol for the window close (X button)
        self.root.protocol("WM_DELETE_WINDOW", self.quit_application)

    def upload_files(self):
        # Hide the upload button and show the loading frame
        self.start_screen_frame.pack_forget()
        self.loading_frame.pack(expand=True)

        # Start a separate thread for file selection and processing
        thread = Thread(target=self.process_files)
        thread.start()

    def process_files(self):
        # Open the file dialog and allow multiple file selection
        file_paths = filedialog.askopenfilenames(
            title="Select Excel Files",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )

        if file_paths:
            # Update the message to "Generating the report" after storing files
            self.root.after(0, lambda: self.loading_label.config(text="Generating the report..."))
            self.root.update_idletasks()  # Ensure the label update is shown immediately

            try:
                Custom.excecute(file_paths)
                # Update the message and show the end screen after processing
                self.root.after(0, self.show_end_screen)
            except Exception as e:
                print(f"Error storing the files as csv: {e}")
            finally:
                # Hide the loading message and wheel
                self.root.after(0, self.reset_ui(show_upload=False))
        else:
            # If no files were selected, reset the UI
            self.root.after(0, self.reset_ui)

    def reset_ui(self, show_upload=True):
        # Hide the loading frame and any other temporary elements
        self.loading_frame.pack_forget()
        
        # Show the upload button
        if show_upload:
            self.start_screen_frame.pack(expand=True, fill=tk.BOTH)
        
        # Ensure the UI updates immediately
        self.root.update_idletasks()

    def show_end_screen(self):
        # Hide the start screen frame
        self.start_screen_frame.pack_forget()
        
        # Ensure the UI updates immediately
        self.root.update_idletasks()

        # Display the end screen frame with the message and buttons
        self.end_screen_frame.pack(expand=True, fill=tk.BOTH)
        
        # Ensure the UI updates immediately
        self.root.update_idletasks()

    def go_to_upload_screen(self):
        # Hide the end screen frame and buttons
        self.end_screen_frame.pack_forget()
        
        # Show the start screen frame
        self.start_screen_frame.pack(expand=True, fill=tk.BOTH)
        
        # Ensure the UI updates immediately
        self.root.update_idletasks()

    def quit_application(self):
        self.root.quit()  # Call quit method to close the application

    def open_result_log(self):
        log_path = os.path.abspath(log_file_path).replace('\\', '/')
        if os.path.exists(log_path):
            if os.name == 'nt':  # For Windows
                os.startfile(log_path)
            else:  # For Unix-based systems
                subprocess.call(['open', log_path])
        else:
            messagebox.showerror("Error", "result.log file not found.")

    def open_new_file(self):
        # Define the directory and prefix for the new file
        directory = os.path.abspath(lattest_report_path).replace('\\', '/')
        prefix = new_file_name
        
        # Search for the file that starts with the given prefix
        found_file = None
        for file_name in os.listdir(directory):
            if file_name.startswith(prefix) and file_name.endswith('.xlsx'):
                found_file = os.path.join(directory, file_name)
                break
        
        if found_file and os.path.exists(found_file):
            # Open the file based on the operating system
            if os.name == 'nt':  # For Windows
                os.startfile(found_file)
            else:  # For Unix-based systems
                subprocess.call(['open', found_file])
        else:
            messagebox.showerror("Error", f"File starting with {prefix} not found.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelUploaderApp(root)
    root.mainloop()
