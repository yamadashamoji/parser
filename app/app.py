import tkinter as tk
from tkinter import filedialog, ttk
import threading
import os
from unzip import extract_zip
from parser import xml_to_csv

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("XML to CSV Converter")
        self.geometry("400x300")

        self.create_widgets()

    def create_widgets(self):
        # ZIP file selection
        self.zip_label = tk.Label(self, text="Select ZIP file:")
        self.zip_label.pack(pady=10)

        self.zip_button = tk.Button(self, text="Browse", command=self.select_zip)
        self.zip_button.pack()

        self.zip_path = tk.StringVar()
        self.zip_entry = tk.Entry(self, textvariable=self.zip_path, width=50)
        self.zip_entry.pack(pady=5)

        # Output directory selection
        self.output_label = tk.Label(self, text="Select output directory:")
        self.output_label.pack(pady=10)

        self.output_button = tk.Button(self, text="Browse", command=self.select_output)
        self.output_button.pack()

        self.output_path = tk.StringVar()
        self.output_entry = tk.Entry(self, textvariable=self.output_path, width=50)
        self.output_entry.pack(pady=5)

        # Convert button
        self.convert_button = tk.Button(self, text="Convert", command=self.start_conversion)
        self.convert_button.pack(pady=20)

        # Progress bar
        self.progress = ttk.Progressbar(self, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=10)

        # Status label
        self.status = tk.StringVar()
        self.status_label = tk.Label(self, textvariable=self.status)
        self.status_label.pack()

    def select_zip(self):
        filename = filedialog.askopenfilename(filetypes=[("ZIP files", "*.zip")])
        self.zip_path.set(filename)

    def select_output(self):
        directory = filedialog.askdirectory()
        self.output_path.set(directory)

    def start_conversion(self):
        zip_file = self.zip_path.get()
        output_dir = self.output_path.get()

        if not zip_file or not output_dir:
            self.status.set("Please select both ZIP file and output directory.")
            return

        self.convert_button.config(state="disabled")
        self.progress["value"] = 0
        self.status.set("Starting conversion...")

        thread = threading.Thread(target=self.convert_process, args=(zip_file, output_dir))
        thread.start()

    def convert_process(self, zip_file, output_dir):
        try:
            # Extract ZIP
            self.status.set("Extracting ZIP file...")
            zip_extracted_dir = os.path.join(output_dir, "extracted")
            extract_zip(zip_file, zip_extracted_dir)
            self.progress["value"] = 25

            # Convert to CSV
            self.status.set("Converting to CSV...")
            csv_extracted_dir = os.path.join(output_dir, "csv")
            xml_to_csv(zip_extracted_dir, csv_extracted_dir)
            self.progress["value"] = 100

            self.status.set("Conversion completed successfully!")
        except Exception as e:
            self.status.set(f"Error occurred: {str(e)}")
        finally:
            self.convert_button.config(state="normal")

if __name__ == "__main__":
    app = App()
    app.mainloop()