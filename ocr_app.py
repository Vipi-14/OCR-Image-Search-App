"""
OCRApp

Initializes the main window and widgets (directory selection, search words input, start button, results display).
browse_directory: Opens a file dialog to select a directory.
start_search: Fetches the directory and words, calls search_images, and displays results.
search_images: Traverses directories, extracts text using Tesseract, searches for words, and collects results.
extract_text: Uses Pytesseract to convert image text.
get_context: Finds the sentence containing the searched word.
display_results: Displays search results in a table, binds double-click event to open the image path.
on_item_click: Opens the clicked image file path in the default web browser.

"""


import os
import pytesseract
from PIL import Image, ImageTk, ImageFilter, ImageOps
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import webbrowser
from threading import Thread
import queue
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image as PILImage

class OCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OCR Application")
        self.root.geometry("800x600")
        
        self.create_widgets()
        self.progress_queue = queue.Queue()
        
    def create_widgets(self):
        # Directory Selection
        self.dir_label = tk.Label(self.root, text="Select Directory:")
        self.dir_label.pack(pady=10)
        
        self.dir_entry = tk.Entry(self.root, width=50)
        self.dir_entry.pack(pady=10)
        
        self.dir_button = tk.Button(self.root, text="Browse", command=self.browse_directory)
        self.dir_button.pack(pady=10)
        
        # Search Words Input
        self.words_label = tk.Label(self.root, text="Enter Words to Search (comma-separated):")
        self.words_label.pack(pady=10)
        
        self.words_entry = tk.Entry(self.root, width=50)
        self.words_entry.pack(pady=10)
        
        # Start Search Button
        self.search_button = tk.Button(self.root, text="Start Search", command=self.start_search)
        self.search_button.pack(pady=10)
        
        # Progress Bar
        self.progress = ttk.Progressbar(self.root, orient=tk.HORIZONTAL, length=400, mode='determinate')
        self.progress.pack(pady=10)
        
        # Results Display
        self.tree = ttk.Treeview(self.root, columns=("Image Name", "Path", "Word Found", "Context"), show='headings')
        self.tree.heading("Image Name", text="Image Name")
        self.tree.heading("Path", text="Path")
        self.tree.heading("Word Found", text="Word Found")
        self.tree.heading("Context", text="Context")
        self.tree.pack(pady=20, fill=tk.BOTH, expand=True)
        
        # Scrollbars
        self.vsb = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        self.hsb = ttk.Scrollbar(self.root, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        self.vsb.pack(side='right', fill='y')
        self.hsb.pack(side='bottom', fill='x')
        
    def browse_directory(self):
        directory = filedialog.askdirectory()
        self.dir_entry.delete(0, tk.END)
        self.dir_entry.insert(0, directory)
        
    def start_search(self):
        directory = self.dir_entry.get()
        words = self.words_entry.get().split(',')
        
        if not directory or not words:
            messagebox.showerror("Input Error", "Please provide a valid directory and search words.")
            return
        
        self.progress['value'] = 0
        self.root.update_idletasks()
        
        thread = Thread(target=self.search_images, args=(directory, words))
        thread.start()
        self.root.after(100, self.check_queue)
        
    def search_images(self, directory, words):
        results = []
        data = []
        total_files = sum([len(files) for r, d, files in os.walk(directory) if any(f.endswith(('png', 'jpg', 'jpeg')) for f in files)])
        processed_files = 0
        
        for root, _, files in os.walk(directory):
            for file in files:
                if file.endswith(('png', 'jpg', 'jpeg')):
                    file_path = os.path.join(root, file)
                    try:
                        text = self.extract_text(file_path)
                        check_text = text.lower()
                        for word in words:
                            if word.lower() in check_text:
                                context = self.get_context(text, word)
                                results.append((file, file_path, word, context))
                                data.append({"Image Name": file, "Path": file_path, "Word Found": word, "Context": context})
                    except Exception as e:
                        print(f"Error processing {file_path}: {e}")
                    processed_files += 1
                    self.progress_queue.put((processed_files / total_files) * 100)
        
        self.display_results(results)
        self.write_sheet(data)
    
    def check_queue(self):
        try:
            progress = self.progress_queue.get_nowait()
            self.progress['value'] = progress
            self.root.update_idletasks()
        except queue.Empty:
            pass
        else:
            self.root.after(100, self.check_queue)

    def preprocess_image(self, image):
        image = image.convert('L')  # Convert to grayscale
        image = image.filter(ImageFilter.MedianFilter())  # Apply median filter
        image = ImageOps.invert(image)  # Invert colors
        image = ImageOps.autocontrast(image)  # Apply auto contrast
        return image
    
    def extract_text(self, image_path):
        image = Image.open(image_path)
        # image = self.preprocess_image(image)
        text = pytesseract.image_to_string(image)
        return text
    
    def get_context(self, text, word):
        sentences = text.split('.')
        for sentence in sentences:
            if word.lower() in sentence.lower():
                return sentence.strip()
        return ""
    
    def display_results(self, results):
        for row in self.tree.get_children():
            self.tree.delete(row)
        
        if not results:
            messagebox.showinfo("No Results", "No results found for the specified words.")
        else:
            for result in results:
                self.tree.insert('', tk.END, values=result)
                self.tree.bind("<Double-1>", self.on_item_click)

    def write_sheet(self, data):
        if os.path.exists("data.xlsx"):
            os.remove("data.xlsx")
        df = pd.DataFrame(data)

        # Create an Excel writer object
        excel_path = "data.xlsx"
        writer = pd.ExcelWriter(excel_path, engine='openpyxl')

        # Save the DataFrame to the Excel file
        df.to_excel(writer, index=False, startrow=0)

        # Load the workbook and the sheet
        workbook = writer.book
        sheet = workbook.active


        # Save the workbook
        writer.close()

        # print(f"Excel file '{excel_path}' created successfully.")
    
    def on_item_click(self, event):
        item = self.tree.selection()[0]
        item_values = self.tree.item(item, "values")
        file_path = item_values[1]
        webbrowser.open(file_path)

if __name__ == "__main__":
    root = tk.Tk()
    app = OCRApp(root)
    root.mainloop()
