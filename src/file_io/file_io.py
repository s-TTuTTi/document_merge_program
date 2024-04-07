import tkinter as tk
from tkinter import filedialog

class FileIO:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()

    def open_files(self):
        file_paths = []

        while True:
            file_path = filedialog.askopenfilename()
            if not file_path:
                break
            file_paths.append(file_path)

        return file_paths

    def save_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension='.pdf')
        return file_path
