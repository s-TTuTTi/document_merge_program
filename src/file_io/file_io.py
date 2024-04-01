import tkinter as tk
from tkinter import filedialog

class FileSelector:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()

    def select_files(self):
        file_paths = filedialog.askopenfilenames()
        return file_paths

if __name__ == "__main__":
    file_selector = FileSelector()
    selected_files = file_selector.select_files()
    print("Selected files:", selected_files)
