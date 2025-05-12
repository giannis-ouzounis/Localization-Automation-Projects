# -*- coding: utf-8 -*-
"""
GUI tool to rename Word .docx files originated from XTM in one folder by matching their truncated names (first 50 characters) with files in another folder.
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox

def get_word_files(path):
    return [f for f in os.listdir(path) if f.lower().endswith(".docx") and os.path.isfile(os.path.join(path, f))]

def truncate_name(name):
    return name[:50] if len(name) > 50 else name

def select_folder(entry_field):
    folder = filedialog.askdirectory()
    if folder:
        entry_field.delete(0, tk.END)
        entry_field.insert(0, folder)

def run_renaming(path1, path2):
    if not os.path.isdir(path1) or not os.path.isdir(path2):
        messagebox.showerror("Invalid Paths", "One or both folder paths are invalid.")
        return

    files1 = get_word_files(path1)
    files2 = get_word_files(path2)

    map1 = {truncate_name(f): f for f in files1 if len(f) > 50}
    map2 = {truncate_name(f): f for f in files2 if len(f) > 50}

    renamed_count = 0

    for key in map1:
        if key in map2:
            old_name = map2[key]
            new_name = map1[key]
            if old_name != new_name:
                try:
                    os.rename(os.path.join(path2, old_name), os.path.join(path2, new_name))
                    renamed_count += 1
                except Exception:
                    pass  # silently skip errors

    messagebox.showinfo("Renaming Complete", f"Process complete.\nTotal files renamed: {renamed_count}")

def create_ui():
    root = tk.Tk()
    root.title("XTM bilingual Words - Renamer")

    # Folder 1
    tk.Label(root, text="Folder 1 (Source):").grid(row=0, column=0, sticky="e", padx=5, pady=10)
    folder1_entry = tk.Entry(root, width=50)
    folder1_entry.grid(row=0, column=1, padx=5)
    tk.Button(root, text="Browse", command=lambda: select_folder(folder1_entry)).grid(row=0, column=2, padx=5)

    # Folder 2
    tk.Label(root, text="Folder 2 (To Rename):").grid(row=1, column=0, sticky="e", padx=5, pady=10)
    folder2_entry = tk.Entry(root, width=50)
    folder2_entry.grid(row=1, column=1, padx=5)
    tk.Button(root, text="Browse", command=lambda: select_folder(folder2_entry)).grid(row=1, column=2, padx=5)

    # Rename Button
    tk.Button(root, text="Rename Files", width=20, height=2,
              command=lambda: run_renaming(folder1_entry.get().strip('"'), folder2_entry.get().strip('"'))
    ).grid(row=2, column=0, columnspan=3, pady=20)

    root.mainloop()

if __name__ == "__main__":
    create_ui()
