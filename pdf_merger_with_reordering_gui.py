# -*- coding: utf-8 -*-
"""
GUI tool to merge multiple PDF files with support for reordering, removing, and saving to a new file.
"""

import PyPDF2
import os
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, END, MULTIPLE, Scrollbar

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Merger with Reordering")

        self.file_list = []  # Stores full file paths

        # Frame to hold the listbox and scrollbars
        list_frame = tk.Frame(root)
        list_frame.pack(padx=10, pady=10)

        # Listbox to display filenames only
        self.listbox = Listbox(list_frame, selectmode=MULTIPLE, width=60, height=15, xscrollcommand=lambda *args: h_scroll.set(*args))
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH)

        # Vertical Scrollbar
        v_scroll = Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=v_scroll.set)

        # Horizontal Scrollbar
        h_scroll = Scrollbar(root, orient=tk.HORIZONTAL, command=self.listbox.xview)
        h_scroll.pack(fill=tk.X)
        self.listbox.config(xscrollcommand=h_scroll.set)

        # Buttons for file operations
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=5)

        tk.Button(btn_frame, text="Add PDFs", command=self.add_files).grid(row=0, column=0, padx=5)
        tk.Button(btn_frame, text="Move Up", command=self.move_up).grid(row=0, column=1, padx=5)
        tk.Button(btn_frame, text="Move Down", command=self.move_down).grid(row=0, column=2, padx=5)
        tk.Button(btn_frame, text="Remove Selected", command=self.remove_selected).grid(row=0, column=3, padx=5)
        tk.Button(btn_frame, text="Merge PDFs", command=self.merge_pdfs).grid(row=0, column=4, padx=5)

    def add_files(self):
        files = filedialog.askopenfilenames(title="Select PDF Files", filetypes=[("PDF Files", "*.pdf")])
        for file in files:
            if file not in self.file_list:
                self.file_list.append(file)
                filename = os.path.basename(file)  # Show only the filename
                self.listbox.insert(END, filename)

    def move_up(self):
        selected = self.listbox.curselection()
        for i in selected:
            if i > 0:
                self.file_list[i], self.file_list[i - 1] = self.file_list[i - 1], self.file_list[i]
                self.refresh_listbox()

    def move_down(self):
        selected = reversed(self.listbox.curselection())
        for i in selected:
            if i < len(self.file_list) - 1:
                self.file_list[i], self.file_list[i + 1] = self.file_list[i + 1], self.file_list[i]
                self.refresh_listbox()

    def remove_selected(self):
        selected = list(self.listbox.curselection())
        for i in reversed(selected):
            del self.file_list[i]
        self.refresh_listbox()

    def refresh_listbox(self):
        self.listbox.delete(0, END)
        for file in self.file_list:
            filename = os.path.basename(file)
            self.listbox.insert(END, filename)

    def merge_pdfs(self):
        if not self.file_list:
            messagebox.showwarning("No Files", "Please add PDF files to merge.")
            return

        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not output_path:
            return

        pdf_writer = PyPDF2.PdfWriter()
        try:
            for file in self.file_list:
                with open(file, 'rb') as pdf_file:
                    pdf_reader = PyPDF2.PdfReader(pdf_file)
                    for page in pdf_reader.pages:
                        pdf_writer.add_page(page)

            with open(output_path, 'wb') as output_pdf:
                pdf_writer.write(output_pdf)

            messagebox.showinfo("Success", f"PDFs merged successfully into:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop()
