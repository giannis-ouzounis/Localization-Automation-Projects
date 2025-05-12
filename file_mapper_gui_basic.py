# -*- coding: utf-8 -*-
"""
GUI tool to manually map one source file to one or more target files.
Displays mappings in a scrollable text box for easy review and manual removal.
Allows adding and removing files from both source and target lists via dialogs.
Supports single selection for source and multiple selections for target files.
Does not include logic for processing or saving the mappings beyond display.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class FileMapperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Mapper")
        self.root.minsize(800, 400)  # Set the minimum size of the window

        self.source_files = []
        self.target_files = []
        self.selected_source_index = None
        self.selected_target_indices = []

        self.create_widgets()

    def create_widgets(self):
        
        frame = ttk.Frame(self.root, padding="10")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Source Listbox and Scrollbars
        ttk.Label(frame, text="Source Files").grid(row=0, column=0, padx=5, pady=5)

        source_frame = ttk.Frame(frame)
        source_frame.grid(row=1, column=0, rowspan=2, padx=5, pady=5, sticky='nsew')

        self.source_listbox = tk.Listbox(source_frame, selectmode=tk.SINGLE, width=80)
        self.source_listbox.grid(row=0, column=0, sticky='nsew')
        self.source_listbox.bind('<<ListboxSelect>>', self.update_selected_source)

        source_vsb = ttk.Scrollbar(source_frame, orient="vertical", command=self.source_listbox.yview)
        source_vsb.grid(row=0, column=1, sticky='ns')
        self.source_listbox.configure(yscrollcommand=source_vsb.set)

        source_hsb = ttk.Scrollbar(source_frame, orient="horizontal", command=self.source_listbox.xview)
        source_hsb.grid(row=1, column=0, sticky='ew')
        self.source_listbox.configure(xscrollcommand=source_hsb.set)

        ttk.Button(frame, text="Add Source Files", command=self.add_source_files).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Remove Source Files", command=self.remove_source_files).grid(row=2, column=1, padx=5, pady=5)

        # Target Listbox and Scrollbars
        ttk.Label(frame, text="Target Files").grid(row=3, column=0, padx=5, pady=5)

        target_frame = ttk.Frame(frame)
        target_frame.grid(row=4, column=0, rowspan=2, padx=5, pady=5, sticky='nsew')

        self.target_listbox = tk.Listbox(target_frame, selectmode=tk.MULTIPLE, width=80)
        self.target_listbox.grid(row=0, column=0, sticky='nsew')
        self.target_listbox.bind('<<ListboxSelect>>', self.update_selected_targets)

        target_vsb = ttk.Scrollbar(target_frame, orient="vertical", command=self.target_listbox.yview)
        target_vsb.grid(row=0, column=1, sticky='ns')
        self.target_listbox.configure(yscrollcommand=target_vsb.set)

        target_hsb = ttk.Scrollbar(target_frame, orient="horizontal", command=self.target_listbox.xview)
        target_hsb.grid(row=1, column=0, sticky='ew')
        self.target_listbox.configure(xscrollcommand=target_hsb.set)

        ttk.Button(frame, text="Add Target Files", command=self.add_target_files).grid(row=4, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Remove Target Files", command=self.remove_target_files).grid(row=5, column=1, padx=5, pady=5)

        ttk.Button(frame, text="Map Files", command=self.map_files).grid(row=6, column=0, columnspan=2, pady=5)

        # Mapping display and Scrollbars
        mapping_frame = ttk.Frame(frame)
        mapping_frame.grid(row=7, column=0, columnspan=2, pady=5, sticky='nsew')

        self.mapping_display = tk.Text(mapping_frame, width=100, height=10)
        self.mapping_display.grid(row=0, column=0, sticky='nsew')

        mapping_vsb = ttk.Scrollbar(mapping_frame, orient="vertical", command=self.mapping_display.yview)
        mapping_vsb.grid(row=0, column=1, sticky='ns')
        self.mapping_display.configure(yscrollcommand=mapping_vsb.set)

        mapping_hsb = ttk.Scrollbar(mapping_frame, orient="horizontal", command=self.mapping_display.xview)
        mapping_hsb.grid(row=1, column=0, sticky='ew')
        self.mapping_display.configure(xscrollcommand=mapping_hsb.set)

        ttk.Button(frame, text="Remove Selected Mapping", command=self.remove_selected_mapping).grid(row=8, column=0, columnspan=2, pady=5)

    def add_source_files(self):
        files = filedialog.askopenfilenames(title="Select Source Files")
        for file in files:
            self.source_files.append(file)
            self.source_listbox.insert(tk.END, file)

    def add_target_files(self):
        files = filedialog.askopenfilenames(title="Select Target Files")
        for file in files:
            self.target_files.append(file)
            self.target_listbox.insert(tk.END, file)

    def remove_source_files(self):
        selected_indices = self.source_listbox.curselection()
        for index in selected_indices[::-1]:  # Reverse the order to avoid issues while deleting
            self.source_listbox.delete(index)
            del self.source_files[index]

    def remove_target_files(self):
        selected_indices = self.target_listbox.curselection()
        for index in selected_indices[::-1]:  # Reverse the order to avoid issues while deleting
            self.target_listbox.delete(index)
            del self.target_files[index]

    def update_selected_source(self, event):
        selection = self.source_listbox.curselection()
        if selection:
            self.selected_source_index = selection[0]

    def update_selected_targets(self, event):
        self.selected_target_indices = self.target_listbox.curselection()

    def map_files(self):
        if self.selected_source_index is None:
            messagebox.showerror("Error", "Please select a source file.")
            return

        if not self.selected_target_indices:
            messagebox.showerror("Error", "Please select at least one target file.")
            return

        source_file = self.source_files[self.selected_source_index]
        target_files = [self.target_files[i] for i in self.selected_target_indices]

        mapping = f"{source_file} -> {', '.join(target_files)}\n"
        self.mapping_display.insert(tk.END, mapping)

    def remove_selected_mapping(self):
        try:
            selected_text = self.mapping_display.selection_get()
            if selected_text:
                start_index = self.mapping_display.search(selected_text, "1.0", tk.END)
                end_index = f"{start_index} + {len(selected_text)}c"
                self.mapping_display.delete(start_index, end_index)
        except tk.TclError:
            messagebox.showerror("Error", "Please select a mapping to remove.")

if __name__ == "__main__":
    root = tk.Tk()
    app = FileMapperApp(root)
    root.mainloop()
