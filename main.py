import re

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
import os
import PyPDF2


class PDFMergerApp(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()
        self.TkdndVersion = TkinterDnD._require(self)

        self.title("PDF Merger")
        self.geometry("600x550")

        self.pdf_files = []
        self.last_used_paths = self.load_last_used_paths()

        self.create_widgets()

    def create_widgets(self):
        # File selection frame
        self.file_frame = ctk.CTkFrame(self)
        self.file_frame.pack(pady=10, padx=10, fill="x")

        self.file_label = ctk.CTkLabel(self.file_frame, text="Selected PDF Files:")
        self.file_label.pack(anchor="w")

        self.file_listbox = tk.Listbox(self.file_frame, selectmode=tk.SINGLE)
        self.file_listbox.pack(fill="x", padx=5, pady=5)
        self.file_listbox.bind('<Delete>', self.remove_selected_pdf)
        self.file_listbox.bind('<Double-1>', self.remove_selected_pdf)

        self.buttons_frame = ctk.CTkFrame(self.file_frame)
        self.buttons_frame.pack(fill="x", padx=5, pady=5)

        # Make columns expand equally
        self.buttons_frame.grid_columnconfigure(0, weight=1)
        self.buttons_frame.grid_columnconfigure(1, weight=1)
        self.buttons_frame.grid_columnconfigure(2, weight=1)

        self.up_button = ctk.CTkButton(self.buttons_frame, text="↑", command=self.move_up)
        self.up_button.grid(row=0, column=0, padx=5, pady=5, sticky='nsew')

        self.down_button = ctk.CTkButton(self.buttons_frame, text="↓", command=self.move_down)
        self.down_button.grid(row=0, column=1, padx=5, pady=5, sticky='nsew')

        self.remove_button = ctk.CTkButton(self.buttons_frame, text="-", command=self.remove_selected_pdf)
        self.remove_button.grid(row=0, column=2, padx=5, pady=5, sticky='nsew')

        self.browse_button = ctk.CTkButton(self.file_frame, text="Browse Files", command=self.browse_files)
        self.browse_button.pack(pady=5)

        # Output name and path frame
        self.output_frame = ctk.CTkFrame(self)
        self.output_frame.pack(pady=10, padx=10, fill="x")

        self.output_name_label = ctk.CTkLabel(self.output_frame, text="Merged PDF Name:")
        self.output_name_label.pack(anchor="w")

        self.output_name_entry = ctk.CTkEntry(self.output_frame)
        self.output_name_entry.pack(fill="x", padx=5, pady=5)
        self.output_name_entry.insert(0, "merged.pdf")

        self.output_path_label = ctk.CTkLabel(self.output_frame, text="Output Path:")
        self.output_path_label.pack(anchor="w")

        self.output_entry = ctk.CTkEntry(self.output_frame)
        self.output_entry.pack(fill="x", padx=5, pady=5)
        self.output_entry.insert(0, self.last_used_paths.get("output_path", ""))

        self.output_browse_button = ctk.CTkButton(self.output_frame, text="Browse Output Folder",
                                                  command=self.browse_output_folder)
        self.output_browse_button.pack(pady=5)

        # Merge button
        self.merge_button = ctk.CTkButton(self, text="Merge PDFs", command=self.merge_pdfs)
        self.merge_button.pack(pady=20)

        # Enable drag and drop
        self.file_listbox.drop_target_register(DND_FILES)
        self.file_listbox.dnd_bind('<<Drop>>', self.get_path)

    def browse_files(self):
        initial_dir = self.last_used_paths.get("initial_pdf_dir", os.path.expanduser("~"))
        file_paths = filedialog.askopenfilenames(initialdir=initial_dir, title="Select PDF files",
                                                 filetypes=[("PDF Files", "*.pdf")])
        if file_paths:
            self.pdf_files.extend(file_paths)
            self.update_file_listbox()
            self.last_used_paths["initial_pdf_dir"] = os.path.dirname(file_paths[0])

    def update_file_listbox(self):
        self.file_listbox.delete(0, tk.END)
        for file in self.pdf_files:
            self.file_listbox.insert(tk.END, file)

    def remove_selected_pdf(self, event=None):
        selected_indices = self.file_listbox.curselection()
        if selected_indices:
            selected_index = selected_indices[0]
            self.pdf_files.pop(selected_index)
            self.update_file_listbox()

    def move_up(self):
        selected_indices = self.file_listbox.curselection()
        if selected_indices:
            index = selected_indices[0]
            if index > 0:
                self.pdf_files[index], self.pdf_files[index - 1] = self.pdf_files[index - 1], self.pdf_files[index]
                self.update_file_listbox()
                self.file_listbox.selection_set(index - 1)

    def move_down(self):
        selected_indices = self.file_listbox.curselection()
        if selected_indices:
            index = selected_indices[0]
            if index < len(self.pdf_files) - 1:
                self.pdf_files[index], self.pdf_files[index + 1] = self.pdf_files[index + 1], self.pdf_files[index]
                self.update_file_listbox()
                self.file_listbox.selection_set(index + 1)

    def browse_output_folder(self):
        folder_path = filedialog.askdirectory(title="Select Output Folder", initialdir=self.output_entry.get())
        if folder_path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, folder_path)
            self.last_used_paths["output_path"] = folder_path

    def merge_pdfs(self):
        output_path = self.output_entry.get()
        output_name = self.output_name_entry.get()
        if not output_path:
            messagebox.showerror("Error", "Please specify an output folder.")
            return
        if not output_name:
            messagebox.showerror("Error", "Please specify a name for the merged PDF.")
            return
        if not self.pdf_files:
            messagebox.showerror("Error", "No PDF files selected.")
            return

        output_file = os.path.join(output_path, output_name)

        merger = PyPDF2.PdfMerger()
        for pdf in self.pdf_files:
            merger.append(pdf)

        merger.write(output_file)
        merger.close()

        messagebox.showinfo("Success", f"PDF files merged successfully into {output_file}")
        self.save_last_used_paths()

    def save_last_used_paths(self):
        with open("last_paths.txt", "w") as file:
            for key, value in self.last_used_paths.items():
                file.write(f"{key}:{value}\n")

    def load_last_used_paths(self):
        paths = {}
        if os.path.exists("last_paths.txt"):
            with open("last_paths.txt", "r") as file:
                for line in file:
                    key, value = line.strip().split(":", 1)
                    paths[key] = value
        return paths

    def get_path(self, event):
        dropped_files = event.data.replace("{", "").replace("}", "")
        pattern = r'(?<=\S) (?=[A-Z]:)'
        file_list = re.split(pattern, dropped_files)

        for file in file_list:
            self.pdf_files.append(file)
            self.update_file_listbox()


# Run the application
if __name__ == "__main__":
    ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light"
    ctk.set_default_color_theme("blue")  # Themes: "blue" (default), "green", "dark-blue"
    app = PDFMergerApp()
    app.mainloop()
