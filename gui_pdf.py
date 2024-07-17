import tkinter as tk
from tkinter import filedialog, messagebox
import pikepdf
import PyPDF2
from pdf2image import convert_from_path
from pdf2docx import Converter
import os

# Function to open file dialog and select a PDF
def select_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        pdf_path.set(file_path)

# Function to rotate PDF
def rotate_pdf():
    input_path = pdf_path.get()
    if not input_path:
        messagebox.showerror("Error", "Please select a PDF file first")
        return

    output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if not output_path:
        return

    try:
        with open(input_path, 'rb') as infile:
            reader = PyPDF2.PdfReader(infile)
            writer = PyPDF2.PdfWriter()
            for page in reader.pages:
                page.rotate(180)  # Example rotation
                writer.add_page(page)
            with open(output_path, 'wb') as outfile:
                writer.write(outfile)
        messagebox.showinfo("Success", f"Rotated PDF saved as {output_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Function to encrypt PDF
def encrypt_pdf():
    input_path = pdf_path.get()
    if not input_path:
        messagebox.showerror("Error", "Please select a PDF file first")
        return

    output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if not output_path:
        return

    password = password_entry.get()
    if not password:
        messagebox.showerror("Error", "Please enter a password")
        return

    try:
        with pikepdf.open(input_path) as pdf:
            pdf.save(output_path, encryption=pikepdf.Encryption(owner=password, user=password, R=4))
        messagebox.showinfo("Success", f"Encrypted PDF saved as {output_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Function to convert PDF to Word
def pdf_to_word():
    input_path = pdf_path.get()
    if not input_path:
        messagebox.showerror("Error", "Please select a PDF file first")
        return

    output_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
    if not output_path:
        return

    try:
        cv = Converter(input_path)
        cv.convert(output_path)
        cv.close()
        messagebox.showinfo("Success", f"PDF converted to Word and saved as {output_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))


# Function to merge PDFs
def merge_pdfs():
    file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if not file_paths:
        return

    output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if not output_path:
        return

    try:
        writer = PyPDF2.PdfWriter()
        for path in file_paths:
            reader = PyPDF2.PdfReader(path)
            for page in range(len(reader.pages)):  
                writer.add_page(reader.pages[page])  
        with open(output_path, 'wb') as outfile:
            writer.write(outfile)
        messagebox.showinfo("Success", f"Merged PDF saved as {output_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))


# Set up the main window
root = tk.Tk()
root.title("PDF Toolkit")

pdf_path = tk.StringVar()

# Create the GUI components
tk.Label(root, text="PDF File:").grid(row=0, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=pdf_path, width=50).grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_pdf).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Password (for encryption):").grid(row=1, column=0, padx=10, pady=10)
password_entry = tk.Entry(root, show='*')
password_entry.grid(row=1, column=1, padx=10, pady=10)

tk.Button(root, text="Rotate PDF", command=rotate_pdf).grid(row=2, column=0, padx=10, pady=10)
tk.Button(root, text="Encrypt PDF", command=encrypt_pdf).grid(row=2, column=1, padx=10, pady=10)
tk.Button(root, text="PDF to Word", command=pdf_to_word).grid(row=2, column=2, padx=10, pady=10)
tk.Button(root, text="Merge PDFs", command=merge_pdfs).grid(row=3, column=1, padx=10, pady=10)

# Start the main event loop
root.mainloop()
