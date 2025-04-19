import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from PIL import Image
import pytesseract
from docx import Document
from fpdf import FPDF
import os
import threading

# Set path to tesseract.exe
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # Adjust if needed

def select_image():
    filepath = filedialog.askopenfilename(
        title="Select an image file",
        filetypes=[("Image Files", "*.png *.jpg *.jpeg *.bmp *.tiff")]
    )
    if filepath:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, filepath)

def extract_text_threaded():
    threading.Thread(target=extract_text).start()

def extract_text():
    image_path = entry_file.get()
    if not os.path.exists(image_path):
        messagebox.showerror("Error", "Image file does not exist!")
        return

    try:
        progress_bar.place(relx=0.5, rely=0.5, anchor="center")  # Centered inside text box
        progress_bar.start()

        img = Image.open(image_path)
        text = pytesseract.image_to_string(img)

        if not text.strip():
            messagebox.showinfo("No Text", "No text detected in the image.")
            return

        text_box.config(state=tk.NORMAL)
        text_box.delete(1.0, tk.END)
        text_box.insert(tk.END, text)
        text_box.config(state=tk.NORMAL)

    except Exception as e:
        messagebox.showerror("Error", str(e))
    finally:
        progress_bar.stop()
        progress_bar.place_forget()  # Hide progress bar

def save_as_word():
    text = text_box.get(1.0, tk.END).strip()
    if not text:
        messagebox.showwarning("Warning", "No text to save!")
        return

    filepath = filedialog.asksaveasfilename(
        defaultextension=".docx",
        filetypes=[("Word Document", "*.docx")]
    )
    if filepath:
        try:
            doc = Document()
            doc.add_paragraph(text)
            doc.save(filepath)
            messagebox.showinfo("Saved", f"Text saved as Word file:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Error", str(e))



def save_as_pdf():
    text = text_box.get(1.0, tk.END).strip()
    if not text:
        messagebox.showwarning("Warning", "No text to save!")
        return

    filepath = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF file", "*.pdf")]
    )
    if filepath:
        try:
            pdf = FPDF()
            pdf.add_page()
            
            # Add a Unicode font (DejaVuSans.ttf must be in the same folder)
            font_path = os.path.join(os.path.dirname(__file__), r"D:\Python\extract-text\DejavuSans.ttf")
            pdf.add_font('DejaVu', '', font_path, uni=True)
            pdf.set_font('DejaVu', '', 12)
            
            for line in text.split('\n'):
                pdf.multi_cell(0, 10, line)

            pdf.output(filepath)
            messagebox.showinfo("Saved", f"Text saved as PDF file:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Error", str(e))





# Helper function to get correct path for icon
def resource_path(relative_path):
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)



# --- GUI ---
root = tk.Tk()
root.title("Image to Text Extractor")
root.minsize(400, 600)  # Minimum window size

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(fill=tk.BOTH, expand=True)

# Widgets
btn_select = tk.Button(frame, text="Select Image", command=select_image)
entry_file = tk.Entry(frame, width=50)
btn_extract = tk.Button(frame, text="Extract Text", command=extract_text_threaded)

text_frame = tk.Frame(frame)
text_box = ScrolledText(text_frame, width=80, height=20, wrap=tk.WORD)

progress_bar = ttk.Progressbar(text_frame, orient=tk.HORIZONTAL, mode='indeterminate', length=200)
progress_bar.place_forget()

btn_save_word = tk.Button(frame, text="Save as Word", command=save_as_word)
btn_save_pdf = tk.Button(frame, text="Save as PDF", command=save_as_pdf)

# Layout function (dynamic)
def layout_widgets(event=None):
    width = root.winfo_width()

    # Clear old layout
    for widget in frame.winfo_children():
        widget.grid_forget()
        widget.pack_forget()

    if width < 500:  # Mobile Layout
        btn_select.pack(pady=5, fill=tk.X)
        entry_file.pack(pady=5, fill=tk.X)
        btn_extract.pack(pady=5, fill=tk.X)

        text_frame.pack(pady=10, fill=tk.BOTH, expand=True)
        text_box.pack(fill=tk.BOTH, expand=True)

        btn_save_word.pack(pady=5, fill=tk.X)
        btn_save_pdf.pack(pady=5, fill=tk.X)

    else:  # Desktop Layout
        btn_select.grid(row=0, column=0, padx=5, pady=5)
        entry_file.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        btn_extract.grid(row=0, column=2, padx=5, pady=5)

        frame.grid_columnconfigure(1, weight=1)

        text_frame.grid(row=1, column=0, columnspan=3, pady=10, sticky="nsew")
        text_box.pack(fill=tk.BOTH, expand=True)

        btn_save_word.grid(row=2, column=0, pady=5)
        btn_save_pdf.grid(row=2, column=2, pady=5)

    # Make sure frame expands correctly
    frame.pack(fill=tk.BOTH, expand=True)

# Bind resize event
root.bind("<Configure>", layout_widgets)

# Initialize layout
layout_widgets()

root.mainloop()