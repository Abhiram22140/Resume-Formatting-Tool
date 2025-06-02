# gui.py

import os
import tkinter as tk
from tkinter import filedialog, messagebox

from parser import parse_resume
from ppt_merger import merge_into_template

# Path to the PPT template you uploaded as Resume.pptx
TEMPLATE_PATH = os.path.join("templates", "Resume.pptx")
OUTPUT_DIR    = "output"

def on_select_and_format():
    """
    Invoked when the user clicks “Select & Format Résumé”.
    1) Let them pick a résumé file (.docx or .pdf)
    2) Parse it → dict
    3) Merge into the PPT template
    4) Save the result in output/
    5) Show success or error dialog
    """
    # 1) File dialog
    filepath = filedialog.askopenfilename(
        title="Select résumé to format",
        filetypes=[("Word Documents", "*.docx"), ("PDF Files", "*.pdf"), ("All Files", "*.*")]
    )
    if not filepath:
        return  # user cancelled

    # 2) Parse résumé
    try:
        parsed = parse_resume(filepath)
    except Exception as e:
        messagebox.showerror("Parsing Error", f"Failed to parse résumé:\n{e}")
        return

    # 3) Build output filename
    base = os.path.basename(filepath)
    name_no_ext, _ = os.path.splitext(base)
    output_name = f"formatted_{name_no_ext}.pptx"
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_path = os.path.join(OUTPUT_DIR, output_name)

    # 4) Merge into template
    try:
        merge_into_template(parsed, TEMPLATE_PATH, output_path)
    except Exception as e:
        messagebox.showerror("Merge Error", f"Failed to generate formatted PPT:\n{e}")
        return

    # 5) Success message
    messagebox.showinfo("Success", f"Formatted PPT saved to:\n{output_path}")

def build_gui():
    """
    Build and run the Tkinter window.
    """
    root = tk.Tk()
    root.title("Résumé → PPT Formatter")
    root.geometry("380x150")
    root.resizable(False, False)

    tk.Label(
        root,
        text="Upload your résumé to generate the formatted PPT:",
        font=("Segoe UI", 11)
    ).pack(pady=15)

    tk.Button(
        root,
        text="Select & Format Résumé",
        command=on_select_and_format,
        width=24,
        height=2
    ).pack()

    root.mainloop()

if __name__ == "__main__":
    build_gui()
