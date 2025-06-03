# gui.py

import os
import tkinter as tk
from tkinter import filedialog, messagebox

from parser import parse_resume
from ppt_merger import merge_into_template

# Path to the PPT template
TEMPLATE_PATH = os.path.join("templates", "Resume.pptx")
OUTPUT_DIR    = "output"

def on_select_and_format():
    """
    1) Let user pick a résumé file (.docx or .pdf)
    2) parse_resume(filepath)
    3) merge_into_template(parsed, TEMPLATE_PATH, output_path)
    4) show success or error
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

    # 5) Success
    messagebox.showinfo("Success", f"Formatted PPT saved to:\n{output_path}")

def ask_for_details(parsed_data):
    """Opens a new window to ask for user's Name and Role."""
    details_window = tk.Toplevel()
    details_window.title("Enter Details")
    details_window.geometry("300x150")
    details_window.transient(tk.Tk()) # Keep window on top of the main window
    details_window.grab_set() # Modal window

    tk.Label(details_window, text="Enter your Name:").pack(pady=5)
    name_entry = tk.Entry(details_window, width=40)
    name_entry.pack()
    if 'name' in parsed_data and parsed_data['name']:
        name_entry.insert(0, parsed_data['name'])

    tk.Label(details_window, text="Enter your Role:").pack(pady=5)
    role_entry = tk.Entry(details_window, width=40)
    role_entry.pack()
    if 'role' in parsed_data and parsed_data['role']:
        role_entry.insert(0, parsed_data['role'])

    def submit_details():
        name = name_entry.get().strip()
        role = role_entry.get().strip()
        if name:
            parsed_data['name'] = name
        if role:
            parsed_data['role'] = role
        details_window.destroy()

    tk.Button(details_window, text="Submit", command=submit_details).pack(pady=10)

    # Wait until the details window is closed
    details_window.wait_window()


def build_gui():
    root = tk.Tk()
    root.title("Résumé → PPT Formatter")
    root.geometry("400x160")
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
        width=26,
        height=2
    ).pack()

    root.mainloop()

if __name__ == "__main__":
    build_gui()
