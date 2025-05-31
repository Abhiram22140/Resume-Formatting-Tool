import tkinter as tk
from tkinter import filedialog, messagebox
from parser import parse_resume
from formatter import format_resume

def select_file():
    filepath = filedialog.askopenfilename(filetypes=[("Resume files", "*.docx *.pdf")])
    if not filepath:
        return
    try:
        parsed_data = parse_resume(filepath)
        output_path = format_resume(parsed_data)
        messagebox.showinfo("Success", f"Formatted resume saved at:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to process file:\n{e}")

app = tk.Tk()
app.title("Client-Specific Resume Formatter")
app.geometry("400x200")

label = tk.Label(app, text="Upload Candidate Resume (.docx or .pdf)", font=("Helvetica", 14))
label.pack(pady=20)

upload_btn = tk.Button(app, text="Choose Resume", command=select_file, bg="#4CAF50", fg="white", padx=20, pady=10)
upload_btn.pack()

app.mainloop()
