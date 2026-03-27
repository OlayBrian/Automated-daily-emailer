import win32com.client
import time
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os

# === Preset Directory on Desktop ===
DESKTOP_PATH = os.path.join(os.path.expanduser("~"), "Desktop")
PRESET_DIR = os.path.join(DESKTOP_PATH, "presets")
os.makedirs(PRESET_DIR, exist_ok=True)

def run_report():
    try:
        path = path_entry.get()
        table_range_str = range_entry.get()
        to = to_entry.get()
        cc = cc_entry.get()
        subject = subject_entry.get()
        body_before = body_before_text.get("1.0", "end").strip()
        body_after = body_after_text.get("1.0", "end").strip()

        excel = win32com.client.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = True


        try:
            wb = excel.Workbooks.Open(path)
        except Exception as e:
            messagebox.showerror("Workbook Error", f"Failed to open workbook:\n{e}")
            excel.Quit()
            return

        wb.RefreshAll()

        def is_refreshing():
            for sheet in wb.Sheets:
                for lo in sheet.ListObjects:
                    try:
                        if lo.QueryTable.Refreshing:
                            return True
                    except:
                        continue
            return False

        while is_refreshing():
            time.sleep(0.5)

        ws = wb.Sheets("Name of Worksheet goes here")
        table_range = ws.Range(table_range_str)

        # Copy the range (with formatting)
        table_range.Copy()


        # Create email and paste table with formatting using WordEditor
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = to
        mail.CC = cc
        mail.Subject = subject or f"Daily Report – {datetime.now():%B %d, %Y}"
        mail.Display()
        time.sleep(1)  # Add this delay


        # Paste into WordEditor
        inspector = mail.GetInspector
        word_editor = inspector.WordEditor

        selection = word_editor.Application.Selection
        selection.TypeText(body_before + "\n\n")
        selection.Paste()
        selection.TypeText("\n" + body_after)

        wb.Close(SaveChanges=False)
        excel.Quit()

        messagebox.showinfo("Success", "Formatted email is ready with Excel table.")

    except Exception as e:
        messagebox.showerror("Error", str(e))

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
    if file_path:
        path_entry.delete(0, tk.END)
        path_entry.insert(0, file_path)

def save_preset():
    data = {
        "path": path_entry.get(),
        "range": range_entry.get(),
        "to": to_entry.get(),
        "cc": cc_entry.get(),
        "subject": subject_entry.get(),
        "body_before": body_before_text.get("1.0", "end").strip(),
        "body_after": body_after_text.get("1.0", "end").strip()
    }

    preset_name = filedialog.asksaveasfilename(
        initialdir=PRESET_DIR,
        defaultextension=".json",
        filetypes=[("JSON files", "*.json")],
        title="Save Preset"
    )

    if preset_name:
        try:
            with open(preset_name, 'w') as f:
                json.dump(data, f, indent=2)
            messagebox.showinfo("Saved", f"Preset saved to:\n{preset_name}")
        except Exception as e:
            messagebox.showerror("Save Error", str(e))

def load_preset():
    preset_file = filedialog.askopenfilename(
        initialdir=PRESET_DIR,
        filetypes=[("JSON files", "*.json")],
        title="Load Preset"
    )
    if preset_file:
        try:
            with open(preset_file, 'r') as f:
                data = json.load(f)

            def safe_insert(entry, value):
                entry.delete(0, tk.END)
                entry.insert(0, value or "")

            safe_insert(path_entry, data.get("path"))
            safe_insert(range_entry, data.get("range"))
            safe_insert(to_entry, data.get("to"))
            safe_insert(cc_entry, data.get("cc"))
            safe_insert(subject_entry, data.get("subject"))

            body_before_text.delete("1.0", tk.END)
            body_before_text.insert("1.0", data.get("body_before", "Good morning,"))

            body_after_text.delete("1.0", tk.END)
            body_after_text.insert("1.0", data.get("body_after", "Thank you."))

            messagebox.showinfo("Loaded", f"Preset loaded from:\n{preset_file}")
        except Exception as e:
            messagebox.showerror("Load Error", str(e))

# === Build the GUI ===
root = tk.Tk()
root.title("Daily Report Tool")
root.geometry("750x720")

style = ttk.Style()
style.configure("TLabel", font=("Segoe UI", 10))
style.configure("TEntry", font=("Segoe UI", 10))
style.configure("TButton", font=("Segoe UI", 10))

main_frame = ttk.Frame(root, padding=15)
main_frame.pack(fill="both", expand=True)

# === Excel Section ===
ttk.Label(main_frame, text="Excel File Path:").grid(row=0, column=0, sticky="w")
path_entry = ttk.Entry(main_frame, width=70)
path_entry.grid(row=0, column=1, sticky="we")
ttk.Button(main_frame, text="Browse", command=browse_file).grid(row=0, column=2, padx=5)

ttk.Label(main_frame, text="Excel Table Range (e.g., B8:H12):").grid(row=1, column=0, columnspan=3, sticky="w", pady=(10, 0))
range_entry = ttk.Entry(main_frame)
range_entry.grid(row=2, column=0, columnspan=3, sticky="we")
range_entry.insert(0, "B8:H12")

# === Email Section ===
ttk.Label(main_frame, text="Email To:").grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))
to_entry = ttk.Entry(main_frame)
to_entry.grid(row=4, column=0, columnspan=3, sticky="we")

ttk.Label(main_frame, text="Email CC:").grid(row=5, column=0, columnspan=3, sticky="w", pady=(10, 0))
cc_entry = ttk.Entry(main_frame)
cc_entry.grid(row=6, column=0, columnspan=3, sticky="we")

ttk.Label(main_frame, text="Email Subject:").grid(row=7, column=0, columnspan=3, sticky="w", pady=(10, 0))
subject_entry = ttk.Entry(main_frame)
subject_entry.grid(row=8, column=0, columnspan=3, sticky="we")
subject_entry.insert(0, f"Daily Report – {datetime.now():%B %d, %Y}")

# === Email Body Before ===
ttk.Label(main_frame, text="Body Text (Before Table):").grid(row=9, column=0, columnspan=3, sticky="w", pady=(10, 0))
body_before_text = tk.Text(main_frame, height=5)
body_before_text.grid(row=10, column=0, columnspan=3, sticky="we")
body_before_text.insert("1.0", "Good morning,")

# === Email Body After ===
ttk.Label(main_frame, text="Body Text (After Table):").grid(row=11, column=0, columnspan=3, sticky="w", pady=(10, 0))
body_after_text = tk.Text(main_frame, height=3)
body_after_text.grid(row=12, column=0, columnspan=3, sticky="we")
body_after_text.insert("1.0", "Thank you.")

# === Buttons ===
button_frame = ttk.Frame(main_frame, padding=(0, 15))
button_frame.grid(row=13, column=0, columnspan=3)

ttk.Button(button_frame, text="Run Report", command=run_report).grid(row=0, column=0, padx=10, ipadx=10, ipady=5)
ttk.Button(button_frame, text="Save Preset", command=save_preset).grid(row=0, column=1, padx=10, ipadx=10, ipady=5)
ttk.Button(button_frame, text="Load Preset", command=load_preset).grid(row=0, column=2, padx=10, ipadx=10, ipady=5)

root.mainloop()
