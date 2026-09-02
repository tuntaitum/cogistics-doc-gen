"""
launcher.py — GUI wrapper for the product catalog generator
Drop an Excel file onto the window to generate a PDF.
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

# ─────────────────────────────────────────────
#  CONFIGURATION
# ─────────────────────────────────────────────

OUTPUT_FOLDER = os.path.join(Path.home(), "Desktop", "Cogistics Product Suggestions Catalog")

# ─────────────────────────────────────────────
#  IMPORT YOUR SCRIPT
# ─────────────────────────────────────────────

import supplyknowledge2productcat as catalog

# ─────────────────────────────────────────────
#  CORE LOGIC
# ─────────────────────────────────────────────

def process_file(excel_path):
    excel_path = excel_path.strip().strip("{}")
    if not excel_path.lower().endswith(".xlsx"):
        set_status("❌  Please drop an .xlsx file", error=True)
        return

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    stem = Path(excel_path).stem.replace("Supply Knowledge Sheet", "").strip(" -_")
    stem = f"Product Suggestions Catalog {stem}" if stem else "Product Suggestions Catalog"
    output_pdf = os.path.join(OUTPUT_FOLDER, f"{stem}.pdf")

    set_status("⏳  Generating PDF...")
    app.after(0, lambda: btn_browse.config(state=tk.DISABLED))

    try:
        catalog.PATHS["excel_file"] = excel_path
        catalog.PATHS["output_pdf"] = output_pdf
        products = catalog.read_excel(excel_path)
        # Schedule PDF generation on main thread
        app.after(0, lambda: run_generate(products, output_pdf))
    except Exception as e:
        set_status(f"❌  Error: {e}", error=True)
        app.after(0, lambda: btn_browse.config(state=tk.NORMAL))

def run_generate(products, output_pdf):
    try:
        catalog.generate_pdf(products, output_pdf)
        set_status(f"✅  Saved → {output_pdf}", error=False)
    except Exception as e:
        set_status(f"❌  Error: {e}", error=True)
    finally:
        btn_browse.config(state=tk.NORMAL)

def run_in_thread(excel_path):
    threading.Thread(target=process_file, args=(excel_path,), daemon=True).start()

def on_drop(event):
    run_in_thread(event.data)

def on_browse():
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if path:
        run_in_thread(path)

def set_status(msg, error=False):
    def _update():
        status_var.set(msg)
        status_label.config(fg="#cc0000" if error else "#1A3C5E")
    app.after(0, _update)

# ─────────────────────────────────────────────
#  UI
# ─────────────────────────────────────────────

app = tk.Tk()
app.title("Cogistics Product Catalog Generator")
app.geometry("480x280")
app.resizable(False, False)
app.configure(bg="white")

# Drop zone
drop_frame = tk.Frame(app, bg="#F5F7FA", relief="solid", bd=1,
                      highlightbackground="#CCCCCC", highlightthickness=1)
drop_frame.place(x=30, y=20, width=420, height=160)

tk.Label(drop_frame, text="Click Browse to select file",
         font=("Helvetica", 14, "bold"), bg="#F5F7FA", fg="#1A3C5E").pack(expand=True)
tk.Label(drop_frame, text=".xlsx files only",
         font=("Helvetica", 9), bg="#F5F7FA", fg="#888888").pack(pady=(0, 20))


# Browse button
btn_browse = tk.Button(app, text="Browse for file", command=on_browse,
                       bg="#1A3C5E", fg="white", font=("Helvetica", 10),
                       relief="flat", padx=16, pady=6, cursor="hand2")
btn_browse.place(x=175, y=200)

# Status label
status_var = tk.StringVar(value="Ready — select an Excel file to begin")
status_label = tk.Label(app, textvariable=status_var, font=("Helvetica", 9),
                        bg="white", fg="#1A3C5E", wraplength=420)
status_label.place(x=30, y=245)

app.mainloop()