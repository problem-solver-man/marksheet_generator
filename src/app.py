from tkinter import filedialog, messagebox
from pypdf import PdfReader
from pathlib import Path
from io import BytesIO
from PIL import Image, ImageTk, ImageSequence
from datetime import datetime
from utils import base_dir, temp_dir
import config_manager as cm
import data_manager as dm
import generate_marksheet as gm
import sys
import threading
import customtkinter as ctk
import tkinter as tk


# Prevent crash in .exe files
if getattr(sys, "frozen", False):
    sys.stdout = sys.stderr


# -------------------------
# THEME
# -------------------------

ctk.set_appearance_mode(cm.config["theme"])
ctk.set_default_color_theme("blue")


# -------------------------
# APP
# -------------------------

# Application window
app = ctk.CTk()
app.title(f"{cm.APP_NAME} v{cm.APP_VERSION}")
app.geometry("1000x600")
app.minsize(820, 480)
app.resizable(True, True)

# Asset paths
excel_path = Path(cm.config["last_excel_file"]
                  ) if cm.config["last_excel_file"] else None
output_folder = Path(cm.config["last_output_folder"])


# -------------------------
# FUNCTIONS
# -------------------------

# Display the animated national flag
flag_gif = Image.open(temp_dir / "logo/earth.gif")
flag_frames = [
    ImageTk.PhotoImage(frame.copy()) for frame in ImageSequence.Iterator(flag_gif)
]

def animate_flag(frame=0):
    flag_label.configure(image=flag_frames[frame], height=100, width=100)
    frame = (frame + 1) % len(flag_frames)
    app.after(50, animate_flag, frame)


# Select the excel file containing students' details and marks
def select_excel():
    global excel_path

    file = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])

    if file:
        excel_label.configure(text=file)        # Display file path on UI
        cm.config["last_excel_file"] = file
        cm.save_config()                        # Save file path to config.json
        excel_path = Path(file)                 # Convert file path to Path obj


# Choose the location where output will be stored
def select_output():
    global output_folder

    folder = filedialog.askdirectory(initialdir=str(output_folder))

    if folder:
        output_label.configure(text=folder)         # Display folder path on UI
        cm.config["last_output_folder"] = folder
        cm.save_config()                            # Save folder path to config.json
        output_folder = Path(folder)                # Convert folder path to Path obj


# Generate marksheets
def main():
    # Check if an excel file is choosen or not
    if not excel_path:
        messagebox.showerror(
            "Error", "Select an excel (.xlsx or .xls) file first.")
        return

    # Class of students to generate marksheet
    selected_class = class_dropdown.get()
    cm.config["last_class"] = selected_class
    cm.save_config()

    # Order of generated marksheet
    marksheet_order = order_dropdown.get()
    cm.config["marksheet_order"] = marksheet_order
    cm.save_config()

    # Load the data from excel file
    wb = dm.load_data(excel_path, selected_class)

    # Prepare overlay for each student of a class
    for ws in wb:
        # Reset the progressbar
        progressbar.set(0)

        # Get the class of students and log with date and time
        clas = ws.title.replace("Class-", "")
        now = datetime.now().strftime("%d-%m-%Y  %H:%M:%S")
        log(f"[{now}] INFO: Processing {ws.title}")

        # Read the template once per class
        template_path = (
            "templates/marksheet-format-12345.pdf"
            if clas in "1A1B2A2B34A4B5"
            else "templates/marksheet-format-67.pdf"
        )
        with open(base_dir / template_path, "rb") as f:
            template_bytes = f.read()

        # Path where marksheets of current class will be saved
        class_folder = output_folder / ws.title
        class_folder.mkdir(parents=True, exist_ok=True)

        # Format all students data in list from worksheet
        student_list = dm.prepare_data(ws)
        total_students = len(student_list)

        # Generate marksheet of each student of current class
        for rank, student in enumerate(student_list):
            # Convert all values to string in student
            student = dm.polish_data(student)

            # Generate marksheet finally
            template = PdfReader(BytesIO(template_bytes))
            gm.generate_marksheets(clas, student, rank + 1,
                                   class_folder, template, marksheet_order)
            update_progressbar(clas, rank + 1, total_students)

        # Log the successful generation of marksheets with date and time
        now = datetime.now().strftime("%d-%m-%Y  %H:%M:%S")
        log(f"[{now}] INFO: Generated marksheets | {ws.title} | folder={class_folder}")


# Create thread for marksheet generation
# Avoid UI freeze while generating many marksheets
def start_main():
    threading.Thread(target=main, daemon=True).start()


# Show the progress while generating marksheets
def update_progressbar(clas, current, total):
    percent = current / total
    progressbar.set(percent)
    progress_label.configure(text=(
        f"Generating Marksheet of Class-{clas}:  "
        f"{current}/{total} done ({round(percent*100)}%)"
    ))
    app.update_idletasks()


# Log the results
def log(msg):
    log_box.configure(state="normal")   # Make log box editable
    log_box.insert("end", f"{msg}\n")
    log_box.see("end")
    log_box.configure(state="disabled")  # Make log box read-only
    log_box.update_idletasks()


# Change the app theme to dark/light
def toggle_theme(choice):
    if choice == "Dark":
        ctk.set_appearance_mode("dark")
    else:
        ctk.set_appearance_mode("light")

    cm.config["theme"] = choice.lower()
    cm.save_config()


# Show day, date and time
def update_time():
    now = datetime.now().strftime("%A  %d-%m-%Y  %H:%M:%S")
    time_label.configure(text=now)
    app.after(1000, update_time)


# -------------------------
# UI
# -------------------------

# ========================= App Row 0 =========================

app.grid_columnconfigure(1, weight=1)
app.grid_rowconfigure(4, weight=1)

# Organization Logo
logo_image = ctk.CTkImage(
    light_image=Image.open(temp_dir/ "logo/nic-logo.png"),
    dark_image=Image.open(temp_dir / "logo/nic-logo.png"),
    size=(90, 90)
)

logo_label = ctk.CTkLabel(app, image=logo_image, text="")
logo_label.grid(row=0, column=0, rowspan=2, padx=50, pady=5, sticky="w")


# Header of the software : Organization Name
title = ctk.CTkLabel(
    app,
    text="National International College",
    font=("Segoe UI", 40, "bold")
)
title.grid(row=0, column=1, pady=5)


# Flag of the Nation
flag_label = tk.Label(app, text="")
flag_label.grid(row=0, column=2, rowspan=2, padx=50, pady=5, sticky="e")


# ========================= App Row 1 =========================

# Addres of the Organization
address = ctk.CTkLabel(
    app,
    text="India, Asia (Earth)",
    font=("Segoe UI", 20, "normal")
)
address.grid(row=1, column=1, pady=5)


# ========================= App Row 2 =========================

# Main frame
top_frame = ctk.CTkFrame(app)
top_frame.grid(row=2, column=0, columnspan=3, padx=20, pady=10, sticky="ew")
top_frame.grid_columnconfigure(4, weight=1)


# ---------------- Excel Selection ----------------
excel_btn = ctk.CTkButton(
    top_frame,
    text="Select Excel File",
    command=select_excel
)
excel_btn.grid(row=0, column=0, padx=10, pady=8, sticky="w")

excel_label = ctk.CTkLabel(
    top_frame,
    text=str(excel_path) if excel_path else "No file selected"
)
excel_label.grid(row=0, column=1, columnspan=4, padx=10, pady=8, sticky="w")


# ---------------- Class Dropdown ----------------
class_label = ctk.CTkLabel(top_frame, text="Generate marksheet for:")
class_label.grid(row=1, column=0, padx=10, pady=8, sticky="e")

class_dropdown = ctk.CTkComboBox(
    top_frame,
    width=150,
    values=[
        "All Class", "Class-1A", "Class-1B", "Class-2A", "Class-2B", "Class-3",
        "Class-4A", "Class-4B", "Class-5", "Class-6", "Class-7"
    ],
    state="readonly"
)

class_dropdown.set(cm.config["last_class"])
class_dropdown.grid(row=1, column=1, padx=10, pady=8, sticky="w")


# ---------------- Marksheet Order Dropdown ----------------
order_label = ctk.CTkLabel(top_frame, text="Marksheet order:")
order_label.grid(row=1, column=2, padx=10, pady=8, sticky="e")

order_dropdown = ctk.CTkComboBox(
    top_frame,
    width=150,
    values=["Order by Rank", "Order by Roll"],
    state="readonly"
)

order_dropdown.set(cm.config["marksheet_order"])
order_dropdown.grid(row=1, column=3, padx=10, pady=8, sticky="w")


# ---------------- Output Folder ----------------
output_btn = ctk.CTkButton(
    top_frame,
    text="Change Output Folder",
    command=select_output
)
output_btn.grid(row=2, column=0, padx=10, pady=8, sticky="w")

output_label = ctk.CTkLabel(
    top_frame,
    text=str(output_folder)
)
output_label.grid(row=2, column=1, columnspan=4, padx=10, pady=8, sticky="w")


# ---------------- Generate Button ----------------
generate_btn = ctk.CTkButton(
    top_frame,
    text="Generate Marksheets",
    height=40,
    command=start_main
)
generate_btn.grid(row=3, column=0, columnspan=5, pady=20)


# ========================= App Row 3 =========================

progress_frame = ctk.CTkFrame(app)
progress_frame.grid(row=3, column=0, columnspan=3, padx=20, pady=5, sticky="ew")
progress_frame.grid_columnconfigure(1, weight=1)

# ---------------- Progress Bar ----------------
progress_label = ctk.CTkLabel(progress_frame, text="Progress bar:")
progress_label.grid(row=3, column=0, padx=10, pady=5, sticky="e")

progressbar = ctk.CTkProgressBar(progress_frame)
progressbar.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
progressbar.set(0)


# ========================= App Row 4 =========================

# ---------------- Log Box ----------------
log_box = ctk.CTkTextbox(app, height=120)
log_box.configure(state="disabled")  # Make log box read-only
log_box.grid(row=4, column=0, columnspan=3, padx=20, pady=10, sticky="nsew")


# ========================= App Row 5 =========================

# ---------------- Status Bar ----------------
status_bar_font = ("Segoe UI", 14, "bold")

status_frame = ctk.CTkFrame(
    app,
    fg_color="#007ACC",
    corner_radius=0
)
status_frame.grid(row=5, column=0, columnspan=3, sticky="ew")

status_frame.grid_rowconfigure(0, weight=1)
status_frame.grid_columnconfigure(0, weight=1)  # empty stretch
status_frame.grid_columnconfigure(1, weight=0)  # theme toggle
status_frame.grid_columnconfigure(2, weight=0)  # time

theme_btn = ctk.CTkSegmentedButton(
    status_frame,
    corner_radius=0,
    text_color="white",
    font=status_bar_font,
    fg_color="#007ACC",
    selected_color="#005a9e",
    unselected_color="#007ACC",
    selected_hover_color="#004578",
    unselected_hover_color="#006bb3",
    values=["Light", "Dark"],
    command=toggle_theme
)
theme_btn.set(cm.config["theme"])
theme_btn.grid(row=0, column=1, padx=20, pady=1, sticky="e")

time_label = ctk.CTkLabel(
    status_frame,
    text="",
    text_color="white",
    font=status_bar_font
)
time_label.grid(row=0, column=2, padx=20, sticky="e")

# Maximize the window after 100 ms
app.after(100, lambda: app.state("zoomed"))
update_time()
animate_flag()

app.mainloop()
