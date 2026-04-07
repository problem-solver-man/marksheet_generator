from tkinter import filedialog, messagebox
from pathlib import Path
from PIL import Image, ImageTk, ImageSequence
from datetime import datetime
from utils import base_dir, temp_dir, CTkinterHandler, LanguageManager
import marksheet_generator as gm
import config_manager as cm
import customtkinter as ctk
import tkinter as tk
import threading
import logging
import sys
import atexit


# Flag to track whether a task is in progress
is_processing = False

# Preventing crash in .exe files
if getattr(sys, "frozen", False):
    sys.stdout = sys.stderr

# Create logger
logger = logging.getLogger()
logger.setLevel(logging.INFO)

# Log unhandled errors
def handle_exception(exc_type, exc_value, exc_traceback):
    logger.error(
        "Unhandled exception",
        exc_info=(exc_type, exc_value, exc_traceback)
    )
sys.excepthook = handle_exception

def thread_exception_handler(args):
    logger.error(
        "Thread exception",
        exc_info=(args.exc_type, args.exc_value, args.exc_traceback)
    )
threading.excepthook = thread_exception_handler

# Asset paths
excel_path = Path(cm.config["last_excel_file"])
output_folder = Path(cm.config["last_output_folder"])

# -------------------------
# THEME
# -------------------------

ctk.set_appearance_mode(cm.config["theme"].lower())
ctk.set_default_color_theme("blue")


# -------------------------
# APP
# -------------------------

# Application window
app = ctk.CTk()
app.iconbitmap(temp_dir / "logos/nic-logo.ico")
app.title(f"{cm.APP_NAME} v{cm.APP_VERSION}")
app.geometry("1000x600")
app.minsize(1000, 600)
app.resizable(True, True)


# Application UI language
lang = LanguageManager(cm.config["language"])

# -------------------------
# FUNCTIONS
# -------------------------

# Display the animated national flag
flag_gif = Image.open(temp_dir / "logos/earth.gif")
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

    file = filedialog.askopenfilename(
        filetypes=[("Excel", "*.xlsx *.xls")],
        initialdir=str(excel_path.parent)
    )

    if file:
        excel_path = Path(file)

        # Display file path on UI
        excel_label.configure(text=str(excel_path))

        logger.info(
            f"{excel_path.name} is selected successfully "
            f"from {excel_path.parent}"
        )

        # Save file path to config.json
        cm.config["last_excel_file"] = str(excel_path)
        cm.save_config()


# Choose the location where output will be stored
def select_output():
    global output_folder

    folder = filedialog.askdirectory(initialdir=str(output_folder))

    if folder:
        output_folder = Path(folder)

        # Display folder path on UI
        output_label.configure(text=str(output_folder))

        logger.info(f"Output folder changed to {output_folder}")

        # Save folder path to config.json
        cm.config["last_output_folder"] = str(output_folder)
        cm.save_config()


# Generate marksheets
def main():
    global is_processing

    try:
        progressbar.set(0)      # Reset the progress bar

        # Check if an excel file is choosen or not
        if not excel_path.is_file():
            logger.error(
                f"{excel_path.name} is no longer available. "
                "It may have been moved, renamed, or deleted. "
                "Please reselect the file to continue."
            )
            messagebox.showerror(lang.t("error"), lang.t("file_abs_err_msg"))
            return

        # Evaluation term to mention in marksheets
        eval_term = term_dropdown.get()
        cm.config["evaluation_term"] = eval_term
        cm.save_config()

        # Evaluation year to mention in marksheets
        eval_year = year_dropdown.get()
        cm.config["evaluation_year"] = eval_year
        cm.save_config()

        # Make evaluation term and year mandatory
        if eval_term == "Select" or eval_year == "Select":
            messagebox.showerror(lang.t("error"), lang.t("no_eval_msg"))
            return

        # Class of students to generate marksheets
        selected_class = class_dropdown.get()
        cm.config["last_class"] = selected_class
        cm.save_config()

        # Order of generated marksheets
        marksheet_order = order_dropdown.get()
        cm.config["marksheet_order"] = marksheet_order
        cm.save_config()

        # Output type of generated marksheets
        output_type = output_type_dropdown.get()
        cm.config["output_type"] = output_type
        cm.save_config()

        # Generate marksheets
        gm.generate_marksheets(
            excel_path, output_folder,
            eval_term, eval_year,
            selected_class, marksheet_order, output_type,
            update_progressbar
        )
    except Exception as e:
        logger.error(f"Unhandled exception in main:\n{e}")
        messagebox.showerror(lang.t("error"), str(e))
    finally:
        # Reset flag, update generate button text and enable language switch
        def reset_ui():
            global is_processing
            is_processing = False
            generate_btn.configure(text=lang.t("generate_btn_inactive"))
            language_btn.configure(state="normal")
            status_label.configure(text=lang.t("status_ready"))

        app.after(0, reset_ui)


# Create thread for marksheet generation
# Avoid UI freeze while generating many marksheets
def start_main():
    global is_processing

     # Check if a task is already running
    if is_processing:
        messagebox.showinfo(lang.t("wait"), lang.t("wait_msg"))
        return

    # Set flag to indicate task is running
    is_processing = True

    # Update generate button text and disable language switch
    generate_btn.configure(text=lang.t("generate_btn_active"))
    language_btn.configure(state="disabled")
    status_label.configure(text=lang.t("status_progress"))

    threading.Thread(target=main, daemon=True).start()


# Show the progress while generating marksheets
def update_progressbar(clas, current, total):
    percent = current / total
    progressbar.set(percent)
    progress_label.configure(text=(
        f"{lang.t("progress_txt1")} {clas}:  "
        f"{current}/{total} {lang.t("progress_txt2")} ({round(percent*100)}%)"
    ))
    app.update_idletasks()


# Change the app theme to dark/light
def toggle_theme(choice):
    ctk.set_appearance_mode(choice.lower())
    logger.info(f"Theme changed to {choice}")

    cm.config["theme"] = choice
    cm.save_config()


# Show day, date and time
def update_time():
    now = datetime.now().strftime("%A  %d-%m-%Y  %H:%M:%S")
    time_label.configure(text=now)
    app.after(1000, update_time)


# Change language
def change_language(choice):
    choice_value = ""
    if choice == "हिंदी":
        choice_value = "hi"
    else:
        choice_value = "en"

    lang.set_language(choice_value)
    refresh_ui(choice_value)
    logger.info(f"Language changed to {choice}")

    cm.config["language"] = choice_value
    cm.save_config()
    
def refresh_ui(mode):
    if mode == "hi":
        excel_btn.configure(text=lang.t("excel_file_btn"), width=200)
        output_btn.configure(text=lang.t("output_folder_btn"), width=200)
    else:
        excel_btn.configure(text=lang.t("excel_file_btn"), width=175)
        output_btn.configure(text=lang.t("output_folder_btn"), width=175)
    
    if is_processing:
        generate_btn.configure(text=lang.t("generate_btn_active"))
        status_label.configure(text=lang.t("status_progress"))
    else:
        generate_btn.configure(text=lang.t("generate_btn_inactive"))
        status_label.configure(text=lang.t("status_ready"))

    title.configure(text=lang.t("school_name"))
    address.configure(text=lang.t("school_address"))
    term_label.configure(text=lang.t("term_txt"))
    year_label.configure(text=lang.t("year_txt"))
    class_label.configure(text=lang.t("class_txt"))
    order_label.configure(text=lang.t("order_txt"))
    output_type.configure(text=lang.t("type_txt"))
    progress_label.configure(text=lang.t("progress_bar"))


# Action before closing the app
def on_closing():
    if is_processing:
        choice = messagebox.askyesnocancel(
            lang.t("task_running"),
            (f"{lang.t("task_running_msg")}\n\n"
             f"{lang.t("task_running_yes")}\n"
             f"{lang.t("task_running_no")}\n"
             f"{lang.t("task_running_cancel")}")
        )

        if choice is True:
            logger.info("Application is shutting down while task is in progress...")
            app.destroy()
        elif choice is False:
            pass
        else:
            return
    else:
        logger.info("Application is shutting down normally...")
        app.destroy()


def log_exit():
    logger.info("Application exited")
    logging.shutdown()


# -------------------------
# UI
# -------------------------

# ========================= App Row 0 : Title =========================

app.grid_columnconfigure(1, weight=1)
app.grid_rowconfigure(4, weight=1)

# Organization Logo
logo_image = ctk.CTkImage(
    light_image=Image.open(temp_dir / "logos/nic-logo.png"),
    dark_image=Image.open(temp_dir / "logos/nic-logo.png"),
    size=(90, 90)
)

logo_label = ctk.CTkLabel(app, image=logo_image, text="")
logo_label.grid(row=0, column=0, rowspan=2, padx=50, pady=5, sticky="w")


# Header of the software : Organization Name
title = ctk.CTkLabel(
    app,
    text=lang.t("National International College"),
    font=("Segoe UI", 30, "bold")
)
title.grid(row=0, column=1, pady=5)


# Flag of the Nation
flag_label = tk.Label(app, text="")
flag_label.grid(row=0, column=2, rowspan=2, padx=50, pady=5, sticky="e")


# ========================= App Row 1 : Address =========================

# Addres of the Organization
address = ctk.CTkLabel(
    app,
    text=lang.t("India, Asia (Earth)"),
    font=("Segoe UI", 20, "normal")
)
address.grid(row=1, column=1, pady=5)


# ========================= App Row 2 : Main frame =========================

top_frame = ctk.CTkFrame(app)
top_frame.grid(row=2, column=0, columnspan=3, padx=20, pady=10, sticky="ew")
top_frame.grid_columnconfigure(6, weight=1)
top_frame_font_normal = ("Segoe UI", 15, "normal")
top_frame_font_bold = ("Segoe UI", 15, "bold")


# ---------------- Excel Selection ----------------
excel_btn = ctk.CTkButton(
    top_frame,
    width=175,
    text=lang.t("Select Excel File"),
    font=top_frame_font_bold,
    command=select_excel
)
excel_btn.grid(row=0, column=0, padx=10, pady=8, sticky="w")

excel_label = ctk.CTkLabel(
    top_frame,
    text=str(excel_path) if excel_path else "No file selected",
    font=top_frame_font_normal
)
excel_label.grid(row=0, column=1, columnspan=4, padx=10, pady=8, sticky="w")


# ---------------- Output Folder ----------------
output_btn = ctk.CTkButton(
    top_frame,
    width=175,
    text=lang.t("Change Output Folder"),
    font=top_frame_font_bold,
    command=select_output
)
output_btn.grid(row=1, column=0, padx=10, pady=8, sticky="w")

output_label = ctk.CTkLabel(
    top_frame,
    text=str(output_folder),
    font=top_frame_font_normal
)
output_label.grid(row=1, column=1, columnspan=4, padx=10, pady=8, sticky="w")


# ---------------- Evaluation Term Dropdown ----------------
term_label = ctk.CTkLabel(
    top_frame, 
    text=lang.t("Evaluation term:"),
    font=top_frame_font_bold
)
term_label.grid(row=2, column=0, padx=10, pady=8, sticky="e")

term_dropdown = ctk.CTkComboBox(
    top_frame,
    width=135,
    font=top_frame_font_normal,
    dropdown_font=top_frame_font_normal,
    values=["Select", "1st","2nd","3rd","Final"],
    state="readonly"
)

term_dropdown.set(cm.config["evaluation_term"])
term_dropdown.grid(row=2, column=1, padx=10, pady=8, sticky="w")


# ---------------- Evaluation Year Dropdown ----------------
year_label = ctk.CTkLabel(
    top_frame, 
    text=lang.t("Evaluation year:"),
    font=top_frame_font_bold
)
year_label.grid(row=2, column=2, padx=10, pady=8, sticky="e")

year_dropdown = ctk.CTkComboBox(
    top_frame,
    width=135,
    font=top_frame_font_normal,
    dropdown_font=top_frame_font_normal,
    values=["Select", "2024", "2025", "2026", "2027", "2028", "2029", "2030"],
    state="readonly"
)

year_dropdown.set(cm.config["evaluation_year"])
year_dropdown.grid(row=2, column=3, padx=10, pady=8, sticky="w")


# ---------------- Class Dropdown ----------------
class_label = ctk.CTkLabel(
    top_frame, 
    text=lang.t("Generate marksheet for:"),
    font=top_frame_font_bold
)
class_label.grid(row=3, column=0, padx=10, pady=8, sticky="e")

class_dropdown = ctk.CTkComboBox(
    top_frame,
    width=135,
    font=top_frame_font_normal,
    dropdown_font=top_frame_font_normal,
    values=[
        "All Class", "Class-1A", "Class-1B", "Class-2A", "Class-2B", "Class-3",
        "Class-4A", "Class-4B", "Class-5", "Class-6", "Class-7"
    ],
    state="readonly"
)

class_dropdown.set(cm.config["last_class"])
class_dropdown.grid(row=3, column=1, padx=10, pady=8, sticky="w")


# ---------------- Marksheet Order Dropdown ----------------
order_label = ctk.CTkLabel(
    top_frame, 
    text=lang.t("Marksheet order:"),
    font=top_frame_font_bold
)
order_label.grid(row=3, column=2, padx=10, pady=8, sticky="e")

order_dropdown = ctk.CTkComboBox(
    top_frame,
    width=135,
    font=top_frame_font_normal,
    dropdown_font=top_frame_font_normal,
    values=["Order by Rank", "Order by Roll"],
    state="readonly"
)

order_dropdown.set(cm.config["marksheet_order"])
order_dropdown.grid(row=3, column=3, padx=10, pady=8, sticky="w")


# ---------------- Output Type Dropdown ----------------
output_type = ctk.CTkLabel(
    top_frame, 
    text=lang.t("Output type:"),
    font=top_frame_font_bold
)
output_type.grid(row=3, column=4, padx=10, pady=8, sticky="e")

output_type_dropdown = ctk.CTkComboBox(
    top_frame,
    width=135,
    font=top_frame_font_normal,
    dropdown_font=top_frame_font_normal,
    values=["Single PDF", "Separate PDFs"],
    state="readonly"
)

output_type_dropdown.set(cm.config["output_type"])
output_type_dropdown.grid(row=3, column=5, padx=10, pady=8, sticky="w")


# ---------------- Generate Button ----------------
generate_btn = ctk.CTkButton(
    top_frame,
    height=40,
    width=165,
    text=lang.t("Generate Marksheets"),
    font=top_frame_font_bold,
    command=start_main
)
generate_btn.grid(row=4, column=0, columnspan=7, pady=20)


# ========================= App Row 3 : Progress frame =========================

progress_frame = ctk.CTkFrame(app)
progress_frame.grid(row=3, column=0, columnspan=3,
                    padx=20, pady=5, sticky="ew")
progress_frame.grid_columnconfigure(1, weight=1)

# ---------------- Progress Bar ----------------
progress_label = ctk.CTkLabel(
    progress_frame, 
    text=lang.t("Progress bar:"),
    font=top_frame_font_normal
)
progress_label.grid(row=3, column=0, padx=10, pady=5, sticky="e")

progressbar = ctk.CTkProgressBar(progress_frame, height=15)
progressbar.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
progressbar.set(0)


# ========================= App Row 4 : Log frame =========================

# ---------------- Log Box ----------------
log_box = ctk.CTkTextbox(app, height=120, font=top_frame_font_normal)
log_box.configure(state="disabled")  # Make log box read-only
log_box.grid(row=4, column=0, columnspan=3, padx=20, pady=10, sticky="nsew")

# Load fresh handlers
if logger.hasHandlers:
    logger.handlers.clear()

# GUI handler
gui_handler = CTkinterHandler(log_box)
gui_handler.setFormatter(
    logging.Formatter(
        "[%(asctime)s] %(levelname)s: %(message)s",
        datefmt="%d-%m-%Y  %H:%M:%S"
    )
)
logger.addHandler(gui_handler)

# Console handler
console_handler = logging.StreamHandler()
console_handler.setFormatter(
    logging.Formatter("%(levelname)s: %(message)s")
)
logger.addHandler(console_handler)

# File handler
file_handler = logging.FileHandler(base_dir / "app.log", encoding="utf-8")
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(
    logging.Formatter(
        "[%(asctime)s] %(levelname)s: %(message)s",
        datefmt="%d-%m-%Y  %H:%M:%S"
    )
)
logger.addHandler(file_handler)


# ========================= App Row 5 : Status frame =========================

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
status_frame.grid_columnconfigure(1, weight=0)  # language dropdown
status_frame.grid_columnconfigure(2, weight=0)  # theme toggle
status_frame.grid_columnconfigure(3, weight=0)  # time

# The current status of app
status_label = ctk.CTkLabel(
    status_frame, 
    text=lang.t("Status: Ready"),
    text_color="white",
    font=status_bar_font
)
status_label.grid(row=0, column=0, padx=20, sticky="w")

# Change language options
language_btn = ctk.CTkSegmentedButton(
    status_frame,
    corner_radius=0,
    text_color="white",
    font=status_bar_font,
    fg_color="#007ACC",
    selected_color="#005a9e",
    unselected_color="#007ACC",
    selected_hover_color="#004578",
    unselected_hover_color="#006bb3",
    values=["हिंदी", "English"],
    command=change_language
)
language_btn.set(cm.config["language"])
language_btn.grid(row=0, column=1, padx=20, pady=1, sticky="e")


# Theme toggle buttons
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
theme_btn.grid(row=0, column=2, padx=20, pady=1, sticky="e")

# Show day, date and time
time_label = ctk.CTkLabel(
    status_frame,
    text="",
    text_color="white",
    font=status_bar_font
)
time_label.grid(row=0, column=3, padx=20, sticky="e")


# Maximize the window after 100 ms
app.after(100, lambda: app.state("zoomed"))
update_time()   # Load day, date and time
animate_flag()  # Load national flag
refresh_ui(cm.config["language"])    # Display the last used language

app.protocol("WM_DELETE_WINDOW", on_closing)    # While closing the window
atexit.register(log_exit)                       # Log before exit

logger.info("Application started successfully")

app.mainloop()
