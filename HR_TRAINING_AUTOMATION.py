import os
import sys
import time
import subprocess
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.scrolledtext import ScrolledText

# ==========================================================
# HR TRAINING AUTOMATION - RUN ALL 3 SCRIPTS IN ORDER
# Keep this launcher in the same folder as the 3 scripts.
# ==========================================================

WAIT_SECONDS = 5

SCRIPT_NEW = "QR_scan_new.py"
SCRIPT_UPDATE = "QR_scan_update.py"
SCRIPT_IMPORT = "HR_IMPORT2_V5.4_GUI_v3.py"
SCRIPT_FINAL = "HR_SCRIPT1_V8.py"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Use python.exe even when this launcher is started as .pyw,
# so the child script's stdout/stderr can be captured.
PYTHON_EXE = sys.executable
if os.path.basename(PYTHON_EXE).lower() == "pythonw.exe":
    PYTHON_EXE = os.path.join(os.path.dirname(PYTHON_EXE), "python.exe")


REQUIREMENTS_FILE = os.path.join(BASE_DIR, "requirements.txt")


def add_log(message):
    """Append one line to the GUI log safely from the worker thread."""
    root.after(0, _add_log_main, message)


def _add_log_main(message):
    log_text.config(state="normal")
    log_text.insert("end", message.rstrip() + "\n")
    log_text.see("end")
    log_text.config(state="disabled")


def install_requirements():
    """Check/install packages from requirements.txt and stream pip output to the GUI log."""
    if not os.path.isfile(REQUIREMENTS_FILE):
        raise FileNotFoundError(
            f"requirements.txt not found in the launcher folder:\n{REQUIREMENTS_FILE}"
        )

    add_log("")
    add_log("=" * 68)
    add_log("Checking required Python libraries...")
    add_log(f"Requirements file: {REQUIREMENTS_FILE}")
    add_log("=" * 68)

    process = subprocess.Popen(
        [
            PYTHON_EXE, "-m", "pip", "install",
            "-r", REQUIREMENTS_FILE,
            "--disable-pip-version-check",
        ],
        cwd=BASE_DIR,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )

    if process.stdout is not None:
        for line in process.stdout:
            add_log(line)
        process.stdout.close()

    return_code = process.wait()

    if return_code != 0:
        add_log(f"Library installation FAILED (exit code {return_code})")
        raise RuntimeError(
            f"Could not install required Python libraries "
            f"(exit code {return_code})."
        )

    add_log("Required Python libraries are ready.")
    add_log("")


def run_script(script_name):
    """Run one child script, show its live output, and wait for it to finish."""
    script_path = os.path.join(BASE_DIR, script_name)

    if not os.path.isfile(script_path):
        raise FileNotFoundError(f"Script not found:\n{script_path}")

    add_log(f"\n{'=' * 68}")
    add_log(f"Starting: {script_name}")
    add_log(f"{'=' * 68}")

    process = subprocess.Popen(
        [PYTHON_EXE, "-u", script_path],
        cwd=BASE_DIR,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )

    if process.stdout is not None:
        for line in process.stdout:
            add_log(line)
        process.stdout.close()

    return_code = process.wait()

    if return_code != 0:
        add_log(f"{script_name} FAILED (exit code {return_code})")
        raise RuntimeError(
            f"{script_name} failed with exit code {return_code}."
        )

    add_log(f"{script_name} completed successfully.")

    # Small buffer between scripts for Excel / file sync / file closing.
    if WAIT_SECONDS > 0:
        add_log(f"Waiting {WAIT_SECONDS} seconds before the next step...")
        time.sleep(WAIT_SECONDS)


def run_pipeline(update_existing=False):
    """Run all 3 scripts in order in a background thread."""
    first_script = SCRIPT_UPDATE if update_existing else SCRIPT_NEW

    steps = [
        ("Step 1 of 3", first_script),
        ("Step 2 of 3", SCRIPT_IMPORT),
        ("Step 3 of 3", SCRIPT_FINAL),
    ]

    try:
        root.after(
            0,
            lambda: update_progress(
                "Preparing",
                "requirements.txt",
                "Checking/installing required libraries..."
            ),
        )
        install_requirements()

        for step_text, script_name in steps:
            root.after(0, lambda s=step_text, n=script_name:
                       update_progress(s, n, "Running..."))
            run_script(script_name)
            root.after(0, lambda s=step_text:
                       update_progress(s, "", "Completed"))

        root.after(0, pipeline_completed)

    except Exception as e:
        add_log(f"\nERROR: {e}")
        root.after(0, lambda err=e: pipeline_failed(err))


def update_progress(step_text, script_name, status):
    progress_step.config(text=step_text)
    progress_script.config(text=script_name)
    progress_status.config(text=status)

    if status == "Completed":
        progress_bar["value"] = min(
            progress_bar["value"] + 33.33, 100
        )
    else:
        progress_bar["value"] = {
            "Preparing": 5,
            "Step 1 of 3": 15,
            "Step 2 of 3": 48,
            "Step 3 of 3": 80,
        }.get(step_text, 0)

    root.update_idletasks()


def pipeline_completed():
    progress_bar["value"] = 100
    progress_step.config(text="All steps completed")
    progress_script.config(text="Training record generated successfully")
    progress_status.config(text="Ready")
    _add_log_main("\nAll 3 steps completed successfully.")
    close_button.config(state="normal", text="Close")
    close_button.config(command=root.destroy)

    messagebox.showinfo(
        "Completed",
        "HR training processing completed successfully.\n\n"
        "All 3 scripts finished in order."
    )


def pipeline_failed(error):
    progress_status.config(text="Failed")
    progress_script.config(text="Processing stopped")
    close_button.config(state="normal", text="Close")
    close_button.config(command=root.destroy)

    messagebox.showerror(
        "HR Training Automation",
        str(error)
    )


def start_pipeline(update_existing):
    global selected_mode

    selected_mode = "Update existing" if update_existing else "Create new"
    new_button.config(state="disabled")
    update_button.config(state="disabled")

    show_progress()
    _add_log_main(
        f"Mode: {selected_mode}\n"
        "The required libraries will be checked first, then the 3 scripts will run automatically in sequence.\n"
    )

    threading.Thread(
        target=run_pipeline,
        args=(update_existing,),
        daemon=True
    ).start()


def show_progress():
    selection_frame.pack_forget()
    progress_frame.pack(fill="both", expand=True)


def choose_mode():
    new_button.config(command=lambda: start_pipeline(False))
    update_button.config(command=lambda: start_pipeline(True))


# ==========================================================
# GUI
# ==========================================================

root = tk.Tk()
root.title("HR Training Automation")
root.geometry("700x500")
# root.minsize(650, 600)
root.resizable(True, True)
root.configure(bg="#F4F7FA")

# ---------- ttk styling ----------
style = ttk.Style()
try:
    style.theme_use("clam")
except tk.TclError:
    pass

style.configure(
    "TButton",
    font=("Segoe UI", 10, "bold"),
    padding=(18, 13)
)

style.configure(
    "Horizontal.TProgressbar",
    troughcolor="#E2E8F0",
    background="#1F7A8C",
    bordercolor="#E2E8F0",
    lightcolor="#1F7A8C",
    darkcolor="#1F7A8C"
)

# ---------- Header ----------
header = tk.Frame(root, bg="#17324D", height=92)
header.pack(fill="x")
header.pack_propagate(False)

tk.Label(
    header,
    text="HR Training Automation",
    bg="#17324D",
    fg="white",
    font=("Segoe UI", 20, "bold")
).pack(anchor="w", padx=28, pady=(18, 2))

tk.Label(
    header,
    text="Training Evidence → Training Record",
    bg="#17324D",
    fg="#D7E6F2",
    font=("Segoe UI", 10)
).pack(anchor="w", padx=30)

# ---------- Main container ----------
main = tk.Frame(root, bg="#F4F7FA")
main.pack(fill="both", expand=True, padx=22, pady=20)

# ---------- Selection frame ----------
selection_frame = tk.Frame(main, bg="#F4F7FA")
selection_frame.pack(fill="both", expand=True)

tk.Label(
    selection_frame,
    text="Select processing mode",
    bg="#F4F7FA",
    fg="#17324D",
    font=("Segoe UI", 15, "bold")
).pack(anchor="w", pady=(2, 4))

tk.Label(
    selection_frame,
    text="Choose how you want to process the Training Evidence file.",
    bg="#F4F7FA",
    fg="#5B6875",
    font=("Segoe UI", 10)
).pack(anchor="w", pady=(0, 16))


def make_mode_card(parent, title, description, color):
    card = tk.Frame(
        parent,
        bg="white",
        highlightbackground="#D9E1E8",
        highlightthickness=1,
        bd=0
    )
    card.pack(fill="x", pady=7)

    accent = tk.Frame(card, bg=color, width=7)
    accent.pack(side="left", fill="y")

    content = tk.Frame(card, bg="white")
    content.pack(side="left", fill="both", expand=True, padx=15, pady=13)

    tk.Label(
        content,
        text=title,
        bg="white",
        fg="#17324D",
        font=("Segoe UI", 11, "bold")
    ).pack(anchor="w")

    tk.Label(
        content,
        text=description,
        bg="white",
        fg="#66727D",
        font=("Segoe UI", 9)
    ).pack(anchor="w", pady=(3, 0))

    return card


new_card = make_mode_card(
    selection_frame,
    "Create NEW Training Evidence",
    "Scan all training PDFs and generate a new evidence workbook.",
    "#1F7A8C"
)

update_card = make_mode_card(
    selection_frame,
    "UPDATE Existing Training Evidence",
    "Process new PDFs and update the existing evidence workbook.",
    "#6C63A8"
)

new_button = ttk.Button(new_card, text="Start  →")
new_button.pack(side="right", padx=16, pady=20)

update_button = ttk.Button(update_card, text="Start  →")
update_button.pack(side="right", padx=16, pady=20)

tk.Label(
    selection_frame,
    text="The 3 scripts will run automatically in sequence.",
    bg="#F4F7FA",
    fg="#7A8793",
    font=("Segoe UI", 9, "italic")
).pack(anchor="w", pady=(14, 0))

# ---------- Progress frame ----------
progress_frame = tk.Frame(main, bg="#F4F7FA")

progress_step = tk.Label(
    progress_frame,
    text="Step 1 of 3",
    bg="#F4F7FA",
    fg="#17324D",
    font=("Segoe UI", 15, "bold")
)
progress_step.pack(anchor="w", pady=(2, 4))

progress_script = tk.Label(
    progress_frame,
    text="",
    bg="#F4F7FA",
    fg="#5B6875",
    font=("Segoe UI", 10)
)
progress_script.pack(anchor="w", pady=(0, 14))

progress_bar = ttk.Progressbar(
    progress_frame,
    orient="horizontal",
    mode="determinate",
    length=510,
    style="Horizontal.TProgressbar"
)
progress_bar.pack(fill="x", pady=(0, 16))

progress_status = tk.Label(
    progress_frame,
    text="Waiting...",
    bg="#F4F7FA",
    fg="#1F7A8C",
    font=("Segoe UI", 10, "bold")
)
progress_status.pack(anchor="w")

tk.Label(
    progress_frame,
    text="Please wait until all 3 steps are completed.",
    bg="#F4F7FA",
    fg="#7A8793",
    font=("Segoe UI", 9)
).pack(anchor="w", pady=(5, 8))

# ---------- Live log window ----------
log_header = tk.Frame(progress_frame, bg="#F4F7FA")
log_header.pack(fill="x", pady=(4, 5))

tk.Label(
    log_header,
    text="Process log",
    bg="#F4F7FA",
    fg="#17324D",
    font=("Segoe UI", 10, "bold")
).pack(side="left")

tk.Label(
    log_header,
    text="Live output from the running scripts",
    bg="#F4F7FA",
    fg="#7A8793",
    font=("Segoe UI", 9)
).pack(side="right")

log_text = ScrolledText(
    progress_frame,
    height=14,
    wrap="word",
    font=("Consolas", 9),
    bg="#17212B",
    fg="#E7EDF2",
    insertbackground="white",
    relief="flat",
    bd=0,
    padx=10,
    pady=8
)
log_text.pack(fill="both", expand=True, pady=(0, 2))
log_text.config(state="disabled")

close_button = ttk.Button(
    progress_frame,
    text="Close",
    state="disabled",
    command=root.destroy
)
close_button.pack(anchor="e", pady=(12, 0))

selected_mode = ""
choose_mode()

root.mainloop()
