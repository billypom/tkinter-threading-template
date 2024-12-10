import sys
from os import path, getcwd
from threading import Thread
from platform import uname
from tempfile import gettempdir
from subprocess import run as sprun
from logging import basicConfig, warning, INFO
from tkinter import (
    Tk,
    Label,
    Menu,
    messagebox,
    Button,
    filedialog,
    font,
    Toplevel,
    scrolledtext,
    WORD,
    END,
    DISABLED,
    BOTH,
)

# CONFIGURATIONS

PRODUCT_NAME = ""
SOFTWARE_VERSION = "0.0.1"
LINUX_INSTALLER_PATH = ""
WINDOWS_INSTALLER_PATH = ""

def in_wsl() -> bool:
    """Windows vs not windows"""
    return "microsoft-standard" in uname().release

if not in_wsl():
    from os import startfile # type: ignore | Windows only

if in_wsl():
    basicConfig(
        format="%(asctime)s %(levelname)-8s %(message)s",
        filename="log.txt",
        filemode="a",
        level=INFO,
    )
else:
    basicConfig(
        format="%(asctime)s %(levelname)-8s %(message)s",
        filename=f"{gettempdir()}/{PRODUCT_NAME}-log.txt",
        filemode="a",
        level=INFO,
    )

# FUNCTIONS

def get_product_version() -> str:
    """Returns this software version as a string"""
    return SOFTWARE_VERSION

def get_latest_version() -> str:
    """Returns the latest software version, as defined in the repo, as a string"""
    linux_repo_path = ""
    windows_repo_path = ""
    if in_wsl():
        file = linux_repo_path
    else:
        file = windows_repo_path
    with open(file, "r") as f:
        content = f.read()
    return content.strip("\n")

def resource_path(relative_path) -> str:
    """Dynamically loads any imported files from the root directory
    e.g., resource_path('license.txt') used to refer to ./license.txt"""
    try:
        base_path = sys._MEIPASS  # type: ignore
    except Exception:
        base_path = path.abspath(".")
    return path.join(base_path, relative_path)

def show_about_menu():
    """About menu"""
    messagebox.showinfo(
        title="About",
        message="Details about this program",
    )

def show_license_menu():
    """Displays text from a list of .txt files"""
    license_window = Toplevel()
    license_window.title("License")
    license_window.geometry("800x640")
    text_area = scrolledtext.ScrolledText(
        license_window, wrap=WORD, width=80, height=20
    )
    text_area.pack(padx=10, pady=10, fill=BOTH, expand=True)
    license_files = [
        resource_path("license.txt"),
    ]
    for filename in license_files:
        try:
            with open(filename, "r", encoding="utf-8") as file:
                text_area.insert(END, file.read() + "\n\n")
        except FileNotFoundError:
            text_area.insert(END, f"ERROR: {filename} not found\n\n")
    text_area.config(state=DISABLED)

def check_for_updates() -> bool:
    """
    Compares the current software version to the latest version listed in the repository
    Opens the file explorer at the directory of the installer if a new version is available

    Returns true if user intends to update
    Returns false if already updated or user chooses not to update
    """
    current_version = get_product_version()
    latest_version = get_latest_version()
    try:
        if current_version != latest_version:
            result = messagebox.askyesno(
                "Check for updates",
                f"Update to {latest_version} is available.\nWould you like to go to the installer?",
            )
            if result:
                if not in_wsl():
                    startfile(WINDOWS_INSTALLER_PATH)
                else:
                    sprun(["explorer.exe", LINUX_INSTALLER_PATH]
                    )
                return True
            else:
                return False
        return False
    except Exception as e:
        warning(f"check_for_updates error: {e}")
        return False

def pprint(message):
    """
    Prints a message to the console and updates the log_label in the GUI with the same message.

    Args:
    message (str): The message to print and display in the GUI.
    """
    print(message)  # Print to console
    if log_label:  # Check if log_label is defined to avoid NameError
        log_label.config(text=message)  # Update GUI log_label text

def add_start_button():
    """
    Adds a [start] button to to the GUI that begins some process
    - run this function in display_gui() before root.mainloop()
    """
    # Load the button to open a file, or automatically process the input file from sys.argv
    button_font = font.Font(family="Calibri", size=20)
    open_file_button = Button(
        root,
        text="Start",
        font=button_font,
        bg="#a1ffc4",
        activebackground="#1151af",
        borderwidth=5,
        command=lambda: [start_processing],
    )
    open_file_button.grid(row=4, column=0)

def add_open_excel_file_button():
    """
    Adds an [open file] button to the GUI, if we did not pass in CLI argument

    - run this function in display_gui() before root.mainloop()
    - .xls files only
    """
    # Load the button to open a file, or automatically process the input file from sys.argv
    button_font = font.Font(family="Calibri", size=20)
    if not INPUT_FILE:
        open_file_button = Button(
            root,
            text="Open file",
            font=button_font,
            bg="#a1ffc4",
            activebackground="#1151af",
            borderwidth=5,
            command=lambda: [dialog_open_excel_file(start_processing)],
        )
        open_file_button.grid(row=4, column=0)
    else:
        dialog_open_excel_file(start_processing)

def dialog_open_excel_file(callback_function):
    """
    Opens a dialog window to choose a .xls file, then runs the callback function
    If a CLI argument was passed, immediately run the callback function
    """
    global INPUT_FILE
    if not INPUT_FILE:
        INPUT_FILE = filedialog.askopenfilename(
            initialdir="",
            title="Select file",
            filetypes=(("Excel 97-03", "*.xls"), ("All files", "*.*")),
        )
    if INPUT_FILE and INPUT_FILE.endswith(".xls"):
        callback_function()  # Call the callback with the selected file
    else:
        messagebox.showerror(
            title="Error: wrong filetype",
            message="This program can only accept *.xls files (Excel 97-03 format)",
        )

def display_gui(root):
    """Displays the tkinter gui"""
    # Menu
    menubar = Menu(root)
    helpmenu = Menu(menubar, tearoff=0)
    helpmenu.add_command(label="About", command=show_about_menu)
    helpmenu.add_command(label="License", command=show_license_menu)
    helpmenu.add_command(label="Check for updates", command=check_for_updates)
    menubar.add_cascade(label="File", menu=helpmenu)
    root.config(menu=menubar)

    # Fonts
    log_font = font.Font(family="Latin Modern Mono", size=12)
    generic_font = font.Font(family="Calibri", size=18)

    # Labels
    global title_label  # Title label
    title_label = Label(root, text="tkinter threading template", font=generic_font)
    title_label.grid(row=1, column=0, padx=3, pady=3)
    global log_label  # Label updated by pprint()
    log_label = Label(root, text="", font=log_font)
    log_label.grid(row=3, column=0, padx=10, pady=5)
    add_start_button()
    root.mainloop()


def start_processing():
    """Starts the animations & runs the work in a new thread"""
    global INPUT_FILE
    # Start thread for processing file - on function completion stop animation
    processing_thread = Thread(
        target=lambda: [work()],
        daemon=True,
    )
    processing_thread.start()
    check_threads([processing_thread])

def check_threads(threads):
    """Check if all threads are finished and run a follow up function"""
    if all(not thread.is_alive() for thread in threads):
        after_threads_complete()
    else:
        # Schedule to check again after 100ms
        root.after(100, check_threads, threads)

def after_threads_complete():
    """Runs after all threads complete"""
    pprint(f"All threads complete\nThis program will close after 10 seconds")
    # Schedule root.destroy() to be called on main thread
    root.after(10000, root.destroy)  # 10,000ms = 10 seconds to wait before destroying

def work():
    """Where we do our processing"""
    for i in range(1000):
        pprint(i)

def main():
    global root
    global INPUT_FILE
    global TEMP_DIR
    root = Tk()
    root.title("Title")

    # Input file determination (allow drag&drop or open file button)
    try:
        INPUT_FILE = path.abspath(sys.argv[1])
    except Exception:
        INPUT_FILE = ""

    # Windows desktop icon - pyinstaller
    if not in_wsl():
        TEMP_DIR = gettempdir()
        # https://stackoverflow.com/questions/45628653/add-ico-file-to-executable-in-pyinstaller
        # Makes the tiny icon work
        if getattr(sys, "frozen", False):
            application_path = sys._MEIPASS # type: ignore
        elif __file__:
            application_path = path.dirname(__file__)
        else:
            return
        icon_file = "icon.ico"
        root.iconbitmap(default=path.join(application_path, icon_file))
    else:
        TEMP_DIR = getcwd()

    display_gui(root)

    # Close program
    try:
        root.quit()
    except Exception as e:
        print("__root quit error: ", str(e))
    return

if __name__ == '__main__':
    main()
