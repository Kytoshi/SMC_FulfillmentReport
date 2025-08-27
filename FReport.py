import time
import os
import threading
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog
import logging
import sys
from tkinter import PhotoImage

import dataDownload  # Your updated DailyOS
import updateReport  # Your updated update_report

class FormPage(tb.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, padding=20)
        self.controller = controller

        # Variables
        self.PDBSusername_var = tb.StringVar()
        self.PDBSpassword_var = tb.StringVar()
        self.folder_path_var = tb.StringVar()

        # Widgets
        self.pdbsUserText = tb.Label(
            self,
            text="PDBS Username:",
            font=("Segoe UI", 12)
        )
        self.pdbsUserText.grid(
            row=3,
            column=0,
            sticky=W,
            pady=5,
            padx=(0, 10)
        )

        self.pdbsUserForm = tb.Entry(
            self,
            textvariable=self.PDBSusername_var,
            font=("Segoe UI", 12)
        )
        self.pdbsUserForm.grid(
            row=3,
            column=1,
            pady=5,
            sticky=EW
        )

        self.pdbsPassText = tb.Label(
            self,
            text="PDBS Password:",
            font=("Segoe UI", 12)
        )
        self.pdbsPassText.grid(
            row=4,
            column=0,
            sticky=W,
            pady=5
        )

        self.pdbsPassForm = tb.Entry(
            self,
            textvariable=self.PDBSpassword_var,
            font=("Segoe UI", 12),
            show="*"
        )
        self.pdbsPassForm.grid(
            row=4,
            column=1,
            pady=5,
            sticky=EW,
            ipadx=(25)
        )

        self.folderText = tb.Label(
            self,
            text="Folder Path:",
            font=("Segoe UI", 12)
        )
        self.folderText.grid(
            row=5,
            column=0,
            sticky=W,
            pady=5
        )

        self.folderForm = tb.Entry(
            self,
            textvariable=self.folder_path_var,
            font=("Segoe UI", 12)
        )
        self.folderForm.grid(
            row=5,
            column=1,
            pady=5,
            sticky=EW
        )

        self.browse_btn = tb.Button(
            self,
            text="Browse...",
            command=self.browse_folder
        )
        self.browse_btn.grid(
            row=5,
            column=2,
            padx=10,
            pady=5
        )

        self.submit_btn = tb.Button(
            self,
            text="Submit",
            bootstyle=SUCCESS,
            width=20,
            command=self.submit
        )
        self.submit_btn.grid(
            row=6,
            column=1,
            pady=20
        )

        self.columnconfigure(1, weight=1)

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path_var.set(folder_selected)

    def submit(self):
        try:
            PDBSusername = self.PDBSusername_var.get()
            PDBSpassword = self.PDBSpassword_var.get()
            folder_path = os.path.normpath(self.folder_path_var.get())

            # Switch to SecondPage
            self.controller.show_frame("SecondPage")

            # Start DailyOS in background
            threading.Thread(
                target=dataDownload.DailyOS,
                args=(
                    PDBSusername,
                    PDBSpassword,
                    folder_path,
                    self.controller.frames["SecondPage"].progress_control,
                ),
                daemon=True
            ).start()

        except Exception as e:
            logging.error("An error occurred", exc_info=True)

class SecondPage(tb.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, padding=20)
        self.controller = controller

        self.title_label = tb.Label(
            self,
            text="Downloading Data From PDBS",
            font=("Segoe UI", 14)
        )
        self.title_label.grid(
            row=0,
            column=0,
            columnspan=3,
            pady=10
        )

        self.status_label = tb.Label(
            self,
            text="Waiting to start...",
            font=("Segoe UI", 12),
            bootstyle=SECONDARY
        )
        self.status_label.grid(
            row=1,
            column=0,
            columnspan=3,
            pady=5
        )

        self.progress = tb.Progressbar(
            self,
            mode="determinate",
            bootstyle=INFO,
            maximum=100
        )
        self.progress.grid(
            row=2,
            column=0,
            columnspan=3,
            sticky=EW,
            padx=20,
            pady=10
        )

        self.continue_btn = tb.Button(
            self,
            text="Update Report",
            bootstyle=SUCCESS,
            width=25,
            command=self.go_to_third
        )
        self.continue_btn.grid(
            row=3,
            column=0,
            columnspan=3,
            pady=20
        )
        self.continue_btn.grid_remove()  # Hide initially

        self.columnconfigure(1, weight=1)

    def go_to_third(self):
        self.controller.show_frame("ThirdPage")
        folder_path = os.path.normpath(self.controller.frames["FormPage"].folder_path_var.get())
        threading.Thread(
            target=updateReport.update_report,
            args=(
                folder_path,
                False,
                self.controller.frames["ThirdPage"].progress_control,
            ),
            daemon=True
        ).start()

    def progress_control(self, action, value=0):
        if action == "start":
            self.status_label.config(
                text="Loading...",
                bootstyle=INFO
            )
            self.progress['value'] = 0
        elif action == "update":
            self.progress['value'] = value
            self.status_label.config(
                text=f"Loading... {int(value)}%",
                bootstyle=INFO
            )
        elif action == "stop":
            self.progress['value'] = 100
            self.status_label.config(
                text="Done!",
                bootstyle=SUCCESS
            )
            self.continue_btn.grid()  # Show button

class ThirdPage(tb.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, padding=20)
        self.controller = controller

        self.title_label = tb.Label(
            self,
            text="Updating Tables in Report",
            font=("Segoe UI", 14)
        )
        self.title_label.grid(
            row=0,
            column=0,
            columnspan=3,
            pady=10
        )

        self.status_label = tb.Label(
            self,
            text="Waiting to start...",
            font=("Segoe UI", 12),
            bootstyle=SECONDARY
        )
        self.status_label.grid(
            row=1,
            column=0,
            columnspan=3,
            pady=5
        )

        self.progress = tb.Progressbar(
            self,
            mode="determinate",
            bootstyle=INFO,
            maximum=100
        )
        self.progress.grid(
            row=2,
            column=0,
            columnspan=3,
            sticky=EW,
            padx=20,
            pady=10
        )

        self.done_btn = tb.Button(
            self,
            text="Back to Form Page",
            bootstyle=SECONDARY,
            width=20,
            command=lambda: controller.show_frame("FormPage")
        )
        self.done_btn.grid(
            row=3,
            column=0,
            columnspan=3,
            pady=20
        )
        self.done_btn.grid_remove()  # Hide initially

        self.columnconfigure(1, weight=1)

    def progress_control(self, action, value=0):
        if action == "start":
            self.status_label.config(
                text="Updating report...",
                bootstyle=INFO
            )
            self.progress['value'] = 0
        elif action == "update":
            self.progress['value'] = value
            self.status_label.config(
                text=f"Updating... {int(value)}%",
                bootstyle=INFO
            )
        elif action == "stop":
            self.progress['value'] = 100
            self.status_label.config(
                text="Update complete!",
                bootstyle=SUCCESS
            )
            self.done_btn.grid()  # Show button

class App(tb.Window):
    def __init__(self):
        super().__init__(themename="flatly")
        self.title("Fulfillment Report Data Downloader")
        self.geometry("500x250+700+300")

        # --- Set custom window icon ---
        icon_filename = "FulfillmentRptIcon.ico"  # Change to your icon file name

        # Determine path correctly for PyInstaller or normal script
        if getattr(sys, 'frozen', False):
            # PyInstaller executable
            base_path = sys._MEIPASS
        else:
            # Normal script
            base_path = os.path.dirname(os.path.abspath(__file__))

        icon_path = os.path.join(base_path, icon_filename)

        # Try .ico first for Windows
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception as e:
                print(f"Failed to set iconbitmap: {e}")

        # Fallback for cross-platform using .png
        png_icon = os.path.join(base_path, "FulfillmentRptPNG.png")
        if os.path.exists(png_icon):
            try:
                self.iconphoto(True, PhotoImage(file=png_icon))
            except Exception as e:
                print(f"Failed to set iconphoto: {e}")

        # --- Continue setting up your UI ---
        container = tb.Frame(self)
        container.pack(fill="both", expand=True)

        self.frames = {}
        for F in (FormPage, SecondPage, ThirdPage):
            page_name = F.__name__
            frame = F(container, self)
            frame.grid(row=0, column=0, sticky="nsew")
            self.frames[page_name] = frame

        self.show_frame("FormPage")

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()

if __name__ == "__main__":
    app = App()
    app.mainloop()
