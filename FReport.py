import time
import os
from datetime import datetime, timedelta
import threading

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog
import logging

import dataDownload 

class FormPage(tb.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, padding=20)
        self.controller = controller
    
        # Variables
        self.PDBSusername_var = tb.StringVar()
        self.PDBSpassword_var = tb.StringVar()
        self.folder_path_var = tb.StringVar()

        # Widgets

        self.pdbsUserText = tb.Label(self, 
            text="PDBS Username:", 
            font=("Segoe UI", 12))
        self.pdbsUserText.grid(row=3, column=0, sticky=W, pady=5, padx=(0, 10))

        self.pdbsUserForm = tb.Entry(self, 
            textvariable=self.PDBSusername_var, 
            font=("Segoe UI", 12))
        self.pdbsUserForm.grid(row=3, column=1, pady=5, sticky=EW)

        self.pdbsPassText = tb.Label(self, 
            text="PDBS Password:", 
            font=("Segoe UI", 12))
        self.pdbsPassText.grid(row=4, column=0, sticky=W, pady=5)

        self.pdbsPassForm = tb.Entry(self, 
            textvariable=self.PDBSpassword_var, 
            font=("Segoe UI", 12), show="*")
        self.pdbsPassForm.grid(row=4, column=1, pady=5, sticky=EW, ipadx=(25))

        self.folderText = tb.Label(self, 
            text="Folder Path:", 
            font=("Segoe UI", 12))
        self.folderText.grid(row=5, column=0, sticky=W, pady=5)

        self.folderForm = tb.Entry(self, 
            textvariable=self.folder_path_var, 
            font=("Segoe UI", 12))
        self.folderForm.grid(row=5, column=1, pady=5, sticky=EW)

        self.browse_btn = tb.Button(self, 
            text="Browse...", 
            command=self.browse_folder)
        self.browse_btn.grid(row=5, column=2, padx=10, pady=5)

        self.submit_btn = tb.Button(self, 
            text="Submit", 
            bootstyle=SUCCESS, 
            command=self.submit)
        self.submit_btn.grid(row=6, column=1, pady=20)

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

            # Switch to loading page
            self.controller.show_frame("SecondPage")

            # Start DailyOS in background
            threading.Thread(
                target=dataDownload.DailyOS,
                args=(
                    PDBSusername,
                    PDBSpassword,
                    folder_path,
                    self.controller.frames["SecondPage"].progress_control
                ),
                daemon=True
            ).start()

        except Exception as e:
            logging.error("An error occurred", exc_info=True)

class SecondPage(tb.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, padding=20)
        self.controller = controller

        tb.Label(self, 
            text="Processing your request...", 
            font=("Segoe UI", 14)).pack(pady=10)

        # Status label
        self.status_label = tb.Label(self, 
            text="Waiting to start...", 
            bootstyle=SECONDARY)
        self.status_label.pack(pady=5)

        # Progress bar (indeterminate mode)
        self.progress = tb.Progressbar(self, 
            mode="indeterminate", 
            bootstyle=INFO)
        self.progress.pack(fill=X, padx=20, pady=10)

        # Back button
        tb.Button(
            self,
            text="Update Report (Excel)",
            bootstyle=SECONDARY,
            width=20,
            command=lambda: controller.show_frame("FormPage")
        ).pack(pady=20)

    def progress_control(self, action):
        """Called by DailyOS to control the loading bar"""
        if action == "start":
            self.status_label.config(
                text="Loading...", 
                bootstyle=INFO)
            self.progress.start(10)
        elif action == "stop":
            self.progress.stop()
            self.status_label.config(
                text="Done!", 
                bootstyle=SUCCESS)

class App(tb.Window):
    def __init__(self):
        super().__init__(themename="flatly")
        self.title("Fulfillment Report Data Downloader")
        self.geometry("500x250+700+300")

        container = tb.Frame(self)
        container.pack(fill=BOTH, expand=YES)
        
        self.frames = {}

        for F in (FormPage, SecondPage):
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