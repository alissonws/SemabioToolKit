#!/usr/bin/env python3
# -*- coding: utf-8 -*-

##LAUNCHER
##Checks for updates and start GUI

import logging, os, sys, ctypes
import tkinter as tk
import tkinter.ttk as ttk
import platform
from multiprocessing import Queue, Process, freeze_support
from utils import autoupdater


TITLE = "XX Semana Acadêmica da Biologia"
APP_NAME = 'Semabio ToolKit'
APP_VERSION = '0.1.1'
PLATFORM = platform.system()
LAUNCHER_PATH = os.path.dirname(os.path.abspath(__file__))


#Logging configuration
if "sys.argv" in locals():
    if sys.argv[1] == "--debug" or sys.argv[1] == "-d":
        logging.basicConfig(level = logging.DEBUG,format='%(process)d-%(levelname)s-%(message)s')

class Launcher:
    def __init__(self, master):

        self.queue_progress = Queue()

        logging.debug("Generating GUI...")

        #STYLES FOR TTK
        style = ttk.Style()
        #Chosing a new theme
        style.theme_use("clam")

        #Title labels
        style.configure("Main.TLabel", background="white",
                        font=("Noto Sans CJK JP Regular",12,"bold"),foreground="black")
        #Common labels
        style.configure("General.TLabel", background="white",
                        font=("Chandas",11),foreground="black")

        #Title

        self.title_frame = tk.Frame(bg="white")
        self.title_frame.grid(stick="nw")

        self.title_label = ttk.Label(self.title_frame, text="Procurando atualizações...", style="General.TLabel")
        self.title_label.grid(padx=10,pady=10)

        #Progressbar

        self.launcher_progressbar_frame = tk.Frame(bg="white")
        self.launcher_progressbar_frame.grid(row=1,sticky="nw")

        self.launcher_progressbar = ttk.Progressbar(self.launcher_progressbar_frame, orient="horizontal", length=300, mode='determinate')
        self.launcher_progressbar.grid(padx=20,pady=10,sticky="wn")


        logging.debug("Creating update process...")
        freeze_support()
        self.update_process = Process(target=autoupdater.autoUpdater, args=(self.queue_progress, APP_NAME, APP_VERSION,))
        logging.debug("Starting update process...")
        self.update_process.start()
        self.updateTaskCheck()


    def updateTaskCheck(self):
        if PLATFORM == "Linux":
            while self.update_process.is_alive():
                if not self.queue_progress.empty():
                    self.progress = self.queue_progress.get(block=False)

                    if self.progress == "DONE":
                        logging.debug("Restarting...")
                        self.update_process.terminate()
                        python = sys.executable
                        os.execl(python, python, *sys.argv)
                    elif self.progress == "NO_UPDATE":
                        logging.debug("Starting application")
                        self.update_process.terminate()
                        self.startGUI()
                    else:
                        self.title_label["text"] = "Baixando atualização..."
                        self.launcher_progressbar["value"] = self.progress
                        root.update()

        elif PLATFORM == "Windows":
            if self.update_process.is_alive():
                if not self.queue_progress.empty():
                    self.progress = self.queue_progress.get(block=False)

                    if self.progress == "DONE":
                        logging.debug("Restarting...")
                        self.update_process.terminate()
                        python = sys.executable
                        os.execl(python, python, *sys.argv)
                    elif self.progress == "NO_UPDATE":
                        logging.debug("Starting application")
                        self.update_process.terminate()
                        self.startGUI()
                    else:
                        self.title_label["text"] = "Baixando atualização..."
                        self.launcher_progressbar["value"] = self.progress
                        root.update()
                root.after(10, self.updateTaskCheck)

    def startGUI(self):
        root.destroy()
        #root.quit()
        from interface import interface
        interface.runMain(APP_VERSION)


if __name__ == "__main__":
    root = tk.Tk()

    try:
        # for freeze version
        if PLATFORM == "Linux":
            root.iconbitmap(sys._MEIPASS + "/icons/iconapp.ico")
        elif PLATFORM == "Windows":
            root.iconbitmap(sys._MEIPASS + r"\icons\iconapp.ico")
        else:
            logging.critical("The program has not been tested on this platform. Errors may occur")
    except:
        # for script version script
        if PLATFORM == "Linux":
            root.iconbitmap(LAUNCHER_PATH+"/icons/iconapp.ico")
        elif PLATFORM == "Windows":
            root.iconbitmap(LAUNCHER_PATH+"\icons\iconapp.ico")
        else:
            logging.critical("The program has not been tested on this platform. Errors may occur")

    Launcher(root)
    root.title(TITLE)
    root.configure(bg="white")
    root.mainloop()
