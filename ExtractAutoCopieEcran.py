#!/usr/bin/env python
# coding: utf-8

# In[1]:


import tkinter as tk
from tkinter import filedialog
import pyautogui
import time
from PIL import ImageGrab
from docx import Document
from docx.shared import Inches

class ScreenshotApp:
    def __init__(self, app):
        self.app = app
        self.capture_active = False
        self.interval_id = None

        self.window_var = tk.StringVar()
        self.window_var.set("Écran entier")

        self.create_ui()

    def create_ui(self):
        self.window_label = tk.Label(self.app, text="Sélectionnez la fenêtre à capturer:")
        self.window_label.pack()

        window_options = ["Écran entier", "Sélectionner une fenêtre"]
        window_menu = tk.OptionMenu(self.app, self.window_var, *window_options)
        window_menu.pack()

        self.interval_label = tk.Label(self.app, text="Intervalle entre les captures (en secondes):")
        self.interval_label.pack()

        self.interval_entry = tk.Entry(self.app)
        self.interval_entry.pack()

        self.capture_button = tk.Button(self.app, text="Commencer la capture", command=self.capture_screenshot)
        self.capture_button.pack()

        self.stop_button = tk.Button(self.app, text="Arrêter la capture", command=self.stop_capture, state=tk.DISABLED)
        self.stop_button.pack()

        self.new_capture_button = tk.Button(self.app, text="Nouvelle capture", command=self.new_capture, state=tk.DISABLED)
        self.new_capture_button.pack()

        self.result_label = tk.Label(self.app, text="")
        self.result_label.pack()

    def capture_screenshot(self):
        self.capture_active = True
        self.capture_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.new_capture_button.config(state=tk.DISABLED)

        selected_window = self.window_var.get()
        interval = self.interval_entry.get()

        try:
            interval = int(interval)
        except ValueError:
            self.result_label.config(text="Veuillez entrer un intervalle valide.")
            self.capture_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            self.new_capture_button.config(state=tk.DISABLED)
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Documents Word", "*.docx")])

        # Créer un document Word
        doc = Document()

        def capture_periodic():
            if self.capture_active:
                self.result_label.config(text="Capture en cours...")
                if selected_window == "Écran entier":
                    screenshot = pyautogui.screenshot()
                elif selected_window == "Sélectionner une fenêtre":
                    screenshot = ImageGrab.grab(bbox=(pyautogui.getActiveWindow().left, pyautogui.getActiveWindow().top, pyautogui.getActiveWindow().right, pyautogui.getActiveWindow().bottom))

                # Enregistrez la capture d'écran dans le document Word
                screenshot_path = "screenshot.png"
                screenshot.save(screenshot_path)
                doc.add_picture(screenshot_path, width=Inches(4))

                # Réexécute la fonction après l'intervalle de temps
                self.interval_id = self.app.after(interval * 1000, capture_periodic)
            else:
                # Enregistrez le document Word
                doc.save(file_path)
                self.result_label.config(text="Capture d'écran enregistrée dans {}".format(file_path))
                self.capture_button.config(state=tk.NORMAL)
                self.stop_button.config(state=tk.DISABLED)
                self.new_capture_button.config(state=tk.NORMAL)

        # Commencer la capture d'écran périodique
        capture_periodic()

    def stop_capture(self):
        self.result_label.config(text="Arrêt de la capture en cours...")
        self.capture_active = False

    def new_capture(self):
        self.result_label.config(text="")
        self.new_capture_button.config(state=tk.DISABLED)
        self.capture_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

if __name__ == "__main__":
    app = tk.Tk()
    app.title("Capture d'écran automatique")

    screenshot_app = ScreenshotApp(app)

    app.mainloop()

