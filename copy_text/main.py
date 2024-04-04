import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os
import pystray
from PIL import Image


class MyApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("File Content Copier")
        self.geometry("480x400")
        self.configure(background='#222222')
        self.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)

        self.file_label1 = tk.Label(self, text="Select Input", font=('calibri', 12), justify="center", wraplength=200, bg='#222222', fg='white')
        self.file_label1.grid(row=0, column=0, padx=25, pady=25, sticky="n")

        self.select_button1 = tk.Button(self, text="Select Source File", font=('calibri', 12), command=self.select_file)
        self.select_button1.grid(row=1, column=0, padx=25, pady=25, sticky="n")

        self.file_label2 = tk.Label(self, text="Select Output", font=('calibri', 12), justify="center", wraplength=200, bg='#222222', fg='white')
        self.file_label2.grid(row=0, column=2, padx=25, pady=25, sticky="n")

        self.select_button2 = tk.Button(self, text="Select Output File", font=('calibri', 12), command=self.select_save_file)
        self.select_button2.grid(row=1, column=2, padx=25, pady=25, sticky="n")

        self.start_button = tk.Button(self, text="Start", font=('Tahoma', 15, 'bold'), command=self.copy_content)
        self.start_button.grid(row=2, column=1, padx=30, pady=30)

        self.result_label = tk.Label(self, text="Output", font=('calibri', 12, 'italic'))
        self.result_label.grid(row=3, column=1, padx=10, pady=10)

    def minimize_to_tray(self):
        self.withdraw()
        image = Image.open("icon.ico")
        menu = (pystray.MenuItem('Quit',  self.quit_window),
                pystray.MenuItem('Show', self.show_window))
        icon = pystray.Icon("name", image, "File Content Copier", menu)
        icon.run()

    def quit_window(self, icon):
        icon.stop()
        self.destroy()

    def show_window(self, icon):
        icon.stop()
        self.after(0, self.deiconify)

    def select_file(self):
        self.file_path1 = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if self.file_path1:
            self.file_label1.config(text=f"Selected file: \n {os.path.basename(self.file_path1)}", justify="center")

    def select_save_file(self):
        self.file_path2 = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if self.file_path2:
            self.file_label2.config(text=f"Save file as: \n {os.path.basename(self.file_path2)}", justify="center")

    def copy_content(self):
        if hasattr(self, 'file_path1') and hasattr(self, 'file_path2'):
            with open(self.file_path1, 'r') as file1:
                content = file1.read()
                with open(self.file_path2, 'w') as file2:
                    file2.write(content)
                    self.result_label.config(text="Done!")
        else:
            self.result_label.config(text="Please select both input and output files.")


if __name__ == "__main__":
    app = MyApp()
    app.mainloop()
