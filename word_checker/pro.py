import customtkinter as ctk
from tkinter import filedialog
import os
import pystray
from PIL import Image
import pandas as pd
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
from bs4 import BeautifulSoup

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class MyApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # configure window
        self.file_path1 = None
        self.file_path2 = None
        self.folder_path = None
        self.words = []
        self.title("Word Checker App")
        self.geometry("600x400")
        self.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)

        # configure grid layout (3x3)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self.grid_rowconfigure(3, weight=1)

        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=7, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="Setting Pane", anchor="center",
                                       font=ctk.CTkFont(size=20))
        self.logo_label.grid(row=2, column=0, padx=20, pady=(20, 10))

        self.file_label1 = ctk.CTkLabel(self, text="Select Websites File (.xlsx)", font=('calibri', 12),
                                        justify="center", wraplength=200)
        self.file_label1.grid(row=0, column=1, padx=10, pady=(15, 5), sticky="nsew")
        self.select_button1 = ctk.CTkButton(self, text="Select Excel File", font=('calibri', 12),
                                            command=self.select_excel_file)
        self.select_button1.grid(row=1, column=1, padx=10, pady=(5, 15), sticky="nsew")

        self.file_label2 = ctk.CTkLabel(self, text="Select Words File (.txt)", font=('calibri', 12), justify="center",
                                        wraplength=200)
        self.file_label2.grid(row=0, column=2, padx=10, pady=(15, 5), sticky="nsew")
        self.select_button2 = ctk.CTkButton(self, text="Select Text File", font=('calibri', 12),
                                            command=self.select_txt_file)
        self.select_button2.grid(row=1, column=2, padx=10, pady=(5, 15), sticky="nsew")

        self.folder_label = ctk.CTkLabel(self, text="Select Output Folder", font=('calibri', 12), justify="center",
                                         wraplength=200)
        self.folder_label.grid(row=2, column=1, columnspan=2, padx=10, pady=(15, 5), sticky="nsew")
        self.select_folder_button = ctk.CTkButton(self, text="Select Folder", font=('calibri', 12),
                                                  command=self.select_folder)
        self.select_folder_button.grid(row=3, column=1, columnspan=2, padx=10, pady=(5, 5), sticky="nsew")

        self.start_button = ctk.CTkButton(self, text="Start", font=('Tahoma', 10, 'bold'), command=self.check_words)
        self.start_button.grid(row=5, column=1, columnspan=2, padx=10, pady=50, sticky="nsew")

        self.appearance_mode_label = ctk.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = ctk.CTkOptionMenu(self.sidebar_frame, values=["Dark", "Light"],
                                                             command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))

        self.scaling_label = ctk.CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=7, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = ctk.CTkOptionMenu(self.sidebar_frame, values=["80%", "90%", "100%", "110%", "120%"],
                                                     command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 20))

        self.website_list = []

    def select_excel_file(self):
        self.file_path1 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.file_path1:
            df = pd.read_excel(self.file_path1)
            logging.info(f"Selected file: {os.path.basename(self.file_path1)}")
            self.website_list = df[df.columns[0]].tolist()[:-1]
            self.file_label1.configure(text=f"Selected file: \n {os.path.basename(self.file_path1)}")

    def select_txt_file(self):
        self.file_path2 = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if self.file_path2:
            try:
                with open(self.file_path2, 'r', encoding='utf-8') as file:
                    content = file.read()
                    self.words = set(content.split("\n"))
                logging.info(f"Selected file: {os.path.basename(self.file_path2)}")
                self.file_label2.configure(text=f"Selected file: \n {os.path.basename(self.file_path2)}")
            except UnicodeDecodeError:
                with open(self.file_path2, 'r', encoding='latin-1') as file:
                    content = file.read()
                    self.words = set(content.split("\n"))
                logging.info(f"Selected file: {os.path.basename(self.file_path2)}")
                self.file_label2.configure(text=f"Selected file: \n {os.path.basename(self.file_path2)}")

    def select_folder(self):
        self.folder_path = filedialog.askdirectory()
        if self.folder_path:
            logging.info(f"Selected folder: {os.path.basename(self.folder_path)}")
            self.folder_label.configure(text=f"Selected folder: \n {os.path.basename(self.folder_path)}")
            self.update_log_file_path()

    def update_log_file_path(self):
        log_file_path = os.path.join(self.folder_path, 'file.log')
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)
        logging.basicConfig(level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(message)s',
                            filename=log_file_path,
                            filemode='w')
        logging.info("Log file path updated")

    def check_words(self):
        if not self.website_list:
            logging.error("No website file selected")
            return

        results = {}
        with ThreadPoolExecutor(max_workers=10) as executor:
            future_to_url = {executor.submit(self.fetch_content, url): url for url in self.website_list}
            for future in as_completed(future_to_url):
                url = future_to_url[future]
                try:
                    data = future.result()
                    results[url] = data
                    logging.info(f"Fetched content from {url}")
                except Exception as exc:
                    logging.error(f"Error fetching content from {url}: {exc}")

        urls_list = []
        words_list = []
        for url, content in results.items():
            for word in self.words:
                if word in content:
                    urls_list.append(url)
                    words_list.append(word)
                    logging.info(f"Word '{word}' found in {url}")

        if urls_list:
            final_df = pd.DataFrame({'Websites_URL': urls_list, 'Word': words_list})
            final_df.to_csv(os.path.join(self.folder_path, 'output.csv'), index=False)
            logging.info("Output file created successfully")
        else:
            logging.info("No words found in any website")

    def fetch_content(self, url):
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(response.content, 'html.parser')
        for script_or_style in soup(['script', 'style']):
            script_or_style.decompose()

        visible_text = soup.get_text(separator=' ')
        return visible_text

    def change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        ctk.set_widget_scaling(new_scaling_float)

    def minimize_to_tray(self):
        self.withdraw()
        image = Image.open("icon.ico")
        menu = (pystray.MenuItem('Quit', self.quit_window),
                pystray.MenuItem('Show', self.show_window))
        icon = pystray.Icon("name", image, "File Content Copier", menu)
        icon.run()

    def quit_window(self, icon):
        icon.stop()
        self.destroy()

    def show_window(self, icon):
        icon.stop()
        self.after(0, self.deiconify)


if __name__ == "__main__":
    app = MyApp()
    app.mainloop()
