import numpy as np
import customtkinter as ctk
from tkinter import filedialog
import os
import pystray
from PIL import Image
import pandas as pd
import logging
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.table import Table

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class MyApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.file_path1 = None
        self.folder_path = None
        self.title("Word Checker App")
        self.geometry("700x600")
        self.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self.grid_rowconfigure(3, weight=1)

        self.sidebar_frame = ctk.CTkFrame(self, width=50, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=15, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="Setting Pane", anchor="center",
                                       font=ctk.CTkFont(size=15))
        self.logo_label.grid(row=1, column=0, padx=20, pady=(5, 10))

        self.file_label1 = ctk.CTkLabel(self, text="", font=('calibri', 12),
                                        justify="center", wraplength=200)
        self.file_label1.grid(row=0, column=1, columnspan=2, padx=10, pady=5, sticky="nsew")
        self.select_button1 = ctk.CTkButton(self, text="Select Excel File (.xlsx)", font=('calibri', 12),
                                            command=self.select_excel_file)
        self.select_button1.grid(row=1, column=1, columnspan=2, padx=10, pady=5, sticky="nsew", ipadx=5)

        self.folder_label = ctk.CTkLabel(self, text="", font=('calibri', 12), justify="center",
                                         wraplength=200)
        self.folder_label.grid(row=2, column=1, columnspan=2, padx=10, pady=(5, 5), sticky="nsew")
        self.select_folder_button = ctk.CTkButton(self, text="Select Output Folder", font=('calibri', 12),
                                                  command=self.select_folder)
        self.select_folder_button.grid(row=3, column=1, columnspan=2, padx=10, pady=(5, 5), sticky="nsew", ipadx=3)
        self.file_label1 = ctk.CTkLabel(self, text="", font=('calibri', 12),
                                        justify="center", wraplength=200)
        self.file_label1.grid(row=4, column=1, columnspan=2, padx=10, pady=5, sticky="nsew")

        self.integer_label = ctk.CTkLabel(self, text="Enter last Index", font=('calibri', 12),
                                          justify="center",
                                          wraplength=200)
        self.integer_label.grid(row=6, column=1, padx=10, pady=(5, 5), sticky="w")
        self.last_index = ctk.CTkEntry(self, font=('calibri', 12), width=100)
        self.last_index.grid(row=6, column=2, padx=10, pady=(5, 5), sticky="w", ipadx=20)

        self.integer_label = ctk.CTkLabel(self, text="Enter new Index", font=('calibri', 12),
                                          justify="center",
                                          wraplength=200)
        self.integer_label.grid(row=7, column=1, padx=10, pady=(5, 5), sticky="w")
        self.new_index = ctk.CTkEntry(self, font=('calibri', 12), width=100)
        self.new_index.grid(row=7, column=2, padx=10, pady=(5, 5), sticky="w", ipadx=20)

        self.integer_label = ctk.CTkLabel(self, text="Enter Min Limit", font=('calibri', 12),
                                          justify="center",
                                          wraplength=200)
        self.integer_label.grid(row=8, column=1, padx=10, pady=(5, 5), sticky="w")
        self.min_limit = ctk.CTkEntry(self, font=('calibri', 12), width=100)
        self.min_limit.grid(row=8, column=2, padx=10, pady=(5, 5), sticky="w", ipadx=20)

        self.integer_label = ctk.CTkLabel(self, text="Enter Max Limit", font=('calibri', 12),
                                          justify="center",
                                          wraplength=200)
        self.integer_label.grid(row=9, column=1, padx=10, pady=(5, 5), sticky="w")
        self.max_limit = ctk.CTkEntry(self, font=('calibri', 12), width=100)
        self.max_limit.grid(row=9, column=2, padx=10, pady=(5, 5), sticky="w", ipadx=20)

        self.name_label = ctk.CTkLabel(self, text="Enter Folder Name", font=('calibri', 12),
                                       justify="center", wraplength=200)
        self.name_label.grid(row=10, column=1, padx=10, pady=(5, 5), sticky="w")
        self.folder_name = ctk.CTkEntry(self, font=('calibri', 12), width=100)
        self.folder_name.grid(row=10, column=2, padx=10, pady=(5, 5), sticky="w", ipadx=20)

        self.start_button = ctk.CTkButton(self, text="Start", font=('Tahoma', 10, 'bold'), command=self.start)
        self.start_button.grid(row=11, column=1, columnspan=2, padx=10, pady=50, sticky="nsew")

        self.appearance_mode_label = ctk.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=9, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = ctk.CTkOptionMenu(self.sidebar_frame, values=["dark", "light"],
                                                             command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=10, column=0, padx=20, pady=(10, 10))

        self.scaling_label = ctk.CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=11, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = ctk.CTkOptionMenu(self.sidebar_frame, values=["80%", "90%", "100%", "110%", "120%"],
                                                     command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=12, column=0, padx=20, pady=(10, 20))

        self.icon = None

    def select_excel_file(self):
        self.file_path1 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

    def select_folder(self):
        self.folder_path = filedialog.askdirectory()

    def change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        ctk.set_widget_scaling(new_scaling_float)

    def minimize_to_tray(self):
        self.withdraw()
        image = Image.open("icon.ico")
        menu = pystray.Menu(
            pystray.MenuItem('Quit', self.quit_window),
            pystray.MenuItem('Show', self.show_window)
        )
        self.icon = pystray.Icon("name", image, "File Content Copier", menu)
        self.icon.run()

    def quit_window(self, icon, item):
        icon.stop()
        self.destroy()

    def show_window(self, icon, item):
        icon.stop()
        self.after(0, self.deiconify)

    # Function to generate sequences of indexes
    def print_sequences(self, n):
        temp = []
        for start in range(1, n):
            for end in range(start + 2, n + 1):
                temp.append(list(range(start, end + 1)))
        return temp

    # Function to generate sequences of indexes
    def new_sequences(self, new_index):
        temp = []
        for start in range(1, new_index - 1):
            for end in range(start, new_index):
                temp.append(list(range(start, new_index + 1)))
                break
        return temp

    # Function to check if all regression line values lie within the range [0.3, 0.7]
    def check_regression_in_range(self, df, low=0.4, high=0.6):
        x = df['index']
        b_linear = np.poly1d(np.polyfit(x, df['b'], 1))(x)
        b_poly = np.poly1d(np.polyfit(x, df['b'], 2))(x)
        s_linear = np.poly1d(np.polyfit(x, df['s'], 1))(x)
        s_poly = np.poly1d(np.polyfit(x, df['s'], 2))(x)

        return all((low <= b_linear) & (b_linear <= high)) and \
            all((low <= b_poly) & (b_poly <= high)) and \
            all((low <= s_linear) & (s_linear <= high)) and \
            all((low <= s_poly) & (s_poly <= high))

    # Function to plot the dataframe
    def plot_df(self, df, item, folder_name="plots", min_limit=0.4, max_limit=0.6):
        if len(df) < 3:
            return None

        if not self.check_regression_in_range(df, min_limit, max_limit):
            return None

        plt.figure(figsize=(12, 6))
        ax = plt.gca()

        sns.regplot(x='index', y='b', data=df, scatter=False, line_kws={'color': '#FFD700', 'linestyle': (0, (1, 1))},
                    ci=None, label='Linear Regression b')
        sns.regplot(x='index', y='b', data=df, scatter=False, line_kws={'color': '#FF1493', 'linestyle': (0, (1, 1))},
                    order=2, ci=None, label='Polynomial Regression b')
        sns.regplot(x='index', y='s', data=df, scatter=False, line_kws={'color': 'black', 'linestyle': (0, (1, 1))},
                    ci=None, label='Linear Regression s')
        sns.regplot(x='index', y='s', data=df, scatter=False, line_kws={'color': '#00FFFF', 'linestyle': (0, (1, 1))},
                    order=2, ci=None, label='Polynomial Regression s')

        plt.legend()
        plt.title('Regression Plot')
        plt.xlabel('Index')
        plt.ylabel('Value')
        plt.grid(True)

        # Add table to the plot
        Points = len(df)
        Date = df.iloc[-1]['Date']
        Time = df.iloc[-1]['Time']
        Density_b = round(df.iloc[-1]['b'], 6)
        Density_s = round(df.iloc[-1]['s'], 6)
        Measure = round(df.iloc[-1]['H'] - df.iloc[-1]['I'], 3)
        Sum_Density_b = round(sum(df['b']), 6)
        Sum_Density_s = round(sum(df['s']), 6)
        Density_Diff = round(round(sum(df['b']), 6) - round(sum(df['s']), 6), 6)
        Color = round(df.iloc[-1]['C'] - df.iloc[-1]['O'], 3)
        Mean_Color = round((sum(df['C']) - sum(df['O'])) / len(df), 3)

        data_table = [['Points', Points], ['Date', Date], ['Time', Time], ['Density_b', Density_b],
                      ['Density_s', Density_s],
                      ['Measure', Measure], ['Sum_Density_b', Sum_Density_b], ['Sum_Density_s', Sum_Density_s],
                      ['Color', Color], ['Mean_Color', Mean_Color], ['Density_Diff', Density_Diff]]

        table = Table(ax, bbox=[-0.3, 0.22, 0.23, 0.55])
        for i, (key, value) in enumerate(data_table):
            table.add_cell(i, 0, width=1.0, height=0.2, text=key, loc='left', edgecolor='black')
            table.add_cell(i, 1, width=1.0, height=0.2, text=value, loc='right', edgecolor='black')
        table.set_fontsize(20)
        table.scale(1, 1.5)
        ax.add_table(table)

        save_path = os.path.join(self.folder_path, folder_name) if folder_name else self.folder_path
        if not os.path.exists(save_path):
            os.makedirs(save_path)
        plt.tight_layout()
        plt.savefig(
            f"{save_path}/[{item[0]}_{item[-1]}]_{df.iloc[-1]['Date']}_{str(df.iloc[-1]['Time']).replace(':', ';')[0:-3]}.png")
        plt.close()

    def start(self):
        if not self.file_path1 or not self.folder_path:
            return

        if self.last_index.get() != "":
            last_index = int(self.last_index.get())
        elif self.new_index.get() != "":
            new_index = int(self.new_index.get())
        if self.min_limit.get() != "":
            min_limit = float(self.min_limit.get())
        if self.max_limit.get() != "":
            max_limit = float(self.max_limit.get())
        if self.folder_name.get() != "":
            folder_name = self.folder_name.get()

        if self.last_index.get() == "":
            last_index = "false"

        if self.min_limit.get() == "":
            min_limit = -99
        if self.max_limit.get() == "":
            max_limit = 99
        if self.folder_name.get() == "":
            folder_name = "plots"

        df = pd.read_excel(self.file_path1)
        df['index'] = df.index
        df['index'] = range(1, len(df) + 1)
        df = df[['index', 'Date', 'Time', 'O', 'H', 'I', 'C']]

        df['b'] = (df['C'] - df['I']) / (df['H'] - df['I'])
        df['s'] = (df['H'] - df['C']) / (df['H'] - df['I'])

        if last_index != "false":
            sequences = self.print_sequences(int(last_index))
            for item in sequences:
                temp_df = df.loc[df['index'].isin(item)]
                self.plot_df(temp_df, item, folder_name, min_limit, max_limit)
        elif new_index:
            sequences = self.new_sequences(int(new_index))
            for item in sequences:
                temp_df = df.loc[df['index'].isin(item)]
                self.plot_df(temp_df, item, folder_name, min_limit, max_limit)


if __name__ == "__main__":
    app = MyApp()
    app.mainloop()
