import tkinter as tk
from tkinter import ttk
import requests

# Function to translate text using Google Translate API
def translate_text():
    source_text = source_text_box.get("1.0", tk.END).strip()
    if source_text:
        url = "https://translate.googleapis.com/translate_a/single"
        headers = {
            "Content-Type": "application/json"
        }
        source_lang = source_lang_var.get()
        target_lang = target_lang_var.get()
        params = {
            "q": source_text,
            "client": "gtx",
            "sl": source_lang,
            "tl": target_lang,
            "dt": 't'
        }
        response = requests.get(url, params=params)
        translated_text = response.json()[0][0][0]
        target_text_box.delete("1.0", tk.END)
        target_text_box.insert(tk.END, translated_text)

# Function to swap languages
def swap_languages():
    current_source = source_lang_var.get()
    current_target = target_lang_var.get()
    source_lang_var.set(current_target)
    target_lang_var.set(current_source)
    source_label.config(text=f"{lang_dict[current_target]} Text")
    target_label.config(text=f"{lang_dict[current_source]} Text")

# Create the main window
root = tk.Tk()
root.title("Language Translator")

# Define language options
lang_dict = {"en": "English", "fa": "Persian"}
lang_options = list(lang_dict.keys())

# Set up language variables
source_lang_var = tk.StringVar(value="en")
target_lang_var = tk.StringVar(value="fa")

# Create and place the language selection dropdowns
language_frame = ttk.Frame(root, padding="10")
language_frame.grid(row=0, column=0, columnspan=2, sticky="ew")

source_lang_menu = ttk.Combobox(language_frame, textvariable=source_lang_var, values=lang_options, state="readonly", width=10)
source_lang_menu.grid(row=0, column=0, padx=5)
swap_button = ttk.Button(language_frame, text="Swap Languages", command=swap_languages)
swap_button.grid(row=0, column=1, padx=5)
target_lang_menu = ttk.Combobox(language_frame, textvariable=target_lang_var, values=lang_options, state="readonly", width=10)
target_lang_menu.grid(row=0, column=2, padx=5)

# Set up the frames
left_frame = ttk.Frame(root, padding="10")
right_frame = ttk.Frame(root, padding="10")

left_frame.grid(row=1, column=0, sticky="nsew")
right_frame.grid(row=1, column=1, sticky="nsew")

# Configure grid layout
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=1)
root.rowconfigure(1, weight=1)

# Create and place the text boxes and labels
source_label = ttk.Label(left_frame, text="Text")
source_label.pack(anchor="w")
source_text_box = tk.Text(left_frame, wrap="word", width=40, height=15)
source_text_box.pack(expand=True, fill="both")

target_label = ttk.Label(right_frame, text="Text")
target_label.pack(anchor="w")
target_text_box = tk.Text(right_frame, wrap="word", width=40, height=15)
target_text_box.pack(expand=True, fill="both")

# Create and place the translate button
translate_button = ttk.Button(root, text="Translate", command=translate_text)
translate_button.grid(row=2, column=0, columnspan=2, pady=10)

# Run the application
root.mainloop()
