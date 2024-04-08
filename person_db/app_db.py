import tkinter as tk
from tkinter import ttk
import sqlite3

def save_to_database():
    name = name_entry.get()
    surname = surname_entry.get()
    age = age_entry.get()

    conn = sqlite3.connect('person_database.db')
    c = conn.cursor()

    c.execute('''CREATE TABLE IF NOT EXISTS persons
                 (name TEXT, surname TEXT, age INTEGER)''')

    c.execute("INSERT INTO persons VALUES (?, ?, ?)", (name, surname, age))
    conn.commit()

    conn.close()

    name_entry.delete(0, tk.END)
    surname_entry.delete(0, tk.END)
    age_entry.delete(0, tk.END)

root = tk.Tk()
root.title("Person Information")

# Set window size and position
window_width = 400
window_height = 400
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# Styling
style = ttk.Style()
style.configure('TLabel', font=('Helvetica', 12))
style.configure('TEntry', font=('Helvetica', 12))
style.configure('TButton', font=('Helvetica', 12))

content_frame = ttk.Frame(root, padding="10")
content_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

tk.Label(content_frame, text="Name:", font=('Helvetica', 10)).grid(row=0, column=0, sticky=tk.W, padx=5, pady=15)
name_entry = ttk.Entry(content_frame)
name_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=50, pady=15)

tk.Label(content_frame, text="Surname:", font=('Helvetica', 10)).grid(row=1, column=0, sticky=tk.W, padx=5, pady=15)
surname_entry = ttk.Entry(content_frame)
surname_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=50, pady=15)

tk.Label(content_frame, text="Age:", font=('Helvetica', 10)).grid(row=2, column=0, sticky=tk.W, padx=5, pady=15)
age_entry = ttk.Entry(content_frame)
age_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=50, pady=15)

# Define a new style for the Save button with Tahoma font and font size 20
style.configure('TSave.TButton', font=('Tahoma', 16))
save_button = ttk.Button(content_frame, text="Save", command=save_to_database, style='TSave.TButton')
save_button.grid(row=3, columnspan=2, pady=30)

# Add padding to the content frame
content_frame.grid(padx=10, pady=10)

root.mainloop()
