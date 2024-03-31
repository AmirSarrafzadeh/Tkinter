import tkinter as tk
from tkinter import messagebox
import pandas as pd


def is_valid_number(number):
    try:
        float(number)
        return True
    except ValueError:
        return False

class BankAccountSimulator:
    def __init__(self, root):
        row = 1
        column = 0
        self.root = root
        self.root.title("Account")
        self.root.geometry("650x650")
        self.root.configure(highlightthickness=2, highlightbackground="black", borderwidth=2)

        # Add a frame to contain all widgets
        self.frame = tk.Frame(root, borderwidth=5, relief='solid', padx=8)
        self.frame.pack(fill="both", expand=True)

        self.tab = tk.Label(self.frame, text="")
        self.tab.grid(row=row - 1, column=column, columnspan=5, ipady=2)

        # Username Label and Entry
        self.username_label = tk.Label(self.frame, text="Username", anchor="w", width=19, font=("Arial", 12))
        self.username_label.grid(row=row, column=column, sticky="e")
        self.username_entry = tk.Entry(self.frame, width=16, bg="light blue", fg="gray", font=("Arial", 15))
        self.username_entry.grid(row=row + 1, column=column, ipady=12)

        # Password Label and Entry
        self.password_label = tk.Label(self.frame, text="Password", anchor="w", width=27, font=("Arial", 12))
        self.password_label.grid(row=row, column=column + 1, sticky="e")
        self.password_entry = tk.Entry(self.frame, show="*", width=22, bg="light blue", fg="gray", font=("Arial", 15))
        self.password_entry.grid(row=row + 1, column=column + 1, ipady=12)

        self.login_button = tk.Button(self.frame, text="Login", command=self.login, width=11, font=("Tahoma", 14))
        self.login_button.grid(row=row + 1, column=column + 2, ipady=4, ipadx=5)

        # Welcome Label
        self.welcome_label = tk.Label(self.frame, text="", fg="green", font=("Arial", 30))
        self.welcome_label.grid(row=row + 4, column=column, columnspan=5, ipady=42)

        # Balance Field
        self.balance_label = tk.Label(self.frame, text="Balance", anchor="w", width=13, state="disabled",
                                      justify="left", font=("Arial", 12))
        self.balance_label.grid(row=row + 7, column=column, sticky="e", columnspan=1, ipady=5)
        self.balance_entry = tk.Entry(self.frame, state="disabled", width=14, font=("Arial", 43), justify="center")
        self.balance_entry.grid(row=row + 8, column=column, ipady=40, columnspan=3)

        self.tab = tk.Label(self.frame, text="")
        self.tab.grid(row=row + 9, column=column, columnspan=5, ipady=5)

        # deposit Field and deposit Button in one row
        self.deposit_entry = tk.Entry(self.frame, state="disabled", width=28, font=("Arial", 14), bg="light blue")
        self.deposit_entry.grid(row=row+10, column=column, ipady=15, columnspan=2)
        self.deposit_button = tk.Button(self.frame, text="deposit", command=self.deposit, state="disabled", width=12,
                                      font=("Arial", 16), pady=8)
        self.deposit_button.grid(row=row + 10, column=column + 2, columnspan=2)

        self.tab = tk.Label(self.frame, text="")
        self.tab.grid(row=row + 11, column=column, columnspan=5)

        # withdraw Field and withdraw Button in one row
        self.withdraw_entry = tk.Entry(self.frame, state="disabled", width=28, font=("Arial", 14), bg="light blue")
        self.withdraw_entry.grid(row=row + 12, column=column, columnspan=2, ipady=15)
        self.withdraw_button = tk.Button(self.frame, text="withdraw", command=self.withdraw, state="disabled", width=12,
                                         font=("Arial", 16), pady=8)
        self.withdraw_button.grid(row=row + 12, column=column + 2, columnspan=2)

        # Read Excel File
        self.data = pd.read_excel("bank_accounts.xlsx")

    def login(self):
        username = self.username_entry.get().lower()
        password = self.password_entry.get()

        if (username, password) in zip(self.data['Username'], self.data['Password']):
            index = self.data[(self.data['Username'] == username) & (self.data['Password'] == password)].index[0]
            balance = self.data.at[index, 'Balance']
            if str(balance).isnumeric():
                balance = float(balance)
            self.balance_entry.config(state="normal")
            self.balance_entry.delete(0, tk.END)
            self.username_entry.config(state="disabled")
            self.password_entry.config(state="disabled")
            self.deposit_entry.config(state="normal")
            self.deposit_button.config(state="normal")
            self.withdraw_entry.config(state="normal")
            self.withdraw_button.config(state="normal")
            self.balance_entry.insert(0, balance)
            name = self.data.at[index, 'Name']
            self.welcome_label.config(text=f"Welcome {name}")
        else:
            messagebox.showerror("Error", "Invalid username or password")

    def deposit(self):
        amount = self.deposit_entry.get()
        current_balance = self.balance_entry.get()
        if is_valid_number(amount) and is_valid_number(current_balance):
            amount = float(amount)
            current_balance = float(self.balance_entry.get())
            new_balance = current_balance + amount
            self.balance_entry.config(state="normal")
            self.balance_entry.delete(0, tk.END)
            self.balance_entry.insert(0, new_balance)
            self.deposit_entry.delete(0, tk.END)
            self.balance_entry.config(state="disabled")

    def withdraw(self):
        amount = self.withdraw_entry.get()
        current_balance = self.balance_entry.get()
        if is_valid_number(amount) and is_valid_number(current_balance):
            amount = float(amount)
            current_balance = float(self.balance_entry.get())

            if amount <= current_balance:
                new_balance = current_balance - amount
                self.balance_entry.config(state="normal")
                self.balance_entry.delete(0, tk.END)
                self.balance_entry.insert(0, new_balance)
                self.balance_entry.config(state="disabled")
            else:
                self.withdraw_entry.delete(0, tk.END)
                self.withdraw_entry.insert(0, "Low Balance")


if __name__ == "__main__":
    root = tk.Tk()
    app = BankAccountSimulator(root)
    root.mainloop()
