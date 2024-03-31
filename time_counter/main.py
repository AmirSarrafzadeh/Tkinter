import tkinter as tk
from tkinter import messagebox


def is_valid_number(number):
    try:
        float(number)
        return True
    except ValueError:
        return False

class TaskTimer:
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry("300x300")
        self.root.resizable(False, False)
        self.root.title("Time Counter")

        # Task Entry
        tk.Label(self.root, text="Task:", font=("Arial", 12), width=10).grid(row=0, column=0, padx=5, pady=5)
        self.task_entry = tk.Entry(self.root, font=("Arial", 12))
        self.task_entry.grid(row=0, column=1)

        # Time Entry (Minutes, Seconds)
        tk.Label(self.root, text="Minutes:", font=("Arial", 12), width=10).grid(row=1, column=0, padx=5, pady=5)
        self.minutes_entry = tk.Entry(self.root, font=("Arial", 12))
        self.minutes_entry.grid(row=1, column=1)


        tk.Label(self.root, text="Seconds:", font=("Arial", 12), width=10).grid(row=2, column=0, padx=5, pady=5)
        self.seconds_entry = tk.Entry(self.root, font=("Arial", 12))
        self.seconds_entry.grid(row=2, column=1)

        # Timer Display
        self.timer_label = tk.Label(self.root, text="00:00", font=('Helvetica', 24))
        self.timer_label.grid(row=3, column=0, columnspan=2)
        timer_font = ('Helvetica', 30, 'bold')
        self.timer_label.config(font=timer_font)
        self.root.configure(bg='lightblue')
        self.timer_label.config(bg='lightblue', fg='white')

        control_frame = tk.Frame(self.root, bg='lightblue', pady=10, padx=10, relief='raised', borderwidth=2)
        control_frame.grid(row=4, column=0, columnspan=2)

        # Buttons
        tk.Button(control_frame, text="Start", command=self.start_timer, font=("Arial", 12), width=10, padx=5, pady=5).pack()
        tk.Button(control_frame, text="Pause", command=self.pause_timer, font=("Arial", 12), width=10, padx=5, pady=5).pack()


        # Initialize variables
        self.remaining = None
        self.paused = False

        self.root.mainloop()

    def start_timer(self):
        minutes = self.minutes_entry.get()
        seconds = self.seconds_entry.get()

        if not minutes and not seconds:
            messagebox.showerror("Error", "Please enter a time")
            return

        if minutes == '':
            minutes = 0
        if seconds == '':
            seconds = 0

        check_minutes = is_valid_number(minutes)
        check_seconds = is_valid_number(seconds)
        if check_minutes and check_seconds:
            minutes = int(minutes)
            seconds = int(seconds)
        else:
            messagebox.showerror("Error", "Please enter a valid time")
            return



        try:
            self.remaining = minutes * 60 + seconds
        except ValueError:
            messagebox.showerror("Error", "Invalid time input")
            return

        if not self.paused:
            self.countdown()

    def pause_timer(self):
        self.paused = True

    def countdown(self):
        if self.remaining > 0 and not self.paused:
            self.timer_label.config(bg='lightblue')
            if self.remaining <= 10:
                self.timer_label.config(bg='red')
            mins, secs = divmod(self.remaining, 60)
            self.timer_label.config(text=f"{mins:02d}:{secs:02d}")
            self.remaining -= 1
            self.root.after(1000, self.countdown)
        else:
            self.paused = False
            if self.remaining == 0:
                messagebox.showinfo("Timer", "Time's up!")

if __name__ == "__main__":
    timer = TaskTimer()
