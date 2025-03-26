import tkinter as tk
from tkinter import messagebox, simpledialog
import pandas as pd
from datetime import datetime, timedelta
import os
import win32com.client

class TrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automation LILO- Tracker")
        self.root.geometry("800x400")
        self.root.configure(bg="blue")

        self.work_mode_var = tk.StringVar(value="WFO")
        self.break_times = []
        self.login_time = None
        self.logout_time = None
        self.total_working_hours = None
        self.register_user()

        self.log_file = "tracker_log.xlsx"
        self.initialize_tracker_log()

        self.create_widgets()

    def initialize_tracker_log(self):
        if not os.path.exists(self.log_file):
            columns = [
                "Date", "Username", "User ID", "Work Mode", "Login Time",
                "Each Break Time", "Logout Time", "Total Working Hours"
            ]
            df = pd.DataFrame(columns=columns)
            df.to_excel(self.log_file, index=False)

    def register_user(self):
        if not os.path.exists("user_info.txt"):
            self.username = simpledialog.askstring("Name", "Please enter your name:")
            self.user_id = simpledialog.askstring("User ID", "Please enter your User ID:")
            self.work_mode = self.work_mode_var.get()
            with open("user_info.txt", "w") as file:
                file.write(f"{self.username}\n{self.user_id}\n{self.work_mode}")
        else:
            with open("user_info.txt", "r") as file:
                self.username, self.user_id, self.work_mode = file.read().splitlines()

    def create_widgets(self):
        tk.Label(self.root, text=f"Welcome {self.username}", bg="blue", fg="white", font=("Helvetica", 16)).pack()
        tk.Label(self.root, text=f"{self.user_id}\n{datetime.now().strftime('%m/%d/%Y')}", bg="blue", fg="white", font=("Helvetica", 14)).pack()
        
        wfo_wfh_frame = tk.Frame(self.root, bg="blue")
        wfo_wfh_frame.pack()
        tk.Radiobutton(wfo_wfh_frame, text="WFO", variable=self.work_mode_var, value="WFO", bg="blue", fg="white", font=("Helvetica", 14)).pack(side="left", padx=10)
        tk.Radiobutton(wfo_wfh_frame, text="WFH", variable=self.work_mode_var, value="WFH", bg="blue", fg="white", font=("Helvetica", 14)).pack(side="left", padx=10)

        button_frame = tk.Frame(self.root, bg="blue")
        button_frame.pack(pady=20)
        self.create_button(button_frame, "Login", self.login, "green")
        self.create_button(button_frame, "Break Start", self.break_start, "orange")
        self.create_button(button_frame, "Break End", self.break_end, "orange")
        self.create_button(button_frame, "Logout", self.logout, "red")

        tk.Button(self.root, text="Send Log", command=self.send_email, font=("Helvetica", 14), bg="orange", width=20).pack(pady=10)
        tk.Button(self.root, text="Exit", command=self.root.quit, font=("Helvetica", 14), width=10).pack(pady=10)

    def create_button(self, parent, text, command, color):
        tk.Button(parent, text=text, command=command, font=("Helvetica", 12), bg=color, fg="white", width=15).pack(side="left", padx=10)

    def log_to_excel(self):
        data = {
            "Date": [datetime.now().strftime('%Y-%m-%d')],
            "Username": [self.username],
            "User ID": [self.user_id],
            "Work Mode": [self.work_mode],
            "Login Time": [self.login_time],
            "Each Break Time": [", ".join(self.break_times)],
            "Logout Time": [self.logout_time],
            "Total Working Hours": [self.total_working_hours]
        }
        new_df = pd.DataFrame(data)
        existing_df = pd.read_excel(self.log_file)
        updated_df = pd.concat([existing_df, new_df], ignore_index=True)
        updated_df.to_excel(self.log_file, index=False)

    def calculate_total_hours(self):
        login_dt = datetime.strptime(self.login_time, '%Y-%m-%d %H:%M:%S')
        logout_dt = datetime.strptime(self.logout_time, '%Y-%m-%d %H:%M:%S')
        total_hours = logout_dt - login_dt
        
        break_duration = timedelta()
        for break_time in self.break_times:
            start, end = map(lambda x: datetime.strptime(x, '%Y-%m-%d %H:%M:%S'), break_time.split(" - "))
            break_duration += (end - start)
        
        self.total_working_hours = str(total_hours - break_duration)

    def login(self):
        self.login_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        messagebox.showinfo("Info", "Logged In")

    def break_start(self):
        self.break_times.append(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - ")
        messagebox.showinfo("Info", "Break Started")

    def break_end(self):
        if self.break_times and self.break_times[-1].endswith(" - "):
            self.break_times[-1] += datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            messagebox.showinfo("Info", "Break Ended")
        else:
            messagebox.showwarning("Warning", "No break in progress!")

    def logout(self):
        self.logout_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.calculate_total_hours()
        self.log_to_excel()
        messagebox.showinfo("Info", "Logged Out")

    def send_email(self):
        try:
            email = simpledialog.askstring("Email", "Enter recipient's email:")
            if not email:
                messagebox.showerror("Error", "Email address is required!")
                return

            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = email
            mail.Subject = "Daily Activity Log"
            mail.Body = "Attached is the daily tracker log."
            mail.Attachments.Add(os.path.abspath(self.log_file))
            mail.Send()
            messagebox.showinfo("Info", "Email sent successfully via Outlook!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send email: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TrackerApp(root)
    root.mainloop()
