import tkinter as tk
from tkinter import messagebox, simpledialog
import pandas as pd
from datetime import datetime, timedelta
import os
import win32com.client

# Global variables
root = None
work_mode_var = None
break_times = []
login_time = None
logout_time = None
total_working_hours = None
username = None
user_id = None
work_mode = None
log_file = "tracker_log.xlsx"


def initialize_tracker_log():
    if not os.path.exists(log_file):
        columns = [
            "Date", "Username", "User ID", "Work Mode", "Login Time",
            "Each Break Time", "Logout Time", "Total Working Hours"
        ]
        df = pd.DataFrame(columns=columns)
        df.to_excel(log_file, index=False)


def register_user():
    global username, user_id, work_mode
    if not os.path.exists("user_info.txt"):
        username = simpledialog.askstring("Name", "Please enter your name:")
        user_id = simpledialog.askstring("User ID", "Please enter your User ID:")
        work_mode = work_mode_var.get()
        with open("user_info.txt", "w") as file:
            file.write(f"{username}\n{user_id}\n{work_mode}")
    else:
        with open("user_info.txt", "r") as file:
            username, user_id, work_mode = file.read().splitlines()


def create_widgets():
    tk.Label(root, text=f"Welcome {username}", bg="blue", fg="white", font=("Helvetica", 16)).pack()
    tk.Label(root, text=f"{user_id}\n{datetime.now().strftime('%m/%d/%Y')}", bg="blue", fg="white", font=("Helvetica", 14)).pack()
    
    wfo_wfh_frame = tk.Frame(root, bg="blue")
    wfo_wfh_frame.pack()
    tk.Radiobutton(wfo_wfh_frame, text="WFO", variable=work_mode_var, value="WFO", bg="blue", fg="white", font=("Helvetica", 14)).pack(side="left", padx=10)
    tk.Radiobutton(wfo_wfh_frame, text="WFH", variable=work_mode_var, value="WFH", bg="blue", fg="white", font=("Helvetica", 14)).pack(side="left", padx=10)

    button_frame = tk.Frame(root, bg="blue")
    button_frame.pack(pady=20)
    create_button(button_frame, "Login", login, "green")
    create_button(button_frame, "Break Start", break_start, "orange")
    create_button(button_frame, "Break End", break_end, "orange")
    create_button(button_frame, "Logout", logout, "red")

    tk.Button(root, text="Send Log", command=send_email, font=("Helvetica", 14), bg="orange", width=20).pack(pady=10)
    tk.Button(root, text="Exit", command=root.quit, font=("Helvetica", 14), width=10).pack(pady=10)


def create_button(parent, text, command, color):
    tk.Button(parent, text=text, command=command, font=("Helvetica", 12), bg=color, fg="white", width=15).pack(side="left", padx=10)


def log_to_excel():
    data = {
        "Date": [datetime.now().strftime('%Y-%m-%d')],
        "Username": [username],
        "User ID": [user_id],
        "Work Mode": [work_mode],
        "Login Time": [login_time],
        "Each Break Time": [", ".join(break_times)],
        "Logout Time": [logout_time],
        "Total Working Hours": [total_working_hours]
    }
    new_df = pd.DataFrame(data)
    existing_df = pd.read_excel(log_file)
    updated_df = pd.concat([existing_df, new_df], ignore_index=True)
    updated_df.to_excel(log_file, index=False)


def calculate_total_hours():
    global total_working_hours
    login_dt = datetime.strptime(login_time, '%Y-%m-%d %H:%M:%S')
    logout_dt = datetime.strptime(logout_time, '%Y-%m-%d %H:%M:%S')
    total_hours = logout_dt - login_dt
    
    break_duration = timedelta()
    for break_time in break_times:
        start, end = map(lambda x: datetime.strptime(x, '%Y-%m-%d %H:%M:%S'), break_time.split(" - "))
        break_duration += (end - start)
    
    total_working_hours = str(total_hours - break_duration)


def login():
    global login_time
    login_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    messagebox.showinfo("Info", "Logged In")


def break_start():
    break_times.append(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - ")
    messagebox.showinfo("Info", "Break Started")


def break_end():
    if break_times and break_times[-1].endswith(" - "):
        break_times[-1] += datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        messagebox.showinfo("Info", "Break Ended")
    else:
        messagebox.showwarning("Warning", "No break in progress!")


def logout():
    global logout_time
    logout_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    calculate_total_hours()
    log_to_excel()
    messagebox.showinfo("Info", "Logged Out")


def send_email():
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
        mail.Attachments.Add(os.path.abspath(log_file))
        mail.Send()
        messagebox.showinfo("Info", "Email sent successfully via Outlook!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to send email: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Automation LILO- Tracker")
    root.geometry("800x400")
    root.configure(bg="blue")
    work_mode_var = tk.StringVar(value="WFO")
    initialize_tracker_log()
    register_user()
    create_widgets()
    root.mainloop()