# gui.py
import ttkbootstrap as tb
from ttkbootstrap.widgets import DateEntry
from tkinter import messagebox
from datetime import datetime

def get_login_and_dates_from_gui():
    """
    Launches a modern GUI to get username, password, start date, and end date.
    Returns: username, password, start_date (YYYY-MM-DD), end_date (YYYY-MM-DD)
    """
    user_input = {}

    def submit_action():
        username = username_entry.get().strip()
        password = password_entry.get().strip()
        try:
            start_date = start_cal.get_date()
            end_date = end_cal.get_date()
        except Exception:
            messagebox.showerror("Error", "Invalid date selection")
            return

        if not username or not password:
            messagebox.showerror("Error", "Please enter both username and password")
            return

        # Convert to YYYY-MM-DD
        user_input["username"] = username
        user_input["password"] = password
        user_input["start_date"] = start_date.strftime("%Y-%m-%d")
        user_input["end_date"] = end_date.strftime("%Y-%m-%d")
        root.destroy()

    root = tb.Window(themename="superhero")
    root.title("WebOC Scraper Login & Dates")
    root.geometry("450x350")
    root.resizable(False, False)

    frame = tb.Frame(root)
    frame.pack(padx=20, pady=20, fill="both", expand=True)

    # Title
    title_label = tb.Label(frame, text="Enter Login Details & Dates", font=("Helvetica", 16, "bold"))
    title_label.pack(pady=(0, 15), anchor="w")

    # Username
    tb.Label(frame, text="Username:").pack(anchor="w")
    username_entry = tb.Entry(frame, width=30)
    username_entry.pack(pady=(0, 10), anchor="w")

    # Password
    tb.Label(frame, text="Password:").pack(anchor="w")
    password_entry = tb.Entry(frame, width=30, show="*")
    password_entry.pack(pady=(0, 15), anchor="w")

    # Date selectors with button on the right
    date_frame = tb.Frame(frame)
    date_frame.pack(fill="x", pady=(0, 10))

    # Start Date
    start_label = tb.Label(date_frame, text="Start Date:")
    start_label.grid(row=0, column=0, sticky="w", padx=(0,5))
    start_cal = DateEntry(date_frame, bootstyle="success", width=25, dateformat="%Y-%m-%d")
    start_cal.set_date(datetime.today())
    start_cal.grid(row=0, column=1, sticky="w")

    # End Date
    end_label = tb.Label(date_frame, text="End Date:")
    end_label.grid(row=1, column=0, sticky="w", padx=(0,5), pady=(5,0))
    end_cal = DateEntry(date_frame, bootstyle="success", width=25, dateformat="%Y-%m-%d")
    end_cal.set_date(datetime.today())
    end_cal.grid(row=1, column=1, sticky="w", pady=(5,0))

    # Start Scraping Button on the right of date selectors
    submit_btn = tb.Button(date_frame, text="Start Scraping", bootstyle="primary", command=submit_action)
    submit_btn.grid(row=0, column=2, rowspan=2, padx=(35,0), pady=(15,0), sticky="n")

    root.mainloop()

    return (
        user_input.get("username"),
        user_input.get("password"),
        user_input.get("start_date"),
        user_input.get("end_date"),
    )
