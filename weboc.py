# weboc.py
from gui import get_login_and_dates_from_gui
from automation_function import start_scraping

if __name__ == "__main__":
    # Launch GUI to get login & dates
    username, password, start_date, end_date = get_login_and_dates_from_gui()

    if not username or not password or not start_date or not end_date:
        print("No input provided. Exiting...")
        exit()

    # Run scraper with GUI input
    start_scraping(username, password, start_date, end_date, progress_callback=print)
