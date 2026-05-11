import pandas as pd
from models import Application

FILE = "applications_tracker.xlsx"


def load_data():
    try:
        return pd.read_excel(FILE)
    except FileNotFoundError:
        return pd.DataFrame(columns=[
            "Company", "Country", "Position",
            "Date Applied", "Platform", "Status", "Notes"
        ])


def save_data(df):
    df.to_excel(FILE, index=False)


def add_application(app: Application):
    df = load_data()

    new_row = pd.DataFrame([app.__dict__])
    df = pd.concat([df, new_row], ignore_index=True)

    save_data(df)
    return df