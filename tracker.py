import pandas as pd

file_path = "applications_tracker.xlsx"

columns = [
    "Company",
    "Country",
    "Position",
    "Date Applied",
    "Platform",
    "Status",
    "Link",
    "Notes",
    "Follow-up Date"
]

df = pd.DataFrame(columns=columns)
df.to_excel(file_path, index=False)

print("Tracker creado correctamente desde cero")