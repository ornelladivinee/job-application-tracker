import pandas as pd

df = pd.read_excel("applications_tracker.xlsx")

total = len(df)
applied = len(df[df["Status"] == "Applied"])
rejected = len(df[df["Status"] == "Rejected"])
pending = len(df[df["Status"] == "Pending"])

print("\nJOB TRACKER DASHBOARD")
print(f"Total: {total}")
print(f"Applied: {applied}")
print(f"Rejected: {rejected}")
print(f"Pending: {pending}")