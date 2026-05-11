import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

# =========================
# CARGAR EXCEL
# =========================
file_path = "applications_tracker.xlsx"

df = pd.read_excel(file_path)

# =========================
# CALCULAR ESTADÍSTICAS
# =========================
total = len(df)

applied = len(df[df["Status"] == "Applied"])
rejected = len(df[df["Status"] == "Rejected"])
pending = len(df[df["Status"] == "Pending"])

# =========================
# ABRIR EXCEL CON OPENPYXL
# =========================
wb = load_workbook(file_path)
ws = wb.active

# =========================
# UBICACIÓN DEL RESUMEN
# =========================
summary_row = ws.max_row + 3

# Título
ws.cell(row=summary_row, column=1, value="JOB TRACKER DASHBOARD")
ws.cell(row=summary_row, column=1).font = Font(bold=True, size=14)

# Datos
ws.cell(row=summary_row + 1, column=1, value="Total Applications")
ws.cell(row=summary_row + 1, column=2, value=total)

ws.cell(row=summary_row + 2, column=1, value="Applied")
ws.cell(row=summary_row + 2, column=2, value=applied)

ws.cell(row=summary_row + 3, column=1, value="Rejected")
ws.cell(row=summary_row + 3, column=2, value=rejected)

ws.cell(row=summary_row + 4, column=1, value="Pending")
ws.cell(row=summary_row + 4, column=2, value=pending)

# =========================
# GUARDAR
# =========================
wb.save(file_path)

print("✔ Dashboard agregado al Excel")