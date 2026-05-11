import pandas as pd
from datetime import datetime

# Cargar Excel existente
df = pd.read_excel("applications_tracker.xlsx")

# Lista de postulaciones
applications = [
    {
        "Company": "NTT Data",
        "Country": "Spain - A Coruña",
        "Position": "Prácticas Ingeniería Informática - Telecomunicaciones",
        "Date Applied": "2026-05-06",
        "Platform": "iAgora",
        "Status": "Applied",
        "Link": "",
        "Notes": "IT internship"
    },
    {
        "Company": "Airbus",
        "Country": "Spain - Getafe",
        "Position": "#Discover II 2026-2027 Internship FDT Stress Team",
        "Date Applied": "2026-05-06",
        "Platform": "iAgora",
        "Status": "Applied",
        "Link": "",
        "Notes": "Engineering / Stress team"
    },
    {
        "Company": "Siemens",
        "Country": "Portugal - Porto",
        "Position": "Azure DevOps Trainee (m/f/d)",
        "Date Applied": "2026-05-06",
        "Platform": "iAgora",
        "Status": "Rejected",
        "Link": "",
        "Notes": "Cloud / DevOps / Azure"
    },
    {
        "Company": "Airbus",
        "Country": "Portugal - Lisbon",
        "Position": "IT Trainee",
        "Date Applied": "2026-05-06",
        "Platform": "iAgora",
        "Status": "Rejected",
        "Link": "",
        "Notes": "IT Systems"
    },
    {
        "Company": "Capgemini",
        "Country": "Spain - Valencia",
        "Position": "Prácticas en Ciberseguridad",
        "Date Applied": "",
        "Platform": "iAgora",
        "Status": "Applied",
        "Link": "",
        "Notes": "Cybersecurity internship"
    },
    {
        "Company": "Bosch",
        "Country": "Portugal",
        "Position": "Data Engineering Internship (Lakehouse / Databricks)",
        "Date Applied": "2026-05-07",
        "Platform": "iAgora",
        "Status": "Applied",
        "Link": "",
        "Notes": "Data engineering, cloud, Hadoop to Databricks migration"
    },
    {
        "Company": "Zurich Insurance",
        "Country": "Spain",
        "Position": "Trainee Cloud Application Developer - 133025",
        "Date Applied": "2026-05-07",
        "Platform": "Zurich Careers / iAgora",
        "Status": "Rejected",
        "Focus": "Cloud Development",
        "Notes": "Application rejected after recruitment process review"
    },
    {
        "Company": "Nokia",
        "Country": "Portugal",
        "City": "Portugal (location not specified)",
        "Position": "Telecommunication Support Trainee - Global Core Network Team",
        "Date Applied": "2026-05-07",
        "Platform": "iAgora / Job Portal",
        "Status": "Applied",
        "Work Mode": "International / Hybrid",
        "Technologies": [
            "IMS Voice",
            "Packet Core",
            "Cloud Infrastructure",
            "Linux",
            "OpenStack",
            "3GPP",
            "VoLTE / VoWiFi / Vo5G"
        ],
        "Focus Area": "Telecommunications / Cloud / Network Engineering",
        "Notes": "Strong technical telecom role with exposure to cloud and network core systems"
    },
    {
        "Company": "Siemens",
        "Country": "Romania",
        "City": "Brasov / Cluj-Napoca / Bucharest",
        "Position": "Siemens Talents Hub Internship 2026 - Data Analytics & Artificial Intelligence",
        "Date Applied": "2026-05-07",
        "Platform": "iAgora / Siemens Careers",
        "Status": "Applied",
        "Focus": "AI, Data Analytics, R&D, Machine Learning, Software Prototyping",
        "Duration": "3 months",
        "Paid": "No",
        "Notes": "R&D internship in AI and Data Analytics"
    },
    {
        "Company": "Novo Nordisk",
        "Country": "Denmark",
        "City": "Bagsværd",
        "Position": "BI, Data & AI Intern - Complaint Insights Team",
        "Date Applied": "2026-05-07",
        "Platform": "Career Portal / Application Form",
        "Status": "Rejected",
        "Focus": "Business Intelligence, Data Analytics, AI, Power BI",
        "Tools": [
            "Power BI",
            "Python",
            "SQL",
            "AI tools"
        ],
        "Duration": "Internship",
        "Paid": "Yes",
        "Notes": "BI + AI internship focused on healthcare data analytics"
    },
    {
    "Company": "NTT DATA",
    "Country": "Spain",
    "City": "A Coruña",
    "Position": "Prácticas Ingeniería Informática - Telecomunicaciones",
    "Date Applied": "2026-05-10",
    "Platform": "NTT DATA Careers / iAgora",
    "Status": "Applied",
    "Focus": "Software Development, Telecommunications, IT",
    "Paid": "Not specified",
    "Notes": "Application confirmed by NTT DATA recruitment team via email"
},

{
    "Company": "BBVA",
    "Country": "Spain",
    "City": "Madrid",
    "Position": "BECAS BBVA ÁREAS DE TECNOLOGÍA E INNOVACIÓN 2026",
    "Date Applied": "2026-05-10",
    "Platform": "BBVA Careers",
    "Status": "Applied",
    "Focus": "Technology, Innovation, IT",
    "Paid": "Not specified",
    "Notes": "Scholarship/internship application confirmed by BBVA recruitment team"
},
{
    "Company": "Natixis",
    "Country": "Portugal",
    "City": "Porto / Lisbon",
    "Position": "IT Trainee | WE WANT YOU 2026 S2",
    "Date Applied": "2026-05-11",
    "Platform": "Natixis Careers",
    "Status": "Applied",
    "Focus": "IT Trainee Program, Development, Data, Cybersecurity, DevOps",
    "Paid": "Not specified",
    "Work Mode": "Hybrid",
    "Notes": "International trainee program focused on IT, software development, data, cybersecurity and analytics"
}
]

# Agregar al Excel
df = pd.concat([df, pd.DataFrame(applications)], ignore_index=True)

# Guardar archivo actualizado
df.to_excel("applications_tracker.xlsx", index=False)

print("Postulaciones agregadas correctamente")

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

# Cargar workbook
wb = load_workbook("applications_tracker.xlsx")
ws = wb.active


# =========================
# AGREGAR / ACTUALIZAR COLUMNA ID (openpyxl) - CORREGIDO
# =========================

from openpyxl import load_workbook

wb = load_workbook("applications_tracker.xlsx")
ws = wb.active

# Eliminar TODAS las columnas que empiecen con "ID" (ID, ID.1, ID.2, ...)
for col_idx in range(ws.max_column, 0, -1):
    header = ws.cell(row=1, column=col_idx).value
    if header and str(header).startswith("ID"):
        ws.delete_cols(col_idx)

# Insertar nueva columna ID al inicio
ws.insert_cols(1)
ws["A1"] = "ID"
for row in range(2, ws.max_row + 1):
    ws[f"A{row}"] = row - 1


# =========================
# ESTILO HEADERS
# =========================
header_fill = PatternFill(start_color="1F4E78",
                          end_color="1F4E78",
                          fill_type="solid")

header_font = Font(color="FFFFFF", bold=True)

for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font

# =========================
# COLORES POR STATUS
# =========================
green_fill = PatternFill(start_color="C6EFCE",
                         end_color="C6EFCE",
                         fill_type="solid")

red_fill = PatternFill(start_color="FFC7CE",
                       end_color="FFC7CE",
                       fill_type="solid")

yellow_fill = PatternFill(start_color="FFEB9C",
                          end_color="FFEB9C",
                          fill_type="solid")

# Buscar columna Status
status_col = None

for cell in ws[1]:
    if cell.value == "Status":
        status_col = cell.column
        break

# Aplicar colores
for row in range(2, ws.max_row + 1):
    status_cell = ws.cell(row=row, column=status_col)

    if status_cell.value == "Applied":
        status_cell.fill = green_fill

    elif status_cell.value == "Rejected":
        status_cell.fill = red_fill

    elif status_cell.value == "Pending":
        status_cell.fill = yellow_fill

# =========================
# AJUSTAR COLUMNAS
# =========================
for column in ws.columns:
    max_length = 0
    column_letter = get_column_letter(column[0].column)

    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass

    adjusted_width = max_length + 3
    ws.column_dimensions[column_letter].width = adjusted_width

# =========================
# FILTROS Y FREEZE
# =========================
ws.auto_filter.ref = ws.dimensions
ws.freeze_panes = "A2"



# =========================
# CONTAR POSTULACIONES RECHAZADAS
# =========================

# Buscar la columna "Status"
status_col = None
for col in range(1, ws.max_column + 1):
    if ws.cell(row=1, column=col).value == "Status":
        status_col = col
        break

# Contar cuántas celdas tienen "Rejected" 
rejected_count = 0
if status_col:
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=status_col).value == "Rejected":
            rejected_count += 1

# Mostrar en consola
print(f"\n Resumen: {rejected_count} postulaciones rechazadas.")

# Escribir el total dentro del Excel 
last_row = ws.max_row + 2
ws.cell(row=last_row, column=1, value="Total Rechazadas:")
ws.cell(row=last_row, column=status_col, value=rejected_count)

# Darle formato 
from openpyxl.styles import Font
ws.cell(row=last_row, column=1).font = Font(bold=True)
ws.cell(row=last_row, column=status_col).font = Font(bold=True)

# =========================
# CONTAR POSTULACIONES APLICADAS
# =========================

# Buscar la columna "Status"
status_col = None
for col in range(1, ws.max_column + 1):
    if ws.cell(row=1, column=col).value == "Status":
        status_col = col
        break

# Contar cuántas celdas tienen "Applied" 
applied_count = 0
if status_col:
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=status_col).value == "Applied":
            applied_count += 1

# Mostrar en consola
print(f"\n Resumen: {applied_count} postulaciones aplicadas.")

# Escribir el total dentro del Excel 
last_row = ws.max_row + 2
ws.cell(row=last_row, column=1, value="Total Aplicadas:")
ws.cell(row=last_row, column=status_col, value=applied_count)

# Darle formato 
from openpyxl.styles import Font
ws.cell(row=last_row, column=1).font = Font(bold=True)
ws.cell(row=last_row, column=status_col).font = Font(bold=True)
applied_count = 0


# Guardar cambios
wb.save("applications_tracker.xlsx")
