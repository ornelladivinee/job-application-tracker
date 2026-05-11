import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

file_path = "applications_tracker.xlsx"

# =========================
# LISTA DE POSTULACIONES 
# =========================
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
    },
    {
    "Company": "TeamViewer",
    "Country": "Portugal",
    "City": "Remote / Hybrid (Portugal / Europe)",
    "Position": "Student, Product Security (all genders)",
    "Date Applied": "2026-05-11",
    "Platform": "TeamViewer Careers",
    "Status": "Applied",
    "Work Mode": "Hybrid / Remote",
    "Focus Area": "Product Security / Cybersecurity / Secure Software Development",
    "Technologies": [
        "Secure Software Development Lifecycle (S-SDLC)",
        "Vulnerability Management",
        "Software Architecture",
        "Programming",
        "Databases",
        "Application Security"
    ],
    "Notes": "Student role focused on Product Security, vulnerability management and secure development practices in a global tech environment. Strong interest in cybersecurity and hands-on learning."
},
    {
    "Company": "Accenture",
    "Country": "Portugal",
    "Position": "Technology Summer Internship",
    "Date Applied": "2026-05-11",
    "Platform": "Accenture Careers",
    "Status": "Applied",
    "Link": "",
    "Notes": "Req ID: R00321525 | Application received"
},
]

# =========================
# REESCRIBIR EXCEL COMPLETO
# =========================
df = pd.DataFrame(applications)

# sobrescribe TODO el archivo
df.to_excel(file_path, index=False)

# =========================
# FORMATO VISUAL
# =========================
wb = load_workbook(file_path)
ws = wb.active

green = PatternFill(
    start_color="C6EFCE",
    end_color="C6EFCE",
    fill_type="solid"
)

red = PatternFill(
    start_color="FFC7CE",
    end_color="FFC7CE",
    fill_type="solid"
)

yellow = PatternFill(
    start_color="FFEB9C",
    end_color="FFEB9C",
    fill_type="solid"
)

# Header bold
for cell in ws[1]:
    cell.font = Font(bold=True)

# Buscar columna Status
status_col = None

for cell in ws[1]:
    if cell.value == "Status":
        status_col = cell.column
        break

# Colores por estado
if status_col:
    for row in range(2, ws.max_row + 1):

        status_cell = ws.cell(row=row, column=status_col)

        if status_cell.value == "Applied":
            status_cell.fill = green

        elif status_cell.value == "Rejected":
            status_cell.fill = red

        elif status_cell.value == "Pending":
            status_cell.fill = yellow

# =========================
# DASHBOARD
# =========================
total = len(df)

applied = len(df[df["Status"] == "Applied"])
rejected = len(df[df["Status"] == "Rejected"])
total = len(df)

applied = len(df[df["Status"] == "Applied"])
rejected = len(df[df["Status"] == "Rejected"])

pending = applied - rejected

print("\nJOB TRACKER")
print(f"Total: {total}")
print(f"Applied: {applied}")
print(f"Rejected: {rejected}")
print(f"Pending: {pending}")


from openpyxl.styles import Font

# =========================
# DASHBOARD EN EXCEL
# =========================

start_row = ws.max_row + 2

ws.cell(row=start_row, column=1, value="JOB TRACKER DASHBOARD").font = Font(bold=True)

ws.cell(row=start_row + 1, column=1, value="Total:")
ws.cell(row=start_row + 1, column=2, value=total)

ws.cell(row=start_row + 2, column=1, value="Applied:")
ws.cell(row=start_row + 2, column=2, value=applied)

ws.cell(row=start_row + 3, column=1, value="Rejected:")
ws.cell(row=start_row + 3, column=2, value=rejected)

ws.cell(row=start_row + 4, column=1, value="Pending:")
ws.cell(row=start_row + 4, column=2, value=pending)
# =========================
# GUARDAR
# =========================
wb.save(file_path)

print("\n✔ Sistema actualizado correctamente")