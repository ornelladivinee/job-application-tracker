import pandas as pd

# Estructura del tracker
data = {
    "Company": [],
    "Country": [],
    "Position": [],
    "Date Applied": [],
    "Platform": [],
    "Status": [],
    "Link": [],
    "Notes": [],
    "Follow-up Date": []
}

# Crear tabla
df = pd.DataFrame(data)

# Guardar Excel
df.to_excel("applications_tracker.xlsx", index=False)

print("Excel creado correctamente")