import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns  # Für die Heatmap

# Excel-Datei laden
file_path = "/Users/erxhanganiu/Downloads/Sales.xlsx" 
sheet_name = "Sales"

# Daten laden
data = pd.read_excel(file_path, sheet_name=sheet_name)

# Relevante Spalten extrahieren (ID, Beschaffung, Nutzungsmöglichkeit)
data_filtered = data.iloc[1:, [0, 8, 9]]  # Spalten: ID (A), Nutzungsmöglichkeit, Beschaffung
data_filtered.columns = ["ID", "Nutzungsmöglichkeit", "Beschaffung"]

# Datentypen konvertieren
data_filtered["Beschaffung"] = pd.to_numeric(data_filtered["Beschaffung"], errors="coerce")
data_filtered["Nutzungsmöglichkeit"] = pd.to_numeric(data_filtered["Nutzungsmöglichkeit"], errors="coerce")

# IDs gruppieren nach Beschaffung und Nutzungsmöglichkeit
data_filtered["IDs"] = data_filtered["ID"].astype(str)  # IDs als Strings konvertieren
heatmap_data = data_filtered.groupby(
    ["Nutzungsmöglichkeit", "Beschaffung"]
)["IDs"].apply(", ".join).unstack(fill_value="")  # IDs als Text gruppieren

# Heatmap erstellen
plt.figure(figsize=(12, 8))

# Manuelle Anzeige der Werte in der Heatmap
for y in range(heatmap_data.shape[0]):
    for x in range(heatmap_data.shape[1]):
        plt.text(
            x + 0.5, y + 0.5,  # Position
            heatmap_data.iloc[y, x],  # ID-Text aus der Tabelle
            ha='center', va='center', fontsize=10
        )

# Achsenbeschriftung und Diagrammtitel
plt.xticks(range(len(heatmap_data.columns)), heatmap_data.columns, rotation=45)
plt.yticks(range(len(heatmap_data.index)), heatmap_data.index)
plt.title("Heatmap mit IDs: Beschaffung vs. Nutzungsmöglichkeit", fontsize=15)
plt.xlabel("Beschaffung (1 - sehr schlecht; 5 - sehr gut)", fontsize=12)
plt.ylabel("Nutzungsmöglichkeit (1 - sehr schlecht; 5 - sehr gut)", fontsize=12)

# Rahmen um die Heatmap zeichnen
plt.grid(False)
plt.gca().invert_yaxis()  # Invertiere die Y-Achse, damit die Darstellung korrekt ist
plt.tight_layout()
plt.show()