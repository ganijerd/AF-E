import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Excel-Datei laden
file_path = "/Users/erxhanganiu/Downloads/Sales.xlsx"  # Passe den Pfad zur Datei an
sheet_name = "Sales"

# Daten laden
data = pd.read_excel(file_path, sheet_name=sheet_name)

# Relevante Spalten extrahieren (ID, Beschaffung, Einzigartigkeit)
data_filtered = data.iloc[1:, [0, 9, 10]]  # Spalten: ID (A), Beschaffung, Einzigartigkeit
data_filtered.columns = ["ID", "Beschaffung", "Einzigartigkeit"]

# Datentypen konvertieren
data_filtered["Beschaffung"] = pd.to_numeric(data_filtered["Beschaffung"], errors="coerce")
data_filtered["Einzigartigkeit"] = pd.to_numeric(data_filtered["Einzigartigkeit"], errors="coerce")

# Scatterplot erstellen
plt.figure(figsize=(12, 8))
plt.scatter(
    data_filtered["Beschaffung"], 
    data_filtered["Einzigartigkeit"], 
    s=100,            # Punktgröße
    c="green",        # Farbe der Punkte
    alpha=0.7         # Transparenz
)

# IDs als Labels hinzufügen, mit einem leichten zufälligen Offset
np.random.seed(42)  # Für reproduzierbare Ergebnisse
x_offsets = np.random.uniform(-0.1, 0.1, len(data_filtered))
y_offsets = np.random.uniform(-0.1, 0.1, len(data_filtered))

for i, label in enumerate(data_filtered["ID"]):
    plt.text(
        data_filtered["Beschaffung"].iloc[i] + x_offsets[i],
        data_filtered["Einzigartigkeit"].iloc[i] + y_offsets[i],
        str(label),
        fontsize=15,  # Schriftgröße der Labels
        ha='center',
        va='center'
    )

# Achsenbeschriftung
plt.xlabel("Beschaffung (1 - sehr schlecht; 5 - sehr gut)", fontsize=12)
plt.ylabel("Einzigartigkeit (1 - sehr üblich; 5 - sehr einzigartig)", fontsize=12)

# X- und Y-Achse auf 1er-Schritte setzen
plt.xticks(range(1, 6, 1))  # X-Achse: Schritte von 1 bis 5
plt.yticks(range(1, 6, 1))  # Y-Achse: Schritte von 1 bis 5

# Titel und Gitter
plt.title("Beschaffung & Einzigartigkeit (mit ID-Labels)", fontsize=15)
plt.grid(True)

# Plot anzeigen
plt.tight_layout()
plt.show()