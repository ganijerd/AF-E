import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from adjustText import adjust_text

# Excel-Datei laden
file_path = "/Users/erxhanganiu/Downloads/Sales.xlsx" 
sheet_name = "Sales"

# Daten laden
data = pd.read_excel(file_path, sheet_name=sheet_name)

# Relevante Spalten extrahieren (ID, Beschaffung, Einzigartigkeit)
data_filtered = data.iloc[1:, [0, 9, 10]]  # Spalten: ID (A), Beschaffung (J), Einzigartigkeit (K)
data_filtered.columns = ["ID", "Beschaffung", "Einzigartigkeit"]

# Datentypen konvertieren
data_filtered["Beschaffung"] = pd.to_numeric(data_filtered["Beschaffung"], errors="coerce")
data_filtered["Einzigartigkeit"] = pd.to_numeric(data_filtered["Einzigartigkeit"], errors="coerce")

# Scatterplot erstellen
plt.figure(figsize=(12, 8))
plt.scatter(
    data_filtered["Beschaffung"], 
    data_filtered["Einzigartigkeit"], 
    s=5000,            # Punktgröße
    c="green",         # Farbe der Punkte
    alpha=0.7         # Transparenz
)

# IDs als Labels hinzufügen
texts = []
for i, label in enumerate(data_filtered["ID"]):
    texts.append(plt.text(
        data_filtered["Beschaffung"].iloc[i],
        data_filtered["Einzigartigkeit"].iloc[i],
        str(label),
        fontsize=10,
        color='white',
        ha='center',
        va='center'
    ))

# Dynamische Anpassung der Label-Positionen
adjust_text(
    texts,
    arrowprops=dict(arrowstyle="->", color='gray', lw=0.5),  # Pfeile für Zuordnung
    only_move={'points': 'y', 'text': 'xy'},  # Texte dürfen sich frei bewegen
)

# Achsenbeschriftung und Titel
plt.xlabel("Beschaffung (1 - sehr schlecht; 5 - sehr gut)", fontsize=12)
plt.ylabel("Einzigartigkeit (1 - sehr üblich; 5 - sehr einzigartig)", fontsize=12)

# X- und Y-Achsen-Ticks auf 1er-Schritte setzen
plt.xticks(range(1, 6, 1))  # X-Achse: Schritte von 1 bis 5 in 1er-Schritten
plt.yticks(range(1, 6, 1))  # Y-Achse: Schritte von 1 bis 5 in 1er-Schritten

plt.title("Scatterplot: Beschaffung vs. Einzigartigkeit (ID's)", fontsize=15)

# Gitter aktivieren
plt.grid(True)

# Plot anzeigen
plt.tight_layout()
plt.show()