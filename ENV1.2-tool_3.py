# Import required libraries
import pandas as pd
from fuzzywuzzy import fuzz, process
import os
import tkinter as tk
from tkinter import ttk, messagebox

# File paths (adjust as necessary)
BIBLIOTEK_FILE = 'ENV1 - Bibliotek.xlsx'
KRAV_FILE = 'Krav.xlsx'

# Load the data
bibliotek_df = pd.read_excel(BIBLIOTEK_FILE, sheet_name='Bibliotek')
krav_df = pd.read_excel(KRAV_FILE, sheet_name='ENV1.2 Database', header=2)

# Clean up the Krav data
krav_df.columns = [
    'Indikator', 'Relevante_bygningsdele', 'Produkttype', 'Kvalitetstrin_1', 
    'Kvalitetstrin_2', 'Kvalitetstrin_3', 'Kvalitetstrin_4', 'Empty', 
    'Dokumentation_1', 'Dokumentation_2', 'Dokumentation_3', 'Dokumentation_4', 'Extra_Column'
]
krav_df.drop(columns=['Empty', 'Extra_Column'], inplace=True)
krav_df.dropna(how='all', inplace=True)

# Modify the function to always display 'Indikator' from bibliotek_df, even without a match in krav_df
def search_material(query, threshold=80):
    results = []

    for index, row in bibliotek_df.iterrows():
        # Check for exact or partial matches in product name, category, or manufacturer
        fields = [
            row.get('Kategori', ''),
            row.get('Materiale', ''),
            row.get('Produktnavn', ''),
            row.get('Producent', '')
        ]
        for field in fields:
            score = fuzz.token_set_ratio(query.lower(), str(field).lower())
            if score >= threshold:
                # Extract indicator and quality step from Bibliotek
                indikator = row.get('Indikator', '')
                print(f"Debug: Bibliotek Indikator = {indikator}")  # Debugging line
                kvalitetstrin_bibliotek = row.get('Kvalitetstrin', '')  # Assuming column G is named 'Kvalitetstrin'

                # No longer require a match in krav_df to display 'Indikator'
                results.append({
                    'Kategori': row.get('Kategori', ''),
                    'Materiale': row.get('Materiale', ''),
                    'Produktnavn': row.get('Produktnavn', ''),
                    'Producent': row.get('Producent', ''),
                    'Indikator': row.get(' Indikator', ''),
                    'Kvalitetstrin_Bibliotek': row.get('Opnået kriterie', ''),
                    'Kvalitetskrav': None,  # Removed dependency on krav_df
                    'Dokumentation': None,  # Removed dependency on krav_df
                    'Score': score
                })
    return sorted(results, key=lambda x: x['Score'], reverse=True)

# Bind the Enter key to trigger the search function and enhance the GUI appearance
# Update the GUI display function to include 'Kvalitetstrin_Bibliotek'
def search_and_display():
    query = search_entry.get().strip()
    if not query:
        messagebox.showwarning("Input Error", "Please enter a search query.")
        return

    matches = search_material(query)
    results_text.delete(1.0, tk.END)  # Clear previous results

    if matches:
        results_text.insert(tk.END, f"Found {len(matches)} possible matches:\n\n")
        for match in matches:
            results_text.insert(tk.END, f"Kategori: {match['Kategori']}\n")
            results_text.insert(tk.END, f"Materiale: {match['Materiale']}\n")
            results_text.insert(tk.END, f"Produktnavn: {match['Produktnavn']}\n")
            results_text.insert(tk.END, f"Producent: {match['Producent']}\n")
            results_text.insert(tk.END, f"Indikator: {match['Indikator']}\n")
            results_text.insert(tk.END, f"Kvalitetstrin: {match['Kvalitetstrin_Bibliotek']}\n")
            results_text.insert(tk.END, f"Match Score: {match['Score']}%\n")
            if match['Kvalitetskrav']:
                results_text.insert(tk.END, "Kvalitetskrav:\n")
                for level, krav in match['Kvalitetskrav'].items():
                    results_text.insert(tk.END, f"  {level}: {krav}\n")
            if match['Dokumentation']:
                results_text.insert(tk.END, "Dokumentation:\n")
                for level, doc in match['Dokumentation'].items():
                    results_text.insert(tk.END, f"  {level}: {doc}\n")
            results_text.insert(tk.END, "\n")
    else:
        results_text.insert(tk.END, "No close matches found.")

# Update the GUI to center the search bar with rounded corners and add a title above it
# Create the main GUI window
root = tk.Tk()
root.title("ENV1.2-tool")
root.configure(bg="#0f172a")  # Set dark background color

# Create and place widgets
frame = tk.Frame(root, bg="#0f172a", padx=20, pady=20)
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Add a title above the search bar
title_label = tk.Label(frame, text="ENV1.2-tool", font=("Arial", 20, "bold"), bg="#0f172a", fg="#e2e8f0")
title_label.grid(row=0, column=0, pady=(0, 10), sticky=tk.N)

# Add subtitle
description_label = tk.Label(frame, text="Søg efter produktnavn, producent eller byggevaretype", font=("Arial", 12, "italic"), bg="#0f172a", fg="#e2e8f0")
description_label.grid(row=1, column=0, pady=(0, 20), sticky=tk.N)

# Create a rounded search bar
search_frame = tk.Frame(frame, bg="#e2e8f0", bd=0, relief="flat", highlightthickness=0)
search_frame.grid(row=2, column=0, pady=10)
search_frame.configure(width=400, height=40)
search_frame.grid_propagate(False)

search_entry = tk.Entry(search_frame, font=("Arial", 12), bg="#e2e8f0", bd=0, relief="flat")
search_entry.pack(fill=tk.BOTH, expand=True, padx=10)

# Add a search icon
search_icon = tk.Label(search_frame, text="🔍", bg="#e2e8f0", font=("Arial", 12))
search_icon.place(relx=0.95, rely=0.5, anchor=tk.CENTER)

# Add a results area
results_text = tk.Text(frame, width=80, height=20, wrap=tk.WORD, font=("Arial", 10), bg="#1e293b", fg="#e2e8f0", bd=0, relief="flat")
results_text.grid(row=3, column=0, pady=10)

# Center the search bar and title
frame.grid_columnconfigure(0, weight=1)

# Bind the Enter key to trigger the search function
root.bind('<Return>', lambda event: search_and_display())

# Enhance the visual appearance with padding and spacing
for widget in frame.winfo_children():
    widget.grid_configure(padx=5, pady=5)

# Start the GUI event loop
root.mainloop()
