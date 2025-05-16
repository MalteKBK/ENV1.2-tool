# Import required libraries
import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz, process
import os

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

# Ensure 'Indikator' columns are strings for accurate matching
bibliotek_df['Indikator'] = bibliotek_df['Indikator'].astype(str).str.strip()
krav_df['Indikator'] = krav_df['Indikator'].astype(str).str.strip()

# Streamlit UI
st.set_page_config(page_title="ENV1.2 Tool", page_icon="🔍", layout="wide")
st.title("ENV1.2 Tool")
query = st.text_input("Indtast produktnavn, producent eller kategori:")

if query:
    results = []
    for index, row in bibliotek_df.iterrows():
        fields = [row.get('Kategori', ''), row.get('Materiale', ''), row.get('Produktnavn', ''), row.get('Producent', '')]
        for field in fields:
            score = fuzz.token_set_ratio(query.lower(), str(field).lower())
            if score >= 80:
                indikator = row.get('Indikator', '').strip()
                kvalitetskrav = dokumentation = None
                if indikator:
                    krav_row = krav_df[krav_df['Indikator'] == indikator]
                    if not krav_row.empty:
                        kvalitetskrav = {f'Opnået kriterie_{i}': krav_row.iloc[0][f'Opnået kriterie_{i}'] for i in range(1, 5)}
                        dokumentation = {f'Dokumentation_{i}': krav_row.iloc[0][f'Dokumentation_{i}'] for i in range(1, 5)}
                
                results.append({
                    'Kategori': row.get('Kategori', ''),
                    'Materiale': row.get('Materiale', ''),
                    'Produktnavn': row.get('Produktnavn', ''),
                    'Producent': row.get('Producent', ''),
                    'Indikator': indikator,
                    'Kvalitetskrav': kvalitetskrav,
                    'Dokumentation': dokumentation,
                    'Score': score
                })
    
    if results:
        for match in sorted(results, key=lambda x: x['Score'], reverse=True):
            st.subheader(match['Produktnavn'])
            st.write(f"**Kategori:** {match['Kategori']}")
            st.write(f"**Materiale:** {match['Materiale']}")
            st.write(f"**Producent:** {match['Producent']}")
            st.write(f"**Indikator:** {match['Indikator']}")
            st.write(f"**Opnået kriterie:** {match['kvalitetskrav']}")
            st.write(f"**Score:** {match['Score']}%")

    else:
        st.write("Ingen resultater fundet.")
