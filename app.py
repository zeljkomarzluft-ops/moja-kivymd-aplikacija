import streamlit as st
import pandas as pd
import os
import io
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Postavke stranice
st.set_page_config(page_title="Naručivanje Materijala", layout="wide", page_icon="📦")

EXCEL_FILE = "lager_pula.xlsx"

@st.cache_data
def ucitaj_podatke():
    if not os.path.exists(EXCEL_FILE):
        return None
    # Učitavanje Excel tablice
    df = pd.read_excel(EXCEL_FILE)
    # Čišćenje naziva stupaca od eventualnih razmaka
    df.columns = df.columns.str.strip()
    return df

def generiraj_pdf(narudzba_df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = styles['Heading1']
    title_style.alignment = 1  # Centrirano
    
    story.append(Paragraph("<b>NARUDŽBENICA MATERIJALA</b>", title_style))
    story.append(Spacer(1, 20))
    
    # Priprema podataka za tablicu
    table_data = [["Šifra", "Naziv dekora / materijala", "Naručena količina"]]
    for _, row in narudzba_df.iterrows():
        table_data.append([str(row['Šifra']), str(row['Naziv']), str(row['Količina za narudžbu'])])
    
    t = Table(table_data, colWidths=[100, 320, 120])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E88E5")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#F5F5F5")),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor("#BDBDBD")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    
    story.append(t)
    doc.build(story)
    buffer.seek(0)
    return buffer

# --- GLAVNI INTERFEJS ---
st.title("📦 Pregled Lagera i Narudžba Materijala")

df = ucitaj_podatke()

if df is None:
    st.error(f"Datoteka `{EXCEL_FILE}` nije pronađena u repozitoriju. Molimo provjerite naziv i lokaciju datoteke.")
else:
    # Pretpostavka naziva stupaca - prilagodite ako se u Excelu zovu drugačije:
    # Stupci: 'Strana', 'Regal', 'Šifra', 'Naziv', 'Količina', 'Slika URL'
    
    # 1. Odabir strane
    strane = df['Strana'].dropna().unique().tolist() if 'Strana' in df.columns else []
    odabrana_strana = st.selectbox("1. Odaberite stranu skladišta:", ["-- Odaberite --"] + strane)
    
    if odabrana_strana != "-- Odaberite --":
        df_strana = df[df['Strana'] == odabrana_strana]
        
        # 2. Odabir regala
        regali = df_strana['Regal'].dropna().unique().tolist() if 'Regal' in df_strana.columns else []
        odabrani_regal = st.selectbox("2. Odaberite regal:", ["-- Odaberite --"] + regali)
        
        if odabrani_regal != "-- Odaberite --":
            df_filtrirano = df_strana[df_strana['Regal'] == odabrani_regal].copy()
            
            st.subheader(f"Dekori na regalu: {odabrani_regal} (Strana: {odabrana_strana})")
            
            # Inicijalizacija rječnika za količine u session_state
            if 'kolicine' not in st.session_state:
                st.session_state.kolicine = {}

            # Prikaz dekora u tabličnom/kartičnom obliku
            narudžba_lista = []
            
            for idx, row in df_filtrirano.iterrows():
                sifra = row.get('Šifra', 'N/A')
                naziv = row.get('Naziv', 'N/A')
                kolicina_lager = row.get('Količina', 0)
                slika_url = row.get('Slika URL', None)
                
                col1, col2, col3, col4, col5 = st.columns([1.5, 3, 1.5, 2, 2])
                
                with col1:
                    st.write(f"**Šifra:** {sifra}")
                with col2:
                    st.write(f"**{naziv}**")
                with col3:
                    st.write(f"Lager: **{kolicina_lager}**")
                with col4:
                    if pd.notna(slika_url) and str(slika_url).startswith("http"):
                        st.markdown(f"[🖼️ Pogledaj dekor]({slika_url})", unsafe_allow_html=True)
                    else:
                        st.write("-")
                with col5:
                    # Polje za unos količine
                    stanje_kolicina = st.number_input(
                        f"Naruči (kom/m):",
                        min_value=0.0,
                        step=1.0,
                        key=f"input_{sifra}_{idx}"
                    )
                    if stanje_kolicina > 0:
                        narudžba_lista.append({
                            'Šifra': sifra,
                            'Naziv': naziv,
                            'Količina za narudžbu': stanje_kolicina
                        })
                st.divider()

            # Gumb za generiranje PDF-a
            if narudžba_lista:
                df_narudzsba = pd.DataFrame(narudžba_lista)
                st.success(f"Odabrano stavki za narudžbu: **{len(df_narudzsba)}**")
                
                pdf_data = generiraj_pdf(df_narudzsba)
                
                st.download_button(
                    label="📄 Preuzmi PDF Narudžbenicu",
                    data=pdf_data,
                    file_name="narudzbenica_materijala.pdf",
                    mime="application/pdf"
                )
            else:
                st.info("Upisite količinu veću od 0 za generiranje PDF narudžbenice.")