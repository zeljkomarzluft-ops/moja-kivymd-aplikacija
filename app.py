import streamlit as st
import pandas as pd
import os
import io
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# Postavke stranice
st.set_page_config(page_title="Prijedlog Narudžbe Materijala - Elgrad", layout="wide", page_icon="📦")

EXCEL_FILE = "lager_pula.xlsx"

@st.cache_data
def ucitaj_podatke():
    if not os.path.exists(EXCEL_FILE):
        return None
    try:
        # Učitavanje Excel datoteke
        df = pd.read_excel(EXCEL_FILE)
        
        # Čišćenje naziva stupaca (uklanjanje praznina)
        df.columns = df.columns.astype(str).str.strip()
        
        # Mapiranje točnih stupaca prema vašoj tablici
        mapa_stupaca = {}
        for col in df.columns:
            c_low = col.lower()
            if c_low == 'šifra' or c_low == 'sifra':
                mapa_stupaca[col] = 'Šifra'
            elif c_low == 'naziv':
                mapa_stupaca[col] = 'Naziv'
            elif c_low == 'strana':
                mapa_stupaca[col] = 'Strana'
            elif c_low == 'regal':
                mapa_stupaca[col] = 'Regal'
            elif c_low == 'polica':
                mapa_stupaca[col] = 'Polica'
            elif c_low == 'lager':
                mapa_stupaca[col] = 'Lager'
            elif 'slika' in c_low or 'url' in c_low:
                mapa_stupaca[col] = 'Slika_URL'
        
        df = df.rename(columns=mapa_stupaca)
        return df
    except Exception as e:
        st.error(f"Greška pri čitanju Excel datoteke: {e}")
        return None

def generiraj_pdf(narudzba_df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = styles['Heading1']
    title_style.alignment = 1  # Centrirano
    
    # Promijenjen naslov u PDF dokumentu
    story.append(Paragraph("<b>PRIJEDLOG NARUDŽBE</b>", title_style))
    story.append(Spacer(1, 20))
    
    # Priprema podataka za tablicu PDF-a
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
st.title("📦 Pregled Lagera i Prijedlog Narudžbe")

df = ucitaj_podatke()

if df is None:
    st.error(f"Datoteka `{EXCEL_FILE}` nije pronađena u repozitoriju ili je neispravna.")
else:
    # Gumb za ručno osvježavanje ako mijenjate Excel
    col_refresh, _ = st.columns([2, 8])
    with col_refresh:
        if st.button("🔄 Osvježi podatke iz Excela"):
            st.cache_data.clear()
            st.rerun()

    # 1. Odabir strane
    if 'Strana' in df.columns:
        strane = sorted([str(x).strip() for x in df['Strana'].dropna().unique()])
        odabrana_strana = st.selectbox("1. Odaberite stranu skladišta:", ["-- Odaberite --"] + strane)
    else:
        st.error("Stupac 'Strana' nije pronađen u Excelu.")
        odabrana_strana = "-- Odaberite --"
    
    if odabrana_strana != "-- Odaberite --":
        df_strana = df[df['Strana'].astype(str).str.strip() == odabrana_strana]
        
        # 2. Odabir regala
        if 'Regal' in df_strana.columns:
            regali_raw = df_strana['Regal'].dropna().unique()
            try:
                regali = sorted([int(x) for x in regali_raw if str(x).isdigit()])
            except:
                regali = sorted([str(x) for x in regali_raw])
                
            odabrani_regal = st.selectbox("2. Odaberite regal:", ["-- Odaberite --"] + [str(r) for r in regali])
        else:
            st.error("Stupac 'Regal' nije pronađen u Excelu.")
            odabrani_regal = "-- Odaberite --"
        
        if odabrani_regal != "-- Odaberite --":
            df_filtrirano = df_strana[df_strana['Regal'].astype(str).str.strip() == str(odabrani_regal)].copy()
            
            st.subheader(f"📍 Dekori na regalu: {odabrani_regal} (Strana: {odabrana_strana})")
            
            narudžba_lista = []
            
            # Zaglavlje tablice
            h1, h2, h3, h4, h5, h6 = st.columns([1.5, 3, 1, 1, 1.5, 2])
            h1.markdown("**Šifra**")
            h2.markdown("**Naziv dekora**")
            h3.markdown("**Polica**")
            h4.markdown("**Lager**")
            h5.markdown("**Slika / Dekor**")
            h6.markdown("**Količina za narudžbu**")
            st.divider()

            for idx, row in df_filtrirano.iterrows():
                sifra = str(row.get('Šifra', 'N/A'))
                naziv = str(row.get('Naziv', 'N/A'))
                polica = str(row.get('Polica', '-'))
                lager = str(row.get('Lager', 0))
                slika_url = row.get('Slika_URL', None)
                
                col1, col2, col3, col4, col5, col6 = st.columns([1.5, 3, 1, 1, 1.5, 2])
                
                with col1:
                    st.write(f"**{sifra}**")
                with col2:
                    st.write(f"{naziv}")
                with col3:
                    st.write(f"Polica: **{polica}**")
                with col4:
                    st.write(f"**{lager}**")
                with col5:
                    if pd.notna(slika_url) and str(slika_url).strip().startswith("http"):
                        url_str = str(slika_url).strip()
                        st.image(url_str, width=70)
                        st.markdown(f"[🔍 Puni prikaz]({url_str})", unsafe_allow_html=True)
                    else:
                        st.write("-")
                with col6:
                    stanje_kolicina = st.number_input(
                        f"Naruči za {sifra}:",
                        min_value=0.0,
                        step=1.0,
                        key=f"input_{sifra}_{idx}",
                        label_visibility="collapsed"
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
                
                # Promijenjen naziv gumba i datoteke
                st.download_button(
                    label="📄 Preuzmi Prijedlog narudžbe",
                    data=pdf_data,
                    file_name="prijedlog_narudzbe.pdf",
                    mime="application/pdf"
                )
            else:
                st.info("Upišite količinu veću od 0 u polja iznad za kreiranje prijedloga narudžbe.")
