import streamlit as st
import pandas as pd
import os
import io
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# Postavke stranice
st.set_page_config(page_title="Upravljanje Skladištem - Elgrad", layout="wide", page_icon="📦")

EXCEL_FILE = "lager_pula.xlsx"

# Trajna memorija za košaricu narudžbe
if "cart" not in st.session_state:
    st.session_state.cart = {}

@st.cache_data
def ucitaj_podatke():
    if not os.path.exists(EXCEL_FILE):
        return None
    try:
        # Učitavanje Excel datoteke
        df = pd.read_excel(EXCEL_FILE)
        df.columns = df.columns.astype(str).str.strip()
        
        mapa_stupaca = {}
        for col in df.columns:
            c_low = col.lower()
            if c_low in ['šifra', 'sifra']:
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
            elif 'faktor' in c_low:
                mapa_stupaca[col] = 'Faktor'
            elif 'mjerna' in c_low or 'jedinica' in c_low or c_low == 'jm':
                mapa_stupaca[col] = 'Mjerna_Jedinica'
        
        df = df.rename(columns=mapa_stupaca)
        
        # Obrada kolone Faktor (zamjena zareza u tačku radi pretvaranja u broj)
        if 'Faktor' in df.columns:
            df['Faktor'] = df['Faktor'].astype(str).str.replace(',', '.', regex=False)
            df['Faktor'] = pd.to_numeric(df['Faktor'], errors='coerce').fillna(1.0)
        else:
            df['Faktor'] = 1.0

        # Obrada kolone Mjerna Jedinica
        if 'Mjerna_Jedinica' not in df.columns:
            df['Mjerna_Jedinica'] = 'm²'
        else:
            df['Mjerna_Jedinica'] = df['Mjerna_Jedinica'].fillna('m²').astype(str)
            
        return df
    except Exception as e:
        st.error(f"Greška pri čitanju Excel datoteke: {e}")
        return None

def generiraj_pdf_narudzba(narudzba_list):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=30)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = styles['Heading1']
    title_style.alignment = 1
    
    story.append(Paragraph("<b>PRIJEDLOG NARUDŽBE MATERIJALA</b>", title_style))
    story.append(Spacer(1, 20))
    
    table_data = [["Šifra", "Naziv dekora / materijala", "Br. Kom / Ploča", "Ukupno"]]
    for item in narudzba_list:
        table_data.append([
            str(item['Šifra']),
            str(item['Naziv']),
            f"{item['Količina']:.0f}",
            f"{item['Ukupno_M2']:.2f} {item['Jedinica']}"
        ])
    
    t = Table(table_data, colWidths=[90, 270, 100, 110])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E88E5")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#F5F5F5")),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor("#BDBDBD")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    
    story.append(t)
    doc.build(story)
    buffer.seek(0)
    return buffer

def generiraj_pdf_inventura(inventura_list, info_regal):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=30)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = styles['Heading1']
    title_style.alignment = 1
    
    story.append(Paragraph("<b>IZVJEŠĆE INVENTURE / KONTROLE REGALA</b>", title_style))
    story.append(Paragraph(f"<font size=10>Lokacija: {info_regal}</font>", styles['Normal']))
    story.append(Spacer(1, 15))
    
    table_data = [["Šifra", "Naziv dekora", "Brojano (kom)", "Ukupno", "Napomena / Bilješka"]]
    for item in inventura_list:
        table_data.append([
            str(item['Šifra']),
            str(item['Naziv']),
            f"{item['Brojano']:.0f}",
            f"{item['Ukupno_M2']:.2f} {item['Jedinica']}",
            ""
        ])
    
    t = Table(table_data, colWidths=[80, 210, 80, 100, 100])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#388E3C")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('TOPPADDING', (0, 1), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#F1F8E9")),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor("#A5D6A7")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    
    story.append(t)
    doc.build(story)
    buffer.seek(0)
    return buffer

# --- GLAVNI INTERFEJS ---
st.title("📦 Skladište - Prijedlog Narudžbe i Inventura")

df = ucitaj_podatke()

if df is None:
    st.error(f"Datoteka `{EXCEL_FILE}` nije pronađena u repozitoriju ili je neispravna.")
else:
    with st.sidebar:
        st.header("⚙️ Opcije i Alati")
        if st.button("🔄 Osvježi podatke iz Excela", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
            
        st.divider()
        st.subheader("➕ Dodaj novi dekor")
        with st.form("forma_novi_dekor"):
            novo_sifra = st.text_input("Šifra dekora*")
            novo_naziv = st.text_input("Naziv dekora*")
            novo_strana = st.text_input("Strana (npr. DESNA)")
            novo_regal = st.text_input("Regal (npr. 5)")
            novo_polica = st.text_input("Polica (npr. A)")
            novo_lager = st.number_input("Početni Lager", min_value=0.0, step=1.0)
            novo_faktor = st.number_input("Faktor (m²/m' po komadu)", min_value=0.001, value=5.796, format="%.3f")
            novo_jedinica = st.selectbox("Mjerna jedinica", ["m²", "m'", "kom"])
            novo_url = st.text_input("Slika URL")
            
            submit_dekor = st.form_submit_button("Spremi u tablicu")
            if submit_dekor:
                if novo_sifra and novo_naziv:
                    novi_red = pd.DataFrame([{
                        'Šifra': novo_sifra,
                        'Naziv': novo_naziv,
                        'Strana': novo_strana if novo_strana else 'DESNA',
                        'Regal': novo_regal if novo_regal else '1',
                        'Polica': novo_polica if novo_polica else 'A',
                        'Lager': novo_lager,
                        'Faktor': novo_faktor,
                        'Mjerna_Jedinica': novo_jedinica,
                        'Slika_URL': novo_url
                    }])
                    try:
                        df_trenutni = pd.read_excel(EXCEL_FILE)
                        df_osvjezeni = pd.concat([df_trenutni, novi_red], ignore_index=True)
                        df_osvjezeni.to_excel(EXCEL_FILE, index=False)
                        st.cache_data.clear()
                        st.success("Novi dekor uspješno dodan!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Greška pri spremanju u Excel: {e}")
                else:
                    st.warning("Upišite barem Šifru i Naziv dekora.")

    tab_narudzba, tab_inventura = st.tabs(["📋 Prijedlog Narudžbe", "📊 Inventura / Kontrola Regala"])

    # ==========================================
    # TAB 1: PRIJEDLOG NARUDŽBE
    # ==========================================
    with tab_narudzba:
        col_s, col_r = st.columns(2)
        with col_s:
            strane = sorted([str(x).strip() for x in df['Strana'].dropna().unique()])
            odabrana_strana = st.selectbox("Odaberite stranu skladišta:", ["-- Odaberite --"] + strane, key="narudzba_strana")
            
        with col_r:
            if odabrana_strana != "-- Odaberite --":
                df_strana = df[df['Strana'].astype(str).str.strip() == odabrana_strana]
                regali_raw = df_strana['Regal'].dropna().unique()
                try:
                    regali = sorted([int(x) for x in regali_raw if str(x).isdigit()])
                except:
                    regali = sorted([str(x) for x in regali_raw])
                odabrani_regal = st.selectbox("Odaberite regal:", ["-- Odaberite --"] + [str(r) for r in regali], key="narudzba_regal")
            else:
                odabrani_regal = "-- Odaberite --"

        if odabrana_strana != "-- Odaberite --" and odabrani_regal != "-- Odaberite --":
            df_filtrirano = df[(df['Strana'].astype(str).str.strip() == odabrana_strana) & 
                               (df['Regal'].astype(str).str.strip() == str(odabrani_regal))].copy()
            
            st.info(f"📍 Prikaz dekora na regalu: **{odabrani_regal}** (Strana: **{odabrana_strana}**)")
            
            h1, h2, h3, h4, h5, h6 = st.columns([1.5, 3, 1, 1, 1.5, 2])
            h1.markdown("**Šifra**")
            h2.markdown("**Naziv dekora**")
            h3.markdown("**Polica / Faktor**")
            h4.markdown("**Lager**")
            h5.markdown("**Slika**")
            h6.markdown("**Naruči komada**")
            st.divider()

            for idx, row in df_filtrirano.iterrows():
                sifra = str(row.get('Šifra', 'N/A'))
                naziv = str(row.get('Naziv', 'N/A'))
                polica = str(row.get('Polica', '-'))
                lager = str(row.get('Lager', 0))
                faktor = float(row.get('Faktor', 1.0))
                jedinica = str(row.get('Mjerna_Jedinica', 'm²'))
                slika_url = row.get('Slika_URL', None)
                
                col1, col2, col3, col4, col5, col6 = st.columns([1.5, 3, 1, 1, 1.5, 2])
                
                with col1:
                    st.write(f"**{sifra}**")
                with col2:
                    st.write(f"{naziv}")
                with col3:
                    st.write(f"Polica: **{polica}**\n\n({faktor} {jedinica})")
                with col4:
                    st.write(f"**{lager}**")
                with col5:
                    if pd.notna(slika_url) and str(slika_url).strip().startswith("http"):
                        st.image(str(slika_url).strip(), width=60)
                    else:
                        st.write("-")
                with col6:
                    trenutna_val = st.session_state.cart.get(sifra, {}).get('Količina', 0.0)
                    
                    unos_kolicina = st.number_input(
                        f"Naruči {sifra}:",
                        min_value=0.0,
                        step=1.0,
                        value=float(trenutna_val),
                        key=f"cart_in_{sifra}",
                        label_visibility="collapsed"
                    )
                    
                    if unos_kolicina > 0:
                        st.session_state.cart[sifra] = {
                            'Šifra': sifra,
                            'Naziv': naziv,
                            'Količina': unos_kolicina,
                            'Ukupno_M2': unos_kolicina * faktor,
                            'Jedinica': jedinica
                        }
                    elif sifra in st.session_state.cart and unos_kolicina == 0:
                        del st.session_state.cart[sifra]
                st.divider()

        st.subheader("🛒 Ukupno odabrano za narudžbu")
        stvarne_stavke = list(st.session_state.cart.values())
        if stvarne_stavke:
            df_kosarica = pd.DataFrame(stvarne_stavke).rename(columns={'Ukupno_M2': 'Ukupna Količina'})
            st.dataframe(df_kosarica, use_container_width=True)
            
            col_pdf, col_clear = st.columns([3, 1])
            with col_pdf:
                pdf_data = generiraj_pdf_narudzba(stvarne_stavke)
                st.download_button(
                    label="📄 Preuzmi Prijedlog narudžbe (PDF)",
                    data=pdf_data,
                    file_name="prijedlog_narudzbe.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            with col_clear:
                if st.button("🧹 Očisti sve uneseno", use_container_width=True):
                    st.session_state.cart = {}
                    st.rerun()
        else:
            st.info("Košarica je trenutno prazna.")

    # ==========================================
    # TAB 2: INVENTURA / BROJANO U SKLADIŠTU
    # ==========================================
    with tab_inventura:
        col_is, col_ir = st.columns(2)
        with col_is:
            strane_inv = sorted([str(x).strip() for x in df['Strana'].dropna().unique()])
            inv_strana = st.selectbox("Odaberite stranu:", ["-- Odaberite --"] + strane_inv, key="inv_strana")
            
        with col_ir:
            if inv_strana != "-- Odaberite --":
                df_inv_strana = df[df['Strana'].astype(str).str.strip() == inv_strana]
                regali_inv_raw = df_inv_strana['Regal'].dropna().unique()
                try:
                    regali_inv = sorted([int(x) for x in regali_inv_raw if str(x).isdigit()])
                except:
                    regali_inv = sorted([str(x) for x in regali_inv_raw])
                inv_regal = st.selectbox("Odaberite regal za brojanje:", ["-- Odaberite --"] + [str(r) for r in regali_inv], key="inv_regal")
            else:
                inv_regal = "-- Odaberite --"

        if inv_strana != "-- Odaberite --" and inv_regal != "-- Odaberite --":
            df_inv_filtrirano = df[(df['Strana'].astype(str).str.strip() == inv_strana) & 
                                   (df['Regal'].astype(str).str.strip() == str(inv_regal))].copy()
            
            st.success(f"📋 Inventurna lista za: **Regal {inv_regal}** (Strana {inv_strana})")
            
            h1, h2, h3, h4 = st.columns([1.5, 3.5, 1.5, 1.5])
            h1.markdown("**Šifra**")
            h2.markdown("**Naziv dekora**")
            h3.markdown("**Brojano (kom)**")
            h4.markdown("**Ukupno**")
            st.divider()

            inventura_lista = []

            for idx, row in df_inv_filtrirano.iterrows():
                sifra = str(row.get('Šifra', 'N/A'))
                naziv = str(row.get('Naziv', 'N/A'))
                faktor = float(row.get('Faktor', 1.0))
                jedinica = str(row.get('Mjerna_Jedinica', 'm²'))
                
                col1, col2, col3, col4 = st.columns([1.5, 3.5, 1.5, 1.5])
                
                with col1:
                    st.write(f"**{sifra}**")
                with col2:
                    st.write(f"{naziv}")
                with col3:
                    brojano = st.number_input(
                        f"Brojano {sifra}:",
                        min_value=0.0,
                        step=1.0,
                        value=0.0,
                        key=f"inv_in_{sifra}_{idx}",
                        label_visibility="collapsed"
                    )
                with col4:
                    ukupno_m2 = brojano * faktor
                    st.write(f"**{ukupno_m2:.2f} {jedinica}**")
                
                inventura_lista.append({
                    'Šifra': sifra,
                    'Naziv': naziv,
                    'Brojano': brojano,
                    'Ukupno_M2': ukupno_m2,
                    'Jedinica': jedinica
                })
                st.divider()

            if inventura_lista:
                pdf_inv_data = generiraj_pdf_inventura(inventura_lista, f"Regal {inv_regal} ({inv_strana})")
                st.download_button(
                    label="📄 Preuzmi Izvješće Inventure (PDF)",
                    data=pdf_inv_data,
                    file_name=f"inventura_regal_{inv_regal}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
