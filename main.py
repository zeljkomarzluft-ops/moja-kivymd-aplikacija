import os
import pandas as pd
from datetime import datetime

from kivy.lang import Builder
from kivy.core.window import Window
from kivymd.app import MDApp
from kivymd.uix.screen import MDScreen
from kivymd.uix.list import MDList, TwoLineAvatarIconListItem, IconLeftWidget
from kivymd.uix.textfield import MDTextField
from kivymd.uix.button import MDRaisedButton
from kivymd.uix.dialog import MDDialog
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.gridlayout import MDGridLayout
from kivymd.uix.label import MDLabel
from kivymd.uix.scrollview import MDScrollView
from kivy.uix.image import AsyncImage

# ReportLab za PDF
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

Window.size = (400, 720)

EXCEL_FILE = "lager pula sa lokacijom i slikom.xlsx"

class SkladisteApp(MDApp):
    def build(self):
        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "Orange"
        
        self.odabrana_strana = None
        self.odabrani_regal = None
        self.kosarica = {}  # {kod: {"naziv": ..., "kolicina": ...}}
        
        self.ucitaj_podatke()
        
        return Builder.load_string('''
MDScreen:
    MDBoxLayout:
        orientation: 'vertical'
        
        MDTopAppBar:
            title: "Skladište Pula - Picking"
            elevation: 4
            left_action_items: [["home", lambda x: app.prikazi_odabir_strana()]]
            right_action_items: [["cart", lambda x: app.otvori_kosaricu()], ["file-pdf-box", lambda x: app.generiraj_pdf()]]
            
        MDBoxLayout:
            id: content_area
            orientation: 'vertical'
            padding: "10dp"
            spacing: "10dp"
''')

    def ucitaj_podatke(self):
        if os.path.exists(EXCEL_FILE):
            try:
                self.df = pd.read_excel(EXCEL_FILE)
                # Očisti nazive stupaca od praznih mjesta
                self.df.columns = [str(c).strip() for c in self.df.columns]
                print("Excel uspješno učitan!")
            except Exception as e:
                print(f"Greška pri učitavanju Excela: {e}")
                self.df = pd.DataFrame()
        else:
            print(f"Datoteka '{EXCEL_FILE}' nije pronađena!")
            self.df = pd.DataFrame()

    def on_start(self):
        self.prikazi_odabir_strana()

    def prikazi_odabir_strana(self):
        content = self.root.ids.content_area
        content.clear_widgets()
        
        lbl = MDLabel(
            text="Odaberite stranu / zonu skladišta:",
            font_style="H6",
            size_hint_y=None,
            height="40dp"
        )
        content.add_widget(lbl)
        
        strane = [
            ("LIJEVA STRANA", "LIJEVA"),
            ("DESNA STRANA", "DESNA"),
            ("LIJEVA KRP", "LIJEVA KRP"),
            ("DESNA KRP", "DESNA KRP"),
            ("EKSPEDIT KRP", "EKSPEDIT KRP")
        ]
        
        scroll = MDScrollView()
        box = MDBoxLayout(orientation='vertical', spacing="10dp", size_hint_y=None)
        box.bind(minimum_height=box.setter('height'))
        
        for naziv_gumba, kljuc in strane:
            btn = MDRaisedButton(
                text=naziv_gumba,
                size_hint=(1, None),
                height="50dp",
                on_release=lambda x, k=kljuc: self.odaberi_stranu(k)
            )
            box.add_widget(btn)
            
        scroll.add_widget(box)
        content.add_widget(scroll)

    def odaberi_stranu(self, strana):
        self.odabrana_strana = strana
        self.prikazi_odabir_regala()

    def prikazi_odabir_regala(self):
        content = self.root.ids.content_area
        content.clear_widgets()
        
        top_box = MDBoxLayout(orientation='horizontal', size_hint_y=None, height="40dp")
        lbl = MDLabel(
            text=f"Strana: {self.odabrana_strana} -> Odaberite regal:",
            font_style="Subtitle1"
        )
        btn_back = MDRaisedButton(text="NATRAG", on_release=lambda x: self.prikazi_odabir_strana())
        top_box.add_widget(lbl)
        top_box.add_widget(btn_back)
        content.add_widget(top_box)
        
        grid = MDGridLayout(cols=3, spacing="10dp", size_hint_y=None)
        grid.bind(minimum_height=grid.setter('height'))
        
        # Dinamičko dohvaćanje SAMO onih regala koji stvarno postoje u Excelu za odabranu stranu
        postojeci_regali = []
        if not self.df.empty and 'Strana' in self.df.columns and 'Regal' in self.df.columns:
            strana_mask = self.df['Strana'].astype(str).str.strip().str.upper() == self.odabrana_strana.strip().upper()
            regali_serija = self.df[strana_mask]['Regal'].dropna().unique()
            
            for r in regali_serija:
                r_str = str(r).strip()
                if r_str and r_str.lower() != 'nan':
                    if r_str.endswith('.0'):
                        r_str = r_str[:-2]
                    postojeci_regali.append(r_str)
            
            try:
                postojeci_regali = sorted(list(set(postojeci_regali)), key=lambda x: int(x))
            except ValueError:
                postojeci_regali = sorted(list(set(postojeci_regali)))

        if not postojeci_regali:
            lbl_prazno = MDLabel(text="Nema pronađenih regala za ovu stranu.", halign="center")
            content.add_widget(lbl_prazno)
            return

        for regal_id in postojeci_regali:
            btn = MDRaisedButton(
                text=f"Regal {regal_id}",
                size_hint_y=None,
                height="50dp",
                on_release=lambda x, r=regal_id: self.odaberi_regal(r)
            )
            grid.add_widget(btn)
            
        scroll = MDScrollView()
        scroll.add_widget(grid)
        content.add_widget(scroll)

    def odaberi_regal(self, regal):
        self.odabrani_regal = regal
        self.prikazi_artikle()

    def prikazi_artikle(self):
        content = self.root.ids.content_area
        content.clear_widgets()
        
        top_box = MDBoxLayout(orientation='horizontal', size_hint_y=None, height="40dp")
        lbl = MDLabel(
            text=f"Lokacija: {self.odabrana_strana} - Regal {self.odabrani_regal}",
            font_style="Subtitle1"
        )
        btn_back = MDRaisedButton(text="NATRAG", on_release=lambda x: self.prikazi_odabir_regala())
        top_box.add_widget(lbl)
        top_box.add_widget(btn_back)
        content.add_widget(top_box)
        
        scroll = MDScrollView()
        list_view = MDList()
        
        filtrirano = pd.DataFrame()
        
        if not self.df.empty and 'Strana' in self.df.columns and 'Regal' in self.df.columns:
            strana_mask = self.df['Strana'].astype(str).str.strip().str.upper() == self.odabrana_strana.strip().upper()
            
            regal_odabrani_str = str(self.odabrani_regal).strip()
            regal_mask = self.df['Regal'].astype(str).str.strip().apply(
                lambda x: x[:-2] if x.endswith('.0') else x
            ) == regal_odabrani_str
            
            filtrirano = self.df[strana_mask & regal_mask]

        col_sifra = 'Šifra' if 'Šifra' in self.df.columns else ('Sifra' if 'Sifra' in self.df.columns else self.df.columns[0])
        col_naziv = 'Naziv' if 'Naziv' in self.df.columns else self.df.columns[1]
        col_lager = 'Lager' if 'Lager' in self.df.columns else 'Kolicina'
        col_slika = 'Slika' if 'Slika' in self.df.columns else 'Putanja'
        col_polica = 'Polica' if 'Polica' in self.df.columns else None

        for _, row in filtrirano.iterrows():
            kod = str(row.get(col_sifra, 'N/A'))
            naziv = str(row.get(col_naziv, 'Nepoznato'))
            
            lager_val = row.get(col_lager, '0')
            try:
                lager_num = float(lager_val)
                lager_str = f"{int(lager_num)}" if lager_num.is_integer() else f"{lager_num}"
            except:
                lager_str = str(lager_val)
                
            polica_str = f" | Polica: {row.get(col_polica)}" if col_polica and pd.notna(row.get(col_polica)) else ""
            slika_path = str(row.get(col_slika, '')) if col_slika in row else ''

            odabrana_kol = self.kosarica.get(kod, {}).get('kolicina', 0)
            
            sec_text = f"Lager: {lager_str} kom{polica_str} | Naručeno: {odabrana_kol} kom"
            
            item = TwoLineAvatarIconListItem(
                text=f"{kod} - {naziv}",
                secondary_text=sec_text,
                on_release=lambda x, k=kod, n=naziv, l=lager_str, img=slika_path: self.dodaj_u_kosaricu_dialog(k, n, l, img)
            )
            icon = IconLeftWidget(icon="package-variant-closed")
            item.add_widget(icon)
            list_view.add_widget(item)
            
        scroll.add_widget(list_view)
        content.add_widget(scroll)

    def dodaj_u_kosaricu_dialog(self, kod, naziv, lager_str, slika_path):
        self.trenutni_kod = kod
        self.trenutni_naziv = naziv
        
        dialog_layout = MDBoxLayout(orientation='vertical', spacing="10dp", size_hint_y=None, height="240dp")
        
        lbl_lager = MDLabel(
            text=f"Predviđeni lager: [b]{lager_str} kom[/b]",
            markup=True,
            font_style="Subtitle1",
            size_hint_y=None,
            height="30dp"
        )
        dialog_layout.add_widget(lbl_lager)
        
        if slika_path and slika_path.lower() != 'nan' and os.path.exists(slika_path):
            img_widget = AsyncImage(source=slika_path, size_hint_y=None, height="120dp")
            dialog_layout.add_widget(img_widget)
        else:
            lbl_no_img = MDLabel(text="(Slika nije dostupna)", halign="center", size_hint_y=None, height="30dp")
            dialog_layout.add_widget(lbl_no_img)
            
        self.tf_kolicina = MDTextField(
            hint_text="Unesite broj ploča za narudžbu",
            input_filter="int",
            text="1"
        )
        dialog_layout.add_widget(self.tf_kolicina)
        
        self.dialog = MDDialog(
            title=f"{naziv}",
            type="custom",
            content_cls=dialog_layout,
            buttons=[
                MDRaisedButton(text="ODUSTANI", on_release=lambda x: self.dialog.dismiss()),
                MDRaisedButton(text="DODAJ U NALOG", on_release=self.potvrdi_unos)
            ]
        )
        self.dialog.open()

    def potvrdi_unos(self, *args):
        kol = self.tf_kolicina.text
        if kol.isdigit() and int(kol) > 0:
            self.kosarica[self.trenutni_kod] = {
                "naziv": self.trenutni_naziv,
                "kolicina": int(kol)
            }
        self.dialog.dismiss()
        self.prikazi_artikle()

    def otvori_kosaricu(self):
        content = self.root.ids.content_area
        content.clear_widgets()
        
        top_box = MDBoxLayout(orientation='horizontal', size_hint_y=None, height="40dp")
        lbl = MDLabel(text="Pregled naloga za odvajanje:", font_style="H6")
        btn_back = MDRaisedButton(text="NATRAG", on_release=lambda x: self.prikazi_odabir_strana())
        top_box.add_widget(lbl)
        top_box.add_widget(btn_back)
        content.add_widget(top_box)
        
        scroll = MDScrollView()
        list_view = MDList()
        
        for kod, data in self.kosarica.items():
            item = TwoLineAvatarIconListItem(
                text=f"{kod} - {data['naziv']}",
                secondary_text=f"Naručeno: {data['kolicina']} kom"
            )
            icon = IconLeftWidget(icon="check-bold")
            item.add_widget(icon)
            list_view.add_widget(item)
            
        scroll.add_widget(list_view)
        content.add_widget(scroll)

    def generiraj_pdf(self):
        if not self.kosarica:
            print("Košarica je prazna!")
            return
            
        pdf_filename = f"Nalog_Odvajanje_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        doc = SimpleDocTemplate(pdf_filename, pagesize=A4)
        elements = []
        
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'TitleStyle',
            parent=styles['Heading1'],
            fontSize=18,
            spaceAfter=12
        )
        
        elements.append(Paragraph("<b>NALOG ZA ODVAJANJE MATERIJALA</b>", title_style))
        elements.append(Paragraph(f"Datum i vrijeme: {datetime.now().strftime('%d.%m.%Y. %H:%M')}", styles['Normal']))
        elements.append(Spacer(1, 20))
        
        table_data = [["Šifra", "Naziv artikla", "Količina"]]
        for kod, data in self.kosarica.items():
            table_data.append([kod, data['naziv'], f"{data['kolicina']} kom"])
            
        t = Table(table_data, colWidths=[100, 280, 100])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#D35400")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#EAEDED")),
            ('GRID', (0, 0), (-1, -1), 1, colors.white)
        ]))
        
        elements.append(t)
        doc.build(elements)
        print(f"PDF uspješno kreiran pod nazivom: {pdf_filename}")

if __name__ == "__main__":
    SkladisteApp().run()