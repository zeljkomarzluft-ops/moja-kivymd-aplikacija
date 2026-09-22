import os
import sys
from kivy.lang import Builder
from kivy.utils import platform
from kivymd.app import MDApp
from kivymd.uix.screen import MDScreen
from kivymd.uix.button import MDRaisedButton
from kivymd.uix.label import MDLabel
from kivymd.uix.boxlayout import MDBoxLayout

# 1. Definiranje sigurne putanje do Excel datoteke u sklopu APK-a
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'lager_pula.xlsx')  # 

class GlavniEkran(MDScreen):
    pass


class MojaKivyMDAplikacija(MDApp):

    def build(self):
        self.theme_cls.primary_palette = "Blue"
        self.theme_cls.theme_style = "Light"

        # Glavni raspored
        layout = MDBoxLayout(
            orientation='vertical',
            padding="20dp",
            spacing="20dp",
            pos_hint={'center_x': 0.5, 'center_y': 0.5}
        )

        # Naslov
        self.labela_status = MDLabel(
            text="Aplikacija je uspješno pokrenuta!",
            halign="center",
            font_style="H5"
        )
        layout.add_widget(self.labela_status)

        # Gumb za testiranje učitavanja Excela
        gumb_excel = MDRaisedButton(
            text="Učitaj Excel podatke",
            pos_hint={'center_x': 0.5},
            on_release=self.ucitaj_excel
        )
        layout.add_widget(gumb_excel)

        ekran = GlavniEkran()
        ekran.add_widget(layout)
        return ekran

    def on_start(self):
        # 2. Automatsko traženje dopuštenja na Androidu pri pokretanju
        if platform == 'android':
            try:
                from android.permissions import request_permissions, Permission
                request_permissions([
                    Permission.READ_EXTERNAL_STORAGE,
                    Permission.WRITE_EXTERNAL_STORAGE
                ])
            except Exception as e:
                print(f"Greška pri traženju dopuštenja: {e}")

    def ucitaj_excel(self, instance):
        # 3. Sigurno otvaranje Excel datoteke pomoću openpyxl
        try:
            import openpyxl

            if os.path.exists(EXCEL_PATH):
                wb = openpyxl.load_workbook(EXCEL_PATH)
                sheet = wb.active
                
                # Primjer čitanja prve ćelije (A1)
                vrijednost = sheet['A1'].value
                self.labela_status.text = f"Excel učitan!\nĆelija A1: {vrijednost}"
            else:
                self.labela_status.text = f"Datoteka nije pronađena na:\n{EXCEL_PATH}"

        except Exception as e:
            self.labela_status.text = f"Greška pri čitanju Excela:\n{str(e)}"


if __name__ == '__main__':
    MojaKivyMDAplikacija().run()
