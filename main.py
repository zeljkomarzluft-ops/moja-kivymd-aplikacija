import os
from kivy.lang import Builder
from kivy.utils import platform
from kivymd.app import MDApp
from kivymd.uix.screen import MDScreen
from kivymd.uix.button import MDRaisedButton
from kivymd.uix.label import MDLabel
from kivymd.uix.boxlayout import MDBoxLayout

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'lager_pula.xlsx')


class GlavniEkran(MDScreen):
    pass


class MojaKivyMDAplikacija(MDApp):

    def build(self):
        self.theme_cls.primary_palette = "Blue"
        self.theme_cls.theme_style = "Light"

        layout = MDBoxLayout(
            orientation='vertical',
            padding="20dp",
            spacing="20dp",
            pos_hint={'center_x': 0.5, 'center_y': 0.5}
        )

        self.labela_status = MDLabel(
            text="Aplikacija pokrenuta!\nPritisnite gumb za testiranje Excela.",
            halign="center",
            font_style="Body1"
        )
        layout.add_widget(self.labela_status)

        gumb_excel = MDRaisedButton(
            text="Učitaj Excel",
            pos_hint={'center_x': 0.5},
            on_release=self.ucitaj_excel
        )
        layout.add_widget(gumb_excel)

        ekran = GlavniEkran()
        ekran.add_widget(layout)
        return ekran

    def on_start(self):
        if platform == 'android':
            try:
                from android.permissions import request_permissions, Permission
                request_permissions([
                    Permission.READ_EXTERNAL_STORAGE,
                    Permission.WRITE_EXTERNAL_STORAGE
                ])
            except Exception as e:
                print(f"Dozvole greška: {e}")

    def ucitaj_excel(self, instance):
        try:
            import openpyxl

            if os.path.exists(EXCEL_PATH):
                wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True)
                sheet = wb.active
                vrijednost = sheet.cell(row=1, column=1).value
                self.labela_status.text = f"Excel uspješno pročitan!\nĆelija (1,1): {vrijednost}"
            else:
                self.labela_status.text = f"Datoteka NIJE pronađena na:\n{EXCEL_PATH}"

        except Exception as e:
            self.labela_status.text = f"Greška u radu s Excelom:\n{str(e)}"


if __name__ == '__main__':
    try:
        MojaKivyMDAplikacija().run()
    except Exception as err:
        print(f"Kritična greška aplikacije: {err}")
