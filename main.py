import flet as ft
import openpyxl
import os
import warnings
import time

warnings.filterwarnings("ignore")

class ToysArmyApp:
    def __init__(self):
        # --- 1. GESTIONE PERCORSI FILE ---
        self.exe_path = os.path.dirname(os.path.abspath(__file__))
        
        possible_paths = [
            os.path.join(self.exe_path, "assets", "Dati ToysArmy.xlsx"),
            os.path.join(self.exe_path, "Dati ToysArmy.xlsx"),
            "Dati ToysArmy.xlsx"
        ]
        
        self.file_path = "Dati ToysArmy.xlsx"
        for p in possible_paths:
            if os.path.exists(p):
                self.file_path = p
                break
        
        if not os.path.exists(self.file_path):
             if os.path.exists("assets/Dati ToysArmy.xlsx"):
                 self.file_path = "assets/Dati ToysArmy.xlsx"

        # --- 2. VARIABILI ---
        self.list_name = "Lista_senza_nome"
        self.selected_units = [] 
        self.compensations = [] 
        self.total_cost = 0.0
        self.alliance = ""
        self.faction = ""
        
        self.bg_image = "background.jpg" 

    def main(self, page: ft.Page):
        self.page = page
        self.page.title = "Toys Army"
        self.page.theme_mode = ft.ThemeMode.LIGHT
        self.page.padding = 0 
        
        self.show_home()

    # --- 3. HELPER EXCEL ---
    def get_excel_data(self, sheet_name):
        try:
            wb = openpyxl.load_workbook(self.file_path, data_only=True)
            if sheet_name not in wb.sheetnames:
                return []
            ws = wb[sheet_name]
            data = []
            for row in ws.iter_rows(values_only=True):
                data.append(list(row))
            
            if len(data) > 0:
                return data[1:] 
            return []
        except Exception as e:
            print(f"Errore lettura foglio {sheet_name}: {e}")
            return []

    # --- 4. BOTTONE INTELLIGENTE ---
    def get_smart_button(self, image_name, fallback_text, action, color="blue", size=30):
        image_full_path = os.path.join(self.exe_path, "assets", image_name)
        
        if os.path.exists(image_full_path):
            return ft.Container(
                content=ft.Image(src=image_name, width=size, height=size, fit="contain"),
                on_click=action,
                padding=5,
                border_radius=5,
                ink=True
            )
        else:
            return ft.ElevatedButton(text=fallback_text, on_click=action, bgcolor=color, color="white")

    # --- HELPER SFONDO ---
    def wrap_with_bg(self, content_column):
        return ft.Container(
            image=ft.DecorationImage(
                src=self.bg_image,
                fit="cover",
                opacity=0.9,
                alignment=ft.Alignment(0, 0)
            ),
            expand=True,
            padding=10,
            content=content_column,
            alignment=ft.Alignment(0, 0)
        )

    # --- 5. HOME ---
    def show_home(self):
        self.page.clean()
        
        main_content = ft.Column([
                ft.ElevatedButton("Asse", on_click=lambda _: self.select_alliance("Asse"), height=60, width=200),
                ft.ElevatedButton("Alleati", on_click=lambda _: self.select_alliance("Alleati"), height=60, width=200),
                ft.Container(height=20),
                
                ft.ElevatedButton("LISTE SALVATE", on_click=lambda _: self.show_saved_lists(), height=60, width=200, bgcolor="orange", color="white"),
                
            ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=20, expand=True)

        self.page.add(
            ft.AppBar(title=ft.Text("Toys Army Menu"), bgcolor="blue", color="white"),
            self.wrap_with_bg(main_content)
        )
        self.page.update()

    # --- VISUALIZZA LISTE SALVATE ---
    def show_saved_lists(self):
        self.page.clean()
        
        files = [f for f in os.listdir(self.exe_path) if f.endswith(".txt") and f != "requirements.txt"]
        
        list_column = ft.Column(spacing=10, horizontal_alignment=ft.CrossAxisAlignment.CENTER)
        
        if not files:
            list_column.controls.append(
                ft.Container(
                    content=ft.Text("Nessuna lista salvata trovata.", size=18, weight="bold"),
                    bgcolor="white", padding=10, border_radius=10
                )
            )
        else:
            for filename in files:
                safe_name = filename.replace(".txt", "")
                
                # NOME LISTA (Cliccabile per aprire la nuova schermata)
                btn_open = ft.Container(
                    content=ft.Text(safe_name, color="blue", size=18, weight="bold"),
                    on_click=lambda e, f=filename: self.open_list_screen(f),
                    padding=10,
                    ink=True, 
                    border_radius=5
                )

                row = ft.Row([
                    btn_open,
                    self.get_smart_button("Cancel.png", "X", lambda e, f=filename: self.delete_list(f), color="red", size=25)
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)
                
                container = ft.Container(
                    content=row,
                    padding=5,
                    bgcolor="white",
                    border_radius=10,
                    border=ft.border.all(1, "grey")
                )
                list_column.controls.append(container)

        self.page.add(
            ft.AppBar(
                title=ft.Text("Liste Salvate"), 
                bgcolor="blue", color="white",
                leading=self.get_smart_button("Back.png", "<", lambda _: self.show_home())
            ),
            self.wrap_with_bg(ft.Column([list_column], scroll="auto", expand=True, horizontal_alignment=ft.CrossAxisAlignment.CENTER))
        )
        self.page.update()

    # --- NUOVA SCHERMATA: LEGGI LISTA (Schermo Intero) ---
    def open_list_screen(self, filename):
        try:
            full_path = os.path.join(self.exe_path, filename)
            content = ""
            with open(full_path, "r", encoding="utf-8") as f:
                content = f.read()
            
            if not content: content = "Il file è vuoto."

            # Creiamo una vera schermata, non un popup
            self.page.clean()

            # Contenitore del testo (stile foglio di carta)
            text_container = ft.Container(
                content=ft.Text(content, size=16, color="black", weight="w500"),
                bgcolor="white",
                padding=15,
                border_radius=10,
                border=ft.border.all(1, "grey"),
                width=350 # Larghezza fissa per renderlo leggibile
            )

            # Contenuto principale
            main_content = ft.Column([
                text_container
            ], scroll="auto", alignment=ft.MainAxisAlignment.START, horizontal_alignment=ft.CrossAxisAlignment.CENTER, expand=True)

            self.page.add(
                ft.AppBar(
                    title=ft.Text(filename.replace(".txt", "")), 
                    bgcolor="blue", color="white",
                    # Il tasto Back riporta alle liste salvate
                    leading=self.get_smart_button("Back.png", "<", lambda _: self.show_saved_lists())
                ),
                self.wrap_with_bg(main_content)
            )
            self.page.update()
            
        except Exception as e:
            self.page.snack_bar = ft.SnackBar(ft.Text(f"Errore apertura: {str(e)}"), bgcolor="red")
            self.page.snack_bar.open = True
            self.page.update()

    def delete_list(self, filename):
        try:
            full_path = os.path.join(self.exe_path, filename)
            os.remove(full_path)
            self.page.snack_bar = ft.SnackBar(ft.Text(f"Lista {filename} eliminata!"), bgcolor="green")
            self.page.snack_bar.open = True
            self.show_saved_lists()
        except Exception as e:
            self.page.snack_bar = ft.SnackBar(ft.Text(f"Errore: {e}"), bgcolor="red")
            self.page.snack_bar.open = True
            self.page.update()

    def select_alliance(self, alliance):
        self.alliance = alliance
        self.show_faction()

    # --- 6. FAZIONE (CENTRATA) ---
    def show_faction(self):
        self.page.clean()
        factions = ["Unione Sovietica", "Stati Uniti d'America", "Regno Unito"] if self.alliance == "Alleati" else ["Terzo Reich", "Impero Giapponese", "Regno D'Italia"]
        
        buttons = []
        for f in factions:
            buttons.append(ft.ElevatedButton(f, on_click=lambda e, x=f: self.load_data(x), height=50, width=250))

        main_content = ft.Column(buttons, alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=20, expand=True)

        self.page.add(
            ft.AppBar(
                title=ft.Text(f"{self.alliance}"), 
                bgcolor="blue", color="white", 
                leading=self.get_smart_button("Back.png", "<", lambda _: self.show_home())
            ),
            self.wrap_with_bg(main_content)
        )
        self.page.update()

    def load_data(self, faction):
        self.faction = faction
        try:
            self.weapons_data = self.get_excel_data(f"Armi {faction}")
            self.roles_data = self.get_excel_data(f"Ruoli {faction}")
            self.tank_data = self.get_excel_data(f"Tank {faction}")
            self.mod_tank_data = self.get_excel_data(f"Mod Tank {faction}")
            self.show_units()
        except Exception as e:
            self.page.snack_bar = ft.SnackBar(ft.Text(f"Errore Dati: {str(e)}"), bgcolor="red")
            self.page.snack_bar.open = True
            self.page.update()

    # --- 7. UNITÀ ---
    def show_units(self):
        self.page.clean()
        container_list = []
        
        def make_block(label, data, cost_idx, desc_idxs):
            opts = []
            for row in data:
                if row[0] is not None:
                    name = str(row[0])
                    try: 
                        cost = row[cost_idx] if row[cost_idx] is not None else 0
                    except: cost = 0
                    opts.append(ft.dropdown.Option(text=f"{name} ({cost})", key=name))

            dd = ft.Dropdown(label="Seleziona", options=opts, expand=True, bgcolor="white")
            qty = ft.TextField(value="1", label="Qta", width=60, bgcolor="white")
            
            def add_btn_click(e):
                self.add_item_logic(dd.value, qty.value, data, cost_idx)
                dd.value = None
                self.page.snack_bar = ft.SnackBar(ft.Text("Aggiunto!"))
                self.page.snack_bar.open = True
                self.page.update()

            return ft.Container(
                content=ft.Column([
                    ft.Text(label, weight="bold"),
                    ft.Row([dd, qty]),
                    ft.ElevatedButton("AGGIUNGI", on_click=add_btn_click) 
                ]), padding=10, border=ft.border.all(1, "grey"), border_radius=10, bgcolor="white"
            )

        try:
            container_list.append(make_block("Armi", self.weapons_data, 3, [1,2]))
            container_list.append(make_block("Ruoli", self.roles_data, 1, [2]))
            container_list.append(make_block("Tank", self.tank_data, 8, [1]))
            container_list.append(make_block("Modifiche", self.mod_tank_data, 3, [1]))
        except Exception as e:
             self.page.add(ft.Text(f"Errore visualizzazione: {e}", color="red"))

        comp_val = ft.TextField(label="Valore da sottrarre", width=150, keyboard_type=ft.KeyboardType.NUMBER, bgcolor="white")
        
        def apply_comp_click(e):
            try:
                val = float(comp_val.value)
                self.compensations.append(val)
                self.total_cost -= val
                comp_val.value = ""
                self.page.snack_bar = ft.SnackBar(ft.Text(f"Sottratto: {val}"))
                self.page.snack_bar.open = True
                self.page.update()
            except:
                pass

        container_list.append(
            ft.Container(
                content=ft.Column([
                    ft.Text("Compensazione / Sconti", weight="bold", color="red"),
                    ft.Row([comp_val, ft.ElevatedButton("APPLICA", on_click=apply_comp_click, color="white", bgcolor="red")])
                ]), padding=10, border=ft.border.all(1, "red"), border_radius=10, margin=ft.margin.only(top=20), bgcolor="white"
            )
        )

        fab = ft.FloatingActionButton(
            content=ft.Text("LISTA", weight="bold"),
            width=80,
            on_click=lambda _: self.show_summary()
        )

        main_scroll_content = ft.Column(container_list, scroll="auto", expand=True)

        self.page.add(
            ft.AppBar(
                title=ft.Text(self.faction), 
                bgcolor="blue", color="white",
                leading=self.get_smart_button("Back.png", "<", lambda _: self.show_faction())
            ),
            self.wrap_with_bg(main_scroll_content),
            fab
        )
        self.page.update()

    def add_item_logic(self, name, qty_str, data, cost_idx):
        if not name: return
        try: q = int(qty_str)
        except: q = 1
        
        cost = 0
        for row in data:
            if str(row[0]) == name:
                try: 
                    cost = float(row[cost_idx])
                except: cost = 0
                break
        
        tot = cost * q
        self.selected_units.append({'id': time.time(), 'text': f"{name} x{q} ({tot})", 'cost': tot})
        self.total_cost += tot

    # --- 8. RIEPILOGO ---
    def show_summary(self):
        self.page.clean()
        items_col = ft.Column()

        def refresh_list_view():
            items_col.controls.clear()
            
            if not self.selected_units and not self.compensations:
                items_col.controls.append(ft.Text("Lista vuota"))
            
            for u in self.selected_units:
                row = ft.Row([
                    ft.Text(u['text'], expand=True),
                    self.get_smart_button("Cancel.png", "X", lambda e, x=u: remove_unit(x), color="red", size=25)
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)
                items_col.controls.append(row)

            if self.compensations:
                items_col.controls.append(ft.Divider())
                items_col.controls.append(ft.Text("Sconti:", italic=True, color="red"))
                for c in self.compensations:
                    row = ft.Row([
                        ft.Text(f"- {c}", color="red", expand=True),
                        self.get_smart_button("Cancel.png", "X", lambda e, x=c: remove_comp(x), color="red", size=25)
                    ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)
                    items_col.controls.append(row)
            
            total_label.value = f"TOTALE: {self.total_cost:.2f}"
            self.page.update()

        def remove_unit(unit_obj):
            self.selected_units.remove(unit_obj)
            self.total_cost -= unit_obj['cost']
            refresh_list_view()

        def remove_comp(comp_val):
            self.compensations.remove(comp_val)
            self.total_cost += comp_val
            refresh_list_view()

        total_label = ft.Text(f"TOTALE: {self.total_cost:.2f}", size=25, weight="bold", color="blue", bgcolor="white")
        refresh_list_view()

        def save_file(e):
            try:
                safe_name = "".join([c for c in self.list_name if c.isalnum() or c in (' ', '_')]).strip()
                if not safe_name: safe_name = "Lista"
                filename = f"{safe_name}.txt"
                full_path = os.path.join(self.exe_path, filename)
                
                with open(full_path, "w", encoding="utf-8") as f:
                    f.write(f"Nome Lista: {self.list_name}\n")
                    f.write(f"Totale Costo: {self.total_cost}\n\n")
                    f.write("--- ELENCO UNITÀ ---\n")
                    for u in self.selected_units: f.write(u['text'] + "\n")
                    if self.compensations:
                        f.write("\n--- SCONTI ---\n")
                        for c in self.compensations: f.write(f"- {c}\n")
                        
                self.page.snack_bar = ft.SnackBar(ft.Text(f"Salvato in: {filename}"), bgcolor="green", duration=4000)
                self.page.snack_bar.open = True
                self.page.update()
            except Exception as ex:
                self.page.snack_bar = ft.SnackBar(ft.Text(f"Errore: {ex}"), bgcolor="red")
                self.page.snack_bar.open = True
                self.page.update()

        name_field = ft.TextField(label="Nome Lista", value=self.list_name, on_change=lambda e: setattr(self, 'list_name', e.control.value), bgcolor="white")

        main_scroll_content = ft.Column([
                name_field,
                ft.Divider(),
                ft.Container(content=items_col, border=ft.border.all(1, "grey"), border_radius=5, padding=5, bgcolor="white"),
                ft.Divider(),
                total_label,
                ft.ElevatedButton("SALVA LISTA", on_click=save_file, bgcolor="green", color="white")
            ], scroll="auto", expand=True)

        self.page.add(
            ft.AppBar(
                title=ft.Text("Riepilogo"), 
                bgcolor="blue", color="white",
                leading=self.get_smart_button("Back.png", "<", lambda _: self.show_units())
            ),
            self.wrap_with_bg(main_scroll_content)
        )
        self.page.update()

if __name__ == "__main__":
    app = ToysArmyApp()
    assets_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
    if os.path.exists(assets_folder):
        ft.app(target=app.main, assets_dir="assets")
    else:
        ft.app(target=app.main)
