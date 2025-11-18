import toga
from toga.style import Pack
from toga.style.pack import COLUMN, ROW
import csv
from openpyxl import Workbook
import os  # <--- add this

class DummyFormApp(toga.App):

    def startup(self):
        self.main_window = toga.MainWindow(title=self.formal_name)

        # Full path to CSV file (relative to this file)
        base_dir = os.path.dirname(__file__)
        csv_path = os.path.join(base_dir, "options.csv")
        self.options = self.load_options(csv_path)  # now it will always find the file

        # Create dropdowns
        self.dropdowns = []
        for i, col_values in enumerate(self.options):
            dropdown = toga.Selection(
                items=col_values,
                style=Pack(flex=1)
            )
            self.dropdowns.append(dropdown)

        # Text input
        self.text_input = toga.TextInput(style=Pack(flex=1))

        # Buttons
        save_csv_btn = toga.Button("Save CSV", on_press=self.save_csv, style=Pack(padding=5))
        save_xlsx_btn = toga.Button("Save XLSX", on_press=self.save_xlsx, style=Pack(padding=5))

        # Layout
        # Layout
        box = toga.Box(style=Pack(direction=COLUMN, margin=10))

        for i, dd in enumerate(self.dropdowns):
            inner_box = toga.Box(
                children=[toga.Label(f"Dropdown {i+1}:", style=Pack(width=100, margin_right=5)), dd],
                style=Pack(direction=ROW, margin_bottom=5)
            )
            box.add(inner_box)

        text_box = toga.Box(
            children=[toga.Label("Text:", style=Pack(width=100, margin_right=5)), self.text_input],
            style=Pack(direction=ROW, margin_bottom=5)
        )
        box.add(text_box)

        button_box = toga.Box(
            children=[save_csv_btn, save_xlsx_btn],
            style=Pack(direction=ROW)
        )
        box.add(button_box)

        self.main_window.content = box

        self.main_window.show()

    def load_options(self, filename):
        """Load CSV columns as lists for dropdowns"""
        columns = []
        with open(filename, newline='') as f:
            reader = csv.reader(f)
            data = list(reader)
            if not data:
                return [[] for _ in range(4)]
            num_cols = len(data[0])
            for c in range(num_cols):
                col_values = [row[c] for row in data if row[c]]
                columns.append(col_values)
        return columns

    def save_csv(self, widget):
        row = [dd.value for dd in self.dropdowns] + [self.text_input.value]
        output_path = os.path.join(os.path.dirname(__file__), "output.csv")
        with open(output_path, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(row)
        self.main_window.info_dialog("Saved", f"Data saved to {output_path}")

    def save_xlsx(self, widget):
        wb = Workbook()
        ws = wb.active
        row = [dd.value for dd in self.dropdowns] + [self.text_input.value]
        ws.append(row)
        output_path = os.path.join(os.path.dirname(__file__), "output.xlsx")
        wb.save(output_path)
        self.main_window.info_dialog("Saved", f"Data saved to {output_path}")


def main():
    return DummyFormApp()
