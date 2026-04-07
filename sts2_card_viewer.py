#!/usr/bin/env python3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
import os
import sys
import glob
import json
from collections import defaultdict
from data.card_classifications import CARD_CLASSIFICATIONS
from data.relic_info import RELIC_INFO

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_FILE = os.path.join(os.getcwd(), "sts2_cards.xlsx")

CLASS_COLORS = {
    "IRONCLAD": "FF8B8B",
    "SILENT": "8BFF8B",
    "DEFECT": "8B8BFF",
    "NECROBINDER": "FF8BFF",
    "REGENT": "FFFF8B",
}








def find_sts2_history_dir():
    if sys.platform == "win32":
        paths = [
            os.path.join(os.environ.get("APPDATA", ""), "SlayTheSpire2", "steam"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "SlayTheSpire2", "steam"),
        ]
    elif sys.platform == "darwin":
        paths = [
            os.path.expanduser("~/Library/Application Support/SlayTheSpire2/steam/")
        ]
    else:
        paths = [os.path.expanduser("~/.local/share/SlayTheSpire2/steam/")]

    for base_path in paths:
        if not os.path.exists(base_path):
            continue
        try:
            for entry in os.listdir(base_path):
                profile_base = os.path.join(base_path, entry)
                for subpath in ["profile1/saves/history", "saves/history"]:
                    profile_path = os.path.join(profile_base, subpath)
                    if os.path.isdir(profile_path) and glob.glob(
                        os.path.join(profile_path, "*.run")
                    ):
                        return profile_path
        except PermissionError:
            continue
    return None


def load_runs(history_dir):
    runs = []
    for f in sorted(glob.glob(os.path.join(history_dir, "*.run"))):
        try:
            with open(f, "r") as file:
                runs.append(json.load(file))
        except:
            continue
    return runs


def get_card_data(runs):
    pick_by_class = defaultdict(
        lambda: defaultdict(lambda: {"offered": 0, "picked": 0})
    )
    win_by_class = defaultdict(lambda: defaultdict(lambda: {"runs": 0, "wins": 0}))

    for run in runs:
        is_win = run.get("win", False)
        players = run.get("players", [])
        if not players:
            continue
        char = players[0].get("character", "Unknown").replace("CHARACTER.", "")
        if char not in CLASS_COLORS:
            continue

        cards_in_run = set()
        for act in run.get("map_point_history", []):
            for point in act:
                for stat in point.get("player_stats", []):
                    for choice in stat.get("card_choices", []):
                        card_id = (
                            choice.get("card", {})
                            .get("id", "Unknown")
                            .replace("CARD.", "")
                        )
                        was_picked = choice.get("was_picked", False)
                        pick_by_class[char][card_id]["offered"] += 1
                        if was_picked:
                            pick_by_class[char][card_id]["picked"] += 1
                            cards_in_run.add(card_id)

        for cid in cards_in_run:
            win_by_class[char][cid]["runs"] += 1
            if is_win:
                win_by_class[char][cid]["wins"] += 1

    return pick_by_class, win_by_class


def get_relic_data(runs):
    relic_stats = defaultdict(lambda: {"total": 0, "wins": 0})
    for run in runs:
        is_win = run.get("win", False)
        players = run.get("players", [])
        if not players:
            continue
        for relic in players[0].get("relics", []):
            rid = relic.get("id", "Unknown").replace("RELIC.", "")
            relic_stats[rid]["total"] += 1
            if is_win:
                relic_stats[rid]["wins"] += 1
    return relic_stats


def create_excel(pick_by_class, win_by_class, relic_stats, output_path):
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for class_name in CLASS_COLORS.keys():
        ws = wb.create_sheet(title=class_name)
        header_fill = PatternFill(
            start_color="FF" + CLASS_COLORS[class_name],
            end_color="FF" + CLASS_COLORS[class_name],
            fill_type="solid",
        )
        headers = [
            "Card",
            "Original Class",
            "Times Offered",
            "Times Picked",
            "Pick Rate %",
            "Runs with Card",
            "Wins with Card",
            "Win Rate %",
        ]

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = header_fill
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")

        row = 2
        class_cards = list(pick_by_class[class_name].items())
        class_cards.sort(key=lambda x: x[1]["picked"], reverse=True)

        for card_id, pick_stats in class_cards:
            win_stats = win_by_class[class_name].get(card_id, {"runs": 0, "wins": 0})
            original_class = CARD_CLASSIFICATIONS.get(card_id, "COLORLESS")
            pick_rate = (
                (pick_stats["picked"] / pick_stats["offered"] * 100)
                if pick_stats["offered"] > 0
                else 0
            )
            win_rate = (
                (win_stats["wins"] / win_stats["runs"] * 100)
                if win_stats["runs"] > 0
                else 0
            )

            ws.cell(row=row, column=1, value=card_id)
            ws.cell(row=row, column=2, value=original_class)
            ws.cell(row=row, column=3, value=pick_stats["offered"])
            ws.cell(row=row, column=4, value=pick_stats["picked"])
            ws.cell(row=row, column=5, value=round(pick_rate, 1))
            ws.cell(row=row, column=6, value=win_stats["runs"])
            ws.cell(row=row, column=7, value=win_stats["wins"])
            ws.cell(row=row, column=8, value=round(win_rate, 1))
            row += 1

        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 15
        for col in range(3, 9):
            ws.column_dimensions[get_column_letter(col)].width = 14

    # Add Relics sheet
    ws = wb.create_sheet(title="Relics")
    header_fill = PatternFill(
        start_color="CCCCCC", end_color="CCCCCC", fill_type="solid"
    )
    headers = ["ID", "Name", "Runs", "Wins", "Win%", "", "Description"]

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    row = 2
    relics_list = [(rid, stats) for rid, stats in relic_stats.items()]
    relics_list.sort(key=lambda x: x[1]["total"], reverse=True)

    for relic_id, stats in relics_list:
        win_rate = (stats["wins"] / stats["total"] * 100) if stats["total"] > 0 else 0
        relic_info = RELIC_INFO.get(
            relic_id, {"name": relic_id, "description": "Unknown"}
        )

        ws.cell(row=row, column=1, value=relic_id)
        ws.cell(row=row, column=2, value=relic_info.get("name", relic_id))
        ws.cell(row=row, column=3, value=stats["total"])
        ws.cell(row=row, column=4, value=stats["wins"])
        ws.cell(row=row, column=5, value=round(win_rate, 1))
        ws.cell(row=row, column=6, value="")
        ws.cell(row=row, column=7, value=relic_info.get("description", ""))
        row += 1

    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 5
    ws.column_dimensions["G"].width = 55

    wb.save(output_path)
    return sum(len(cards) for cards in pick_by_class.values())


class STSCardViewer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("STS2 Statistics Viewer")
        self.geometry("1200x700")
        self.history_dir = None
        self.card_data = {}
        self.relic_data = []
        self.all_data = []
        self.current_class = None

        self.create_widgets()
        self.find_and_load_data()

    def create_widgets(self):
        top_frame = ttk.Frame(self)
        top_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(top_frame, text="STS2 Save Location:").pack(side="left", padx=5)
        self.path_label = ttk.Label(top_frame, text="Not selected", foreground="gray")
        self.path_label.pack(side="left", padx=5)
        ttk.Button(top_frame, text="Browse...", command=self.browse_folder).pack(
            side="left", padx=5
        )
        ttk.Button(top_frame, text="Generate Data", command=self.generate_data).pack(
            side="left", padx=5
        )
        ttk.Button(top_frame, text="Help", command=self.show_help).pack(
            side="right", padx=5
        )

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)

        self.cards_frame = ttk.Frame(self.notebook)
        self.relics_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.cards_frame, text="Cards")
        self.notebook.add(self.relics_frame, text="Relics")

        self.create_cards_tab()
        self.create_relics_tab()

        self.status_label = ttk.Label(self, text="", relief="sunken", anchor="w")
        self.status_label.pack(fill="x", padx=5, pady=2)

    def create_cards_tab(self):
        toolbar = ttk.Frame(self.cards_frame)
        toolbar.pack(fill="x", padx=5, pady=5)

        ttk.Label(toolbar, text="Class:").pack(side="left", padx=5)
        self.class_var = tk.StringVar(value="IRONCLAD")
        for class_name in CLASS_COLORS.keys():
            rb = ttk.Radiobutton(
                toolbar,
                text=class_name,
                value=class_name,
                variable=self.class_var,
                command=self.on_class_change,
            )
            rb.pack(side="left", padx=2)

        ttk.Label(toolbar, text="Search:").pack(side="left", padx=(20, 5))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(toolbar, textvariable=self.search_var, width=20)
        self.search_entry.pack(side="left")
        self.search_entry.bind("<KeyRelease>", lambda e: self.filter_data())
        ttk.Button(toolbar, text="Refresh", command=self.refresh).pack(
            side="right", padx=5
        )

        main_frame = ttk.Frame(self.cards_frame)
        main_frame.pack(fill="both", expand=True, padx=5, pady=5)

        columns = ("Card", "Orig", "Offered", "Picked", "Pick%", "Runs", "Wins", "Win%")
        self.tree = ttk.Treeview(main_frame, columns=columns, show="headings")

        self.tree.heading("Card", text="Card", command=lambda: self.sort_by("card"))
        self.tree.heading(
            "Orig", text="Orig Class", command=lambda: self.sort_by("original_class")
        )
        self.tree.heading(
            "Offered", text="Times Offered", command=lambda: self.sort_by("offered")
        )
        self.tree.heading(
            "Picked", text="Times Picked", command=lambda: self.sort_by("picked")
        )
        self.tree.heading(
            "Pick%", text="Pick Rate %", command=lambda: self.sort_by("pick_rate")
        )
        self.tree.heading(
            "Runs", text="Runs with Card", command=lambda: self.sort_by("runs")
        )
        self.tree.heading(
            "Wins", text="Wins with Card", command=lambda: self.sort_by("wins")
        )
        self.tree.heading(
            "Win%", text="Win Rate %", command=lambda: self.sort_by("win_rate")
        )

        self.tree.column("Card", width=180)
        self.tree.column("Orig", width=80, anchor="center")
        for col in columns[2:]:
            self.tree.column(col, width=70, anchor="e")

        scrollbar = ttk.Scrollbar(
            main_frame, orient="vertical", command=self.tree.yview
        )
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.tree.tag_configure("oddrow", background="#f0f0f0")
        self.tree.tag_configure("evenrow", background="#ffffff")

    def create_relics_tab(self):
        toolbar = ttk.Frame(self.relics_frame)
        toolbar.pack(fill="x", padx=5, pady=5)

        ttk.Label(toolbar, text="Search:").pack(side="left", padx=5)
        self.relic_search_var = tk.StringVar()
        self.relic_search_entry = ttk.Entry(
            toolbar, textvariable=self.relic_search_var, width=20
        )
        self.relic_search_entry.pack(side="left")
        self.relic_search_entry.bind("<KeyRelease>", lambda e: self.filter_relics())
        ttk.Button(toolbar, text="Refresh", command=self.refresh).pack(
            side="right", padx=5
        )

        main_frame = ttk.Frame(self.relics_frame)
        main_frame.pack(fill="both", expand=True, padx=5, pady=5)

        columns = ("Name", "ID", "Runs", "Wins", "Win%", "Description")
        self.relic_tree = ttk.Treeview(
            main_frame, columns=columns, show="headings", selectmode="browse"
        )

        self.relic_tree.heading(
            "Name", text="Relic Name", command=lambda: self.sort_relics("name")
        )
        self.relic_tree.heading(
            "ID", text="Relic ID", command=lambda: self.sort_relics("id")
        )
        self.relic_tree.heading(
            "Runs", text="Runs", command=lambda: self.sort_relics("total")
        )
        self.relic_tree.heading(
            "Wins", text="Wins", command=lambda: self.sort_relics("wins")
        )
        self.relic_tree.heading(
            "Win%", text="Win Rate %", command=lambda: self.sort_relics("win_rate")
        )
        self.relic_tree.heading(
            "Description",
            text="Description",
            command=lambda: self.sort_relics("description"),
        )

        self.relic_tree.column("Name", width=150)
        self.relic_tree.column("ID", width=150)
        self.relic_tree.column("Runs", width=60, anchor="e")
        self.relic_tree.column("Wins", width=60, anchor="e")
        self.relic_tree.column("Win%", width=60, anchor="e")
        self.relic_tree.column("Description", width=500)

        scrollbar_y = ttk.Scrollbar(
            main_frame, orient="vertical", command=self.relic_tree.yview
        )
        scrollbar_x = ttk.Scrollbar(
            main_frame, orient="horizontal", command=self.relic_tree.xview
        )
        self.relic_tree.configure(
            yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set
        )

        self.relic_tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")

        self.relic_tree.tag_configure("oddrow", background="#f0f0f0")
        self.relic_tree.tag_configure("evenrow", background="#ffffff")

        self.relic_desc_label = ttk.Label(
            self.relics_frame, text="", wraplength=700, justify="left"
        )
        self.relic_desc_label.pack(fill="x", padx=5, pady=2)
        self.relic_tree.bind("<<TreeviewSelect>>", self.on_relic_select)

    def find_and_load_data(self):
        self.history_dir = find_sts2_history_dir()
        if self.history_dir:
            self.path_label.config(text=self.history_dir, foreground="black")
            self.generate_data()
        else:
            self.path_label.config(
                text="Not found - click Browse to select", foreground="red"
            )

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select STS2 save folder")
        if folder:
            history_path = os.path.join(folder, "saves", "history")
            if os.path.exists(history_path) and glob.glob(
                os.path.join(history_path, "*.run")
            ):
                self.history_dir = history_path
                self.path_label.config(text=history_path, foreground="black")
                self.generate_data()
            else:
                self.history_dir = folder
                self.path_label.config(text=folder, foreground="black")
                self.generate_data()

    def generate_data(self):
        if not self.history_dir:
            messagebox.showwarning("Warning", "Please select a STS2 save folder first.")
            return

        self.status_label.config(text="Loading runs...")
        self.update()

        runs = load_runs(self.history_dir)
        if not runs:
            messagebox.showwarning(
                "No Data", "No run files found in the selected folder."
            )
            self.status_label.config(text="No runs found")
            return

        self.status_label.config(text=f"Processing {len(runs)} runs...")
        self.update()

        pick_by_class, win_by_class = get_card_data(runs)
        relic_stats = get_relic_data(runs)
        total = create_excel(pick_by_class, win_by_class, relic_stats, EXCEL_FILE)

        self.status_label.config(
            text=f"Generated {total} cards and {len(relic_stats)} relics from {len(runs)} runs"
        )
        self.load_excel_data()

    def load_excel_data(self):
        if not os.path.exists(EXCEL_FILE):
            return

        self.card_data = {}
        self.relic_data = []
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            for class_name in wb.sheetnames:
                ws = wb[class_name]
                if class_name == "Relics":
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        if row[0]:
                            self.relic_data.append(
                                {
                                    "id": row[0],
                                    "name": row[1],
                                    "total": row[2],
                                    "wins": row[3],
                                    "win_rate": row[4],
                                    "description": row[6] if len(row) > 6 else "",
                                }
                            )
                else:
                    self.card_data[class_name] = []
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        if row[0]:
                            self.card_data[class_name].append(
                                {
                                    "card": row[0],
                                    "original_class": row[1],
                                    "offered": row[2],
                                    "picked": row[3],
                                    "pick_rate": row[4],
                                    "runs": row[5],
                                    "wins": row[6],
                                    "win_rate": row[7],
                                }
                            )
            self.load_class("IRONCLAD")
            self.load_relics()
            runs = load_runs(self.history_dir) if self.history_dir else []
            self.status_label.config(
                text=f"Loaded {sum(len(c) for c in self.card_data.values())} cards and {len(self.relic_data)} relics from {len(runs)} runs"
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {e}")

    def load_class(self, class_name):
        self.current_class = class_name
        self.all_data = self.card_data.get(class_name, [])
        self.filter_data()

    def filter_data(self):
        search = self.search_var.get().lower()
        filtered = self.all_data
        if search:
            filtered = [c for c in self.all_data if search in c["card"].lower()]

        for row in self.tree.get_children():
            self.tree.delete(row)

        for i, card in enumerate(filtered):
            tags = ("oddrow",) if i % 2 == 0 else ("evenrow",)
            self.tree.insert(
                "",
                "end",
                values=(
                    card["card"],
                    card["original_class"],
                    card["offered"],
                    card["picked"],
                    card["pick_rate"],
                    card["runs"],
                    card["wins"],
                    card["win_rate"],
                ),
                tags=tags,
            )

    def sort_by(self, key):
        reverse = getattr(self, "_sort_reverse", False)
        self._sort_reverse = not reverse

        def get_val(card):
            val = card.get(key)
            if isinstance(val, str):
                return val.lower()
            return val if val else 0

        self.all_data.sort(key=get_val, reverse=reverse)
        self.filter_data()

    def on_class_change(self):
        self.load_class(self.class_var.get())

    def load_relics(self):
        self.filter_relics()

    def filter_relics(self):
        search = self.relic_search_var.get().lower()
        filtered = self.relic_data
        if search:
            filtered = [
                r
                for r in self.relic_data
                if search in r["name"].lower() or search in r["id"].lower()
            ]

        for row in self.relic_tree.get_children():
            self.relic_tree.delete(row)

        for i, relic in enumerate(filtered):
            tags = ("oddrow",) if i % 2 == 0 else ("evenrow",)
            self.relic_tree.insert(
                "",
                "end",
                values=(
                    relic["name"],
                    relic["id"],
                    relic["total"],
                    relic["wins"],
                    relic["win_rate"],
                    relic["description"],
                ),
                tags=tags,
            )

    def sort_relics(self, key):
        reverse = getattr(self, "_relic_sort_reverse", False)
        self._relic_sort_reverse = not reverse

        def get_val(relic):
            val = relic.get(key)
            if isinstance(val, str):
                return val.lower()
            return val if val else 0

        self.relic_data.sort(key=get_val, reverse=reverse)
        self.filter_relics()

    def on_relic_select(self, event):
        selection = self.relic_tree.selection()
        if selection:
            item = self.relic_tree.item(selection[0])
            values = item["values"]
            if values:
                self.relic_desc_label.config(text=f"Description: {values[5]}")

    def refresh(self):
        self.find_and_load_data()

    def show_help(self):
        help_text = """How to find your STS2 save folder:

WINDOWS:
%LOCALAPPDATA%\\SlayTheSpire2\\steam\\<profile>\\saves\\history

MAC:
~/Library/Application Support/SlayTheSpire2/steam/<profile>/saves/history

LINUX:
~/.local/share/SlayTheSpire2/steam/<profile>/saves/history

NOTE: <profile> is usually a number like 76561198054269638

If auto-detection fails, click "Browse" and select the folder containing your .run files."""
        messagebox.showinfo("Help - Finding STS2 Save Files", help_text)


if __name__ == "__main__":
    app = STSCardViewer()
    app.mainloop()
