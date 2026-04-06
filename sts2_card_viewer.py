#!/usr/bin/env python3
import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
import os
import sys
from collections import defaultdict
import glob
import json

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

HISTORY_DIR = os.path.expanduser(
    "~/.local/share/SlayTheSpire2/steam/76561198054269638/profile1/saves/history/"
)
SPIRE_CODES_FILE = os.path.expanduser(
    "~/.local/share/opencode/tool-output/tool_d652ebd47001nzUP93jzebv47A"
)
EXCEL_FILE = os.path.join(BASE_DIR, "sts2_cards.xlsx")

CLASS_COLORS = {
    "IRONCLAD": "#FF8B8B",
    "SILENT": "#8BFF8B",
    "DEFECT": "#8B8BFF",
    "NECROBINDER": "#FF8BFF",
    "REGENT": "#FFFF8B",
    "COLORLESS": "#E0E0E0",
}


class STSCardViewer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("STS2 Card Statistics Viewer")
        self.geometry("1000x600")

        self.card_data = {}
        self.all_data = []
        self.current_class = None

        self.load_data()
        self.create_widgets()
        self.load_class("IRONCLAD")

    def load_data(self):
        if os.path.exists(EXCEL_FILE):
            wb = openpyxl.load_workbook(EXCEL_FILE)
            for class_name in wb.sheetnames:
                ws = wb[class_name]
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
        else:
            messagebox.showwarning(
                "Warning",
                f"Excel file not found: {EXCEL_FILE}\nPlease run sts2_stats.py first.",
            )

    def create_widgets(self):
        toolbar = ttk.Frame(self)
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

        main_frame = ttk.Frame(self)
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

    def refresh(self):
        self.load_data()
        if self.current_class:
            self.load_class(self.current_class)


if __name__ == "__main__":
    app = STSCardViewer()
    app.mainloop()
