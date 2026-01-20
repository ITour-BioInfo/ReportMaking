from __future__ import annotations

import datetime as dt
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from db import DatabaseError, load_database, validate_database
from report_generator import generate_report

DB_PATH = Path("data/variant_db.csv")
TEMPLATE_PATH = Path("templates/template.docx")
OUTPUT_DIR = Path("output")

PATIENT_FIELDS = [
    ("PATIENT_NAME", "Име и фамилия"),
    ("DOB", "Дата на раждане"),
    ("ID", "Идентификационен номер"),
    ("SAMPLE_DATE", "Материал получен"),
    ("ISSUE_DATE", "Дата на издаване"),
    ("ANALYST", "Извършил анализа"),
    ("LEAD", "Ръководител сектор"),
    ("DOCTOR", "Насочващ лекар"),
    ("PANEL_NAME", "Генетичен панел"),
]


class ScrollableFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        canvas = tk.Canvas(self, borderwidth=0, height=400)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")


class ReportApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Генетичен панел - Генератор на доклади")
        self.db = None
        self.variant_vars = {}
        self.genotype_vars = {}
        self.combo_widgets = {}

        self.patient_vars = {key: tk.StringVar() for key, _ in PATIENT_FIELDS}
        self.sex_var = tk.StringVar(value="F")

        self.preview_result = tk.StringVar()
        self.preview_comment = tk.StringVar()

        self._build_ui()
        self.reload_database()

    def _build_ui(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill="both", expand=True)

        patient_frame = ttk.LabelFrame(main_frame, text="Пациент")
        patient_frame.pack(fill="x")

        for idx, (key, label) in enumerate(PATIENT_FIELDS):
            ttk.Label(patient_frame, text=label).grid(row=idx, column=0, sticky="w", pady=2)
            entry = ttk.Entry(patient_frame, textvariable=self.patient_vars[key], width=40)
            entry.grid(row=idx, column=1, sticky="w", pady=2)

        sex_frame = ttk.Frame(patient_frame)
        sex_frame.grid(row=0, column=2, rowspan=3, sticky="nw", padx=20)
        ttk.Label(sex_frame, text="Пол").pack(anchor="w")
        ttk.Radiobutton(sex_frame, text="Ж", value="F", variable=self.sex_var, command=self._refresh_genotypes).pack(
            anchor="w"
        )
        ttk.Radiobutton(sex_frame, text="М", value="M", variable=self.sex_var, command=self._refresh_genotypes).pack(
            anchor="w"
        )

        variants_frame = ttk.LabelFrame(main_frame, text="Варианти")
        variants_frame.pack(fill="both", expand=True, pady=10)

        headers = ttk.Frame(variants_frame)
        headers.pack(fill="x")
        ttk.Label(headers, text="Полиморфизъм", width=20).grid(row=0, column=0)
        ttk.Label(headers, text="Ген", width=10).grid(row=0, column=1)
        ttk.Label(headers, text="Вариант", width=20).grid(row=0, column=2)
        ttk.Label(headers, text="Генотип", width=15).grid(row=0, column=3)

        self.scrollable = ScrollableFrame(variants_frame)
        self.scrollable.pack(fill="both", expand=True)

        preview_frame = ttk.LabelFrame(main_frame, text="Преглед")
        preview_frame.pack(fill="x", pady=10)
        ttk.Label(preview_frame, text="Резултат:").grid(row=0, column=0, sticky="w")
        ttk.Label(preview_frame, textvariable=self.preview_result).grid(row=0, column=1, sticky="w")
        ttk.Label(preview_frame, text="Коментар:").grid(row=1, column=0, sticky="nw")
        ttk.Label(preview_frame, textvariable=self.preview_comment, wraplength=600, justify="left").grid(
            row=1, column=1, sticky="w"
        )

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)
        ttk.Button(button_frame, text="Презареди база", command=self.reload_database).pack(side="left")
        ttk.Button(button_frame, text="Валидация", command=self.validate_db).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Генерирай DOCX", command=self.generate_docx).pack(side="left", padx=5)
        ttk.Button(
            button_frame,
            text="Генерирай DOCX + PDF",
            command=lambda: self.generate_docx(pdf=True),
        ).pack(side="left", padx=5)

    def reload_database(self):
        try:
            self.db = load_database(DB_PATH)
        except DatabaseError as exc:
            messagebox.showerror("Грешка", str(exc))
            return
        self._build_variant_rows()
        self._refresh_genotypes()

    def _build_variant_rows(self):
        for widget in self.scrollable.scrollable_frame.winfo_children():
            widget.destroy()
        self.variant_vars.clear()
        self.genotype_vars.clear()
        self.combo_widgets.clear()

        if not self.db:
            return

        for row_idx, variant in enumerate(self.db.variants):
            ttk.Label(self.scrollable.scrollable_frame, text=variant.polymorphism_display, width=20).grid(
                row=row_idx, column=0, sticky="w"
            )
            ttk.Label(self.scrollable.scrollable_frame, text=variant.gene_display, width=10).grid(
                row=row_idx, column=1, sticky="w"
            )
            ttk.Label(self.scrollable.scrollable_frame, text=variant.variant_display, width=20).grid(
                row=row_idx, column=2, sticky="w"
            )

            var = tk.StringVar()
            combo = ttk.Combobox(
                self.scrollable.scrollable_frame,
                textvariable=var,
                state="readonly",
                width=12,
            )
            combo.grid(row=row_idx, column=3, sticky="w")
            combo.bind("<<ComboboxSelected>>", lambda e, vid=variant.variant_id: self._update_preview(vid))

            self.variant_vars[variant.variant_id] = variant
            self.genotype_vars[variant.variant_id] = var
            self.combo_widgets[variant.variant_id] = combo

    def _refresh_genotypes(self):
        if not self.db:
            return
        sex = self.sex_var.get()
        for variant_id, combo in self.combo_widgets.items():
            variant = self.variant_vars[variant_id]
            effective_sex = sex if variant.chromosome == "X" else "Any"
            options = self.db.genotype_options(variant_id, effective_sex)
            combo.configure(values=options)
            current = self.genotype_vars[variant_id].get()
            if current not in options:
                self.genotype_vars[variant_id].set(options[0] if options else "")
        self.preview_result.set("")
        self.preview_comment.set("")

    def _update_preview(self, variant_id: str):
        if not self.db:
            return
        sex = self.sex_var.get()
        genotype = self.genotype_vars[variant_id].get()
        interpretation = self.db.lookup_interpretation(variant_id, sex, genotype)
        if interpretation is None:
            self.preview_result.set("")
            self.preview_comment.set("")
            return
        self.preview_result.set(interpretation.summary_result)
        self.preview_comment.set(interpretation.comment_text)

    def validate_db(self):
        if not self.db:
            return
        errors = validate_database(self.db)
        if errors:
            messagebox.showerror("Грешки", "\n".join(errors))
        else:
            messagebox.showinfo("OK", "Базата е валидна.")

    def generate_docx(self, pdf: bool = False):
        if not self.db:
            return
        OUTPUT_DIR.mkdir(exist_ok=True)
        patient_data = {key: var.get() for key, var in self.patient_vars.items()}
        patient_data["SEX"] = self.sex_var.get()
        if not patient_data.get("ISSUE_DATE"):
            patient_data["ISSUE_DATE"] = dt_today()

        selections = {variant_id: var.get() for variant_id, var in self.genotype_vars.items()}
        filename = _output_filename(patient_data)
        output_path = OUTPUT_DIR / filename

        try:
            generate_report(
                TEMPLATE_PATH,
                output_path,
                patient_data,
                selections,
                self.db,
                include_pdf=pdf,
            )
        except Exception as exc:
            messagebox.showerror("Грешка", str(exc))
            return

        messagebox.showinfo("Готово", f"Записан доклад: {output_path}")


def _output_filename(patient_data: dict) -> str:
    name = patient_data.get("PATIENT_NAME", "report").strip().replace(" ", "_")
    date = patient_data.get("ISSUE_DATE", dt_today()).replace("/", "-")
    return f"{name}_{date}.docx"


def dt_today() -> str:
    return dt.date.today().strftime("%d.%m.%Y")


def main() -> None:
    root = tk.Tk()
    ReportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
