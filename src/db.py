from __future__ import annotations

import csv
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
import xml.etree.ElementTree as ET

SHEET_NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


@dataclass(frozen=True)
class Variant:
    variant_id: str
    polymorphism_display: str
    gene_display: str
    variant_display: str
    category: str
    chromosome: str
    block_id: str
    block_title: str
    info_text: str
    display_order: int


@dataclass(frozen=True)
class Interpretation:
    variant_id: str
    sex: str
    genotype: str
    summary_result: str
    comment_text: str


@dataclass
class Database:
    variants: List[Variant]
    interpretations: List[Interpretation]
    interpretation_lookup: Dict[Tuple[str, str, str], Interpretation]
    variant_lookup: Dict[str, Variant]

    def genotype_options(self, variant_id: str, sex: str) -> List[str]:
        options = []
        for interpretation in self.interpretations:
            if interpretation.variant_id != variant_id:
                continue
            if interpretation.sex not in ("Any", sex):
                continue
            options.append(interpretation.genotype)
        return sorted(set(options))

    def lookup_interpretation(self, variant_id: str, sex: str, genotype: str) -> Optional[Interpretation]:
        return self.interpretation_lookup.get((variant_id, sex, genotype)) or self.interpretation_lookup.get(
            (variant_id, "Any", genotype)
        )


class DatabaseError(Exception):
    pass


def load_database(path: Path) -> Database:
    if path.suffix.lower() == ".csv":
        variants, interpretations = _load_from_csv(path)
    elif path.suffix.lower() == ".xlsx":
        variants, interpretations = _load_from_xlsx(path)
    else:
        raise DatabaseError(f"Unsupported database format: {path}")

    interpretation_lookup = {}
    for entry in interpretations:
        key = (entry.variant_id, entry.sex, entry.genotype)
        interpretation_lookup[key] = entry

    variant_lookup = {variant.variant_id: variant for variant in variants}
    return Database(
        variants=variants,
        interpretations=interpretations,
        interpretation_lookup=interpretation_lookup,
        variant_lookup=variant_lookup,
    )


def validate_database(db: Database) -> List[str]:
    errors = []
    required_variant_fields = [
        "variant_id",
        "category",
        "block_id",
        "block_title",
        "info_text",
        "display_order",
    ]

    for variant in db.variants:
        for field in required_variant_fields:
            if not getattr(variant, field):
                errors.append(f"Missing {field} for variant {variant.variant_id}")

    for variant in db.variants:
        options = [
            interp
            for interp in db.interpretations
            if interp.variant_id == variant.variant_id
        ]
        if not options:
            errors.append(f"No interpretations for variant {variant.variant_id}")

    seen = set()
    for interp in db.interpretations:
        key = (interp.variant_id, interp.sex, interp.genotype)
        if key in seen:
            errors.append(
                f"Duplicate interpretation for {interp.variant_id} {interp.sex} {interp.genotype}"
            )
        seen.add(key)

    return errors


def _load_from_csv(path: Path) -> Tuple[List[Variant], List[Interpretation]]:
    if not path.exists():
        raise DatabaseError(f"Database file not found: {path}")

    variant_rows = []
    interpretation_rows = []

    with path.open("r", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            if row.get("sheet") == "VariantCatalog":
                variant_rows.append(row)
            elif row.get("sheet") == "Interpretations":
                interpretation_rows.append(row)

    if not variant_rows or not interpretation_rows:
        raise DatabaseError(
            "CSV database must include rows tagged with sheet=VariantCatalog and sheet=Interpretations"
        )

    return _rows_to_records(variant_rows, interpretation_rows)


def _load_from_xlsx(path: Path) -> Tuple[List[Variant], List[Interpretation]]:
    if not path.exists():
        raise DatabaseError(f"Database file not found: {path}")

    with zipfile.ZipFile(path, "r") as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        sheet_map = {}
        for sheet in workbook.findall("s:sheets/s:sheet", SHEET_NS):
            name = sheet.attrib.get("name")
            rel_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            sheet_map[rel_id] = name

        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_ns = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
        rel_targets = {
            rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall("r:Relationship", rel_ns)
        }

        shared_strings = _load_shared_strings(zf)

        variant_rows = []
        interpretation_rows = []
        for rel_id, sheet_name in sheet_map.items():
            target = rel_targets.get(rel_id)
            if not target:
                continue
            sheet_path = f"xl/{target}"
            rows = _read_sheet(zf, sheet_path, shared_strings)
            if not rows:
                continue
            headers = rows[0]
            data_rows = rows[1:]
            records = [dict(zip(headers, row)) for row in data_rows]
            if sheet_name == "VariantCatalog":
                variant_rows = records
            elif sheet_name == "Interpretations":
                interpretation_rows = records

    if not variant_rows or not interpretation_rows:
        raise DatabaseError("VariantCatalog or Interpretations sheet missing in Excel database")

    return _rows_to_records(variant_rows, interpretation_rows)


def _rows_to_records(
    variant_rows: Iterable[Dict[str, str]],
    interpretation_rows: Iterable[Dict[str, str]],
) -> Tuple[List[Variant], List[Interpretation]]:
    variants = []
    for row in variant_rows:
        if not row.get("variant_id"):
            continue
        display_order = _parse_display_order(row)
        variants.append(
            Variant(
                variant_id=row.get("variant_id", "").strip(),
                polymorphism_display=row.get("polymorphism_display", "").strip(),
                gene_display=row.get("gene_display", "").strip(),
                variant_display=row.get("variant_display", "").strip(),
                category=row.get("category", "").strip(),
                chromosome=row.get("chromosome", "").strip(),
                block_id=row.get("block_id", "").strip(),
                block_title=row.get("block_title", "").strip(),
                info_text=row.get("info_text", "").strip(),
                display_order=display_order,
            )
        )

    interpretations = []
    for row in interpretation_rows:
        if not row.get("variant_id"):
            continue
        interpretations.append(
            Interpretation(
                variant_id=row.get("variant_id", "").strip(),
                sex=row.get("sex", "Any").strip() or "Any",
                genotype=row.get("genotype", "").strip(),
                summary_result=row.get("summary_result", "").strip(),
                comment_text=row.get("comment_text", "").strip(),
            )
        )

    variants.sort(key=lambda v: v.display_order)
    return variants, interpretations


def _parse_display_order(row: Dict[str, str]) -> int:
    value = (row.get("display_order") or "").strip()
    if not value:
        return 0
    try:
        return int(value)
    except ValueError as exc:
        variant_id = row.get("variant_id", "<unknown>")
        raise DatabaseError(
            f"Invalid display_order '{value}' for variant {variant_id}. "
            "Ensure the CSV uses quotes around text fields that contain commas."
        ) from exc


def _load_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []

    root = ET.fromstring(data)
    values = []
    for si in root.findall("s:si", SHEET_NS):
        text_parts = []
        for t in si.findall(".//s:t", SHEET_NS):
            text_parts.append(t.text or "")
        values.append("".join(text_parts))
    return values


def _read_sheet(zf: zipfile.ZipFile, sheet_path: str, shared_strings: List[str]) -> List[List[str]]:
    data = zf.read(sheet_path)
    root = ET.fromstring(data)
    rows = []

    for row in root.findall("s:sheetData/s:row", SHEET_NS):
        row_values = []
        max_col = 0
        cells = row.findall("s:c", SHEET_NS)
        cell_map = {}
        for cell in cells:
            cell_ref = cell.attrib.get("r", "")
            col_letters = re.sub(r"\d", "", cell_ref)
            col_index = _col_letters_to_index(col_letters)
            max_col = max(max_col, col_index)
            value = _cell_value(cell, shared_strings)
            cell_map[col_index] = value
        for col_idx in range(1, max_col + 1):
            row_values.append(cell_map.get(col_idx, ""))
        rows.append(row_values)

    return rows


def _cell_value(cell: ET.Element, shared_strings: List[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        text_el = cell.find("s:is/s:t", SHEET_NS)
        return text_el.text if text_el is not None else ""
    if cell_type == "s":
        v = cell.find("s:v", SHEET_NS)
        if v is None or v.text is None:
            return ""
        idx = int(v.text)
        return shared_strings[idx] if idx < len(shared_strings) else ""
    v = cell.find("s:v", SHEET_NS)
    return v.text if v is not None and v.text is not None else ""


def _col_letters_to_index(letters: str) -> int:
    result = 0
    for char in letters:
        result = result * 26 + (ord(char.upper()) - 64)
    return result
