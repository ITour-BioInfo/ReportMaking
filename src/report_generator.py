from __future__ import annotations

import datetime as dt
import zipfile
from pathlib import Path
from typing import Dict, Iterable, List, Tuple
import xml.etree.ElementTree as ET

from db import Database
from template_tools import (
    TemplateParts,
    clone_element,
    extract_templates,
    insert_after,
    remove_element,
    replace_placeholders_in_element,
)

WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def generate_report(
    template_path: Path,
    output_path: Path,
    patient_data: Dict[str, str],
    selections: Dict[str, str],
    db: Database,
    include_pdf: bool = False,
) -> Path:
    xml_files = _load_docx_files(template_path)

    document_root = ET.fromstring(xml_files["word/document.xml"])
    template_parts = extract_templates(document_root)

    _replace_patient_placeholders(xml_files, patient_data)

    _build_summary_table(document_root, template_parts, selections, db, patient_data.get("SEX", "Any"))
    _build_sections(document_root, template_parts, selections, db, patient_data.get("SEX", "Any"))

    xml_files["word/document.xml"] = ET.tostring(document_root, encoding="utf-8", xml_declaration=True)
    _write_docx_files(output_path, xml_files)

    if include_pdf:
        _convert_to_pdf(output_path)

    return output_path


def _load_docx_files(path: Path) -> Dict[str, bytes]:
    with zipfile.ZipFile(path, "r") as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def _write_docx_files(path: Path, files: Dict[str, bytes]) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)


def _replace_patient_placeholders(xml_files: Dict[str, bytes], patient_data: Dict[str, str]) -> None:
    mapping = {f"[[{key}]]": value for key, value in patient_data.items()}
    for name in list(xml_files.keys()):
        if not name.startswith("word/"):
            continue
        if not (name.endswith("document.xml") or name.startswith("word/header") or name.startswith("word/footer")):
            continue
        root = ET.fromstring(xml_files[name])
        for element in root.findall(".//w:t", WORD_NS):
            if element.text is None:
                continue
            for placeholder, value in mapping.items():
                if placeholder in element.text:
                    element.text = element.text.replace(placeholder, value)
        xml_files[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _build_summary_table(
    root: ET.Element,
    templates: TemplateParts,
    selections: Dict[str, str],
    db: Database,
    sex: str,
) -> None:
    summary_table = templates.summary_table
    template_row = templates.summary_row

    parent = summary_table
    for variant in db.variants:
        row = clone_element(template_row)
        genotype = selections.get(variant.variant_id, "")
        interpretation = db.lookup_interpretation(variant.variant_id, sex, genotype)
        result = interpretation.summary_result if interpretation else ""
        replace_placeholders_in_element(
            row,
            {
                "[[POLYMORPHISM]]": variant.polymorphism_display,
                "[[GENE]]": variant.gene_display,
                "[[VARIANT]]": variant.variant_display,
                "[[GENOTYPE]]": genotype,
                "[[RESULT]]": result,
            },
        )
        parent.append(row)

    remove_element(summary_table, template_row)


def _build_sections(
    root: ET.Element,
    templates: TemplateParts,
    selections: Dict[str, str],
    db: Database,
    sex: str,
) -> None:
    body = root.find("w:body", WORD_NS)
    if body is None:
        return

    elements = list(body)
    summary_table = templates.summary_table
    marker = templates.templates_marker

    if summary_table not in elements or marker not in elements:
        return

    summary_index = elements.index(summary_table)
    marker_index = elements.index(marker)

    for element in elements[summary_index + 1 : marker_index]:
        body.remove(element)

    insert_elements = []
    grouped = _group_variants(db.variants)

    for category, blocks in grouped:
        section_table = clone_element(templates.section_header_table)
        replace_placeholders_in_element(section_table, {"[[SECTION_TITLE]]": category})
        insert_elements.append(section_table)

        for block_variants in blocks:
            _append_block_tables(insert_elements, block_variants, selections, db, templates, sex)

        conclusion_table = clone_element(templates.section_header_table)
        replace_placeholders_in_element(conclusion_table, {"[[SECTION_TITLE]]": "Извод:"})
        insert_elements.append(conclusion_table)

    insert_after(body, summary_table, insert_elements)


def _group_variants(variants) -> List[Tuple[str, List[List]]]:
    grouped = []
    current_category = None
    current_blocks = {}

    for variant in variants:
        if variant.category != current_category:
            if current_category is not None:
                grouped.append((current_category, list(current_blocks.values())))
            current_category = variant.category
            current_blocks = {}

        block = current_blocks.setdefault(variant.block_id, [])
        block.append(variant)

    if current_category is not None:
        grouped.append((current_category, list(current_blocks.values())))
    return grouped


def _append_block_tables(
    insert_elements: List[ET.Element],
    block_variants,
    selections: Dict[str, str],
    db: Database,
    templates: TemplateParts,
    sex: str,
) -> None:
    variants = list(block_variants)
    for chunk_start in range(0, len(variants), 3):
        chunk = variants[chunk_start : chunk_start + 3]
        template_table = templates.box_tables.get(len(chunk))
        if template_table is None:
            raise ValueError(f"Missing box template for {len(chunk)} variants")
        table = clone_element(template_table)
        mapping = {
            "[[BLOCK_TITLE]]": chunk[0].block_title,
            "[[INFO_TEXT]]": chunk[0].info_text,
        }
        for idx, variant in enumerate(chunk, start=1):
            genotype = selections.get(variant.variant_id, "")
            interpretation = db.lookup_interpretation(variant.variant_id, sex, genotype)
            mapping[f"[[RESULT_LEFT_{idx}]]"] = f"{variant.gene_display} {variant.variant_display}".strip()
            mapping[f"[[GENOTYPE_BIG_{idx}]]"] = genotype
            mapping[f"[[COMMENT_{idx}]]"] = interpretation.comment_text if interpretation else ""
        replace_placeholders_in_element(table, mapping)
        insert_elements.append(table)


def _convert_to_pdf(docx_path: Path) -> None:
    import subprocess

    output_dir = docx_path.parent
    subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--convert-to",
            "pdf",
            str(docx_path),
            "--outdir",
            str(output_dir),
        ],
        check=False,
    )
