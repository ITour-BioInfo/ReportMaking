from __future__ import annotations

import copy
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Tuple
import xml.etree.ElementTree as ET

WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


@dataclass
class TemplateParts:
    summary_table: ET.Element
    summary_row: ET.Element
    section_header_table: ET.Element
    box_tables: Dict[int, ET.Element]
    templates_marker: ET.Element


def replace_placeholders_in_element(element: ET.Element, mapping: Dict[str, str]) -> None:
    for text in element.findall(".//w:t", WORD_NS):
        if text.text is None:
            continue
        for placeholder, value in mapping.items():
            if placeholder in text.text:
                text.text = text.text.replace(placeholder, value)


def find_table_by_placeholder(root: ET.Element, placeholder: str) -> Optional[ET.Element]:
    for table in root.findall(".//w:tbl", WORD_NS):
        if placeholder in _element_text(table):
            return table
    return None


def find_row_by_placeholder(table: ET.Element, placeholder: str) -> Optional[ET.Element]:
    for row in table.findall("w:tr", WORD_NS):
        if placeholder in _element_text(row):
            return row
    return None


def extract_templates(root: ET.Element) -> TemplateParts:
    marker = None
    for paragraph in root.findall(".//w:p", WORD_NS):
        if "TEMPLATES_DO_NOT_DELETE" in _element_text(paragraph):
            marker = paragraph
            break

    if marker is None:
        raise ValueError("Template marker TEMPLATES_DO_NOT_DELETE not found in template.docx")

    summary_table = find_table_by_placeholder(root, "[[POLYMORPHISM]]")
    if summary_table is None:
        raise ValueError("Summary table template not found (missing [[POLYMORPHISM]] placeholder)")

    summary_row = find_row_by_placeholder(summary_table, "[[POLYMORPHISM]]")
    if summary_row is None:
        raise ValueError("Summary row template not found in summary table")

    section_header_table = find_table_by_placeholder(root, "[[SECTION_TITLE]]")
    if section_header_table is None:
        raise ValueError("Section header template not found (missing [[SECTION_TITLE]])")

    box_tables = {}
    for count in (1, 2, 3):
        placeholder = f"[[RESULT_LEFT_{count}]]"
        table = find_table_by_placeholder(root, placeholder)
        if table is None:
            continue
        box_tables[count] = table

    return TemplateParts(
        summary_table=summary_table,
        summary_row=summary_row,
        section_header_table=section_header_table,
        box_tables=box_tables,
        templates_marker=marker,
    )


def clone_element(element: ET.Element) -> ET.Element:
    return copy.deepcopy(element)


def _element_text(element: ET.Element) -> str:
    texts = []
    for text in element.findall(".//w:t", WORD_NS):
        if text.text:
            texts.append(text.text)
    return "".join(texts)


def remove_element(parent: ET.Element, child: ET.Element) -> None:
    for idx, elem in enumerate(list(parent)):
        if elem is child:
            parent.remove(child)
            break


def insert_after(parent: ET.Element, reference: ET.Element, new_elements: Iterable[ET.Element]) -> None:
    elements = list(parent)
    try:
        index = elements.index(reference)
    except ValueError:
        parent.extend(new_elements)
        return
    insert_index = index + 1
    for elem in new_elements:
        parent.insert(insert_index, elem)
        insert_index += 1
