#!/usr/bin/env python3
"""Popunjava missing vrednosti u koloni FORMULA za tabelu "Pol" bez eksternih zavisnosti.

Pravila:
- POL == 'M' -> FORMULA = 2 * $G$2
- POL == 'Z' -> FORMULA = 2 * $G$3

Napomena:
- Skripta radi nad .xlsx fajlom koristeci samo Python standardnu biblioteku
  (zipfile + xml.etree), pa je pogodna za offline/proxy ogranicena okruzenja.
"""

from __future__ import annotations

import argparse
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"a": MAIN_NS}


def _q(tag: str) -> str:
    return f"{{{MAIN_NS}}}{tag}"


def _load_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for si in root.findall("a:si", NS):
        text = "".join(node.text or "" for node in si.findall(".//a:t", NS))
        strings.append(text)
    return strings


def _get_cell(row: ET.Element, ref: str) -> ET.Element | None:
    for cell in row.findall("a:c", NS):
        if cell.attrib.get("r") == ref:
            return cell
    return None


def _cell_value(cell: ET.Element | None, shared_strings: list[str]) -> str | None:
    if cell is None:
        return None
    value_node = cell.find("a:v", NS)
    if value_node is None or value_node.text is None:
        return None

    raw = value_node.text
    if cell.attrib.get("t") == "s":
        return shared_strings[int(raw)]
    return raw


def _as_float(value: str | None, cell_ref: str) -> float:
    if value is None:
        raise ValueError(f"Nedostaje vrednost u celiji {cell_ref}.")
    try:
        return float(value)
    except ValueError as exc:
        raise ValueError(f"Vrednost u celiji {cell_ref} nije broj: {value!r}") from exc


def _set_numeric_cell(row: ET.Element, ref: str, value: float) -> None:
    cell = _get_cell(row, ref)
    if cell is None:
        cell = ET.SubElement(row, _q("c"), {"r": ref})

    # Ako je postojeci string tip, uklanjamo atribut da bi vrednost bila numericka.
    if "t" in cell.attrib:
        del cell.attrib["t"]

    formula_node = cell.find("a:f", NS)
    if formula_node is not None:
        cell.remove(formula_node)

    value_node = cell.find("a:v", NS)
    if value_node is None:
        value_node = ET.SubElement(cell, _q("v"))

    value_node.text = str(int(value) if value.is_integer() else value)


def fill_missing_values(input_path: Path, output_path: Path) -> int:
    with zipfile.ZipFile(input_path, "r") as archive:
        shared_strings = _load_shared_strings(archive)
        sheet_xml = archive.read("xl/worksheets/sheet1.xml")

    sheet_root = ET.fromstring(sheet_xml)

    rows_by_idx: dict[int, ET.Element] = {}
    for row in sheet_root.findall(".//a:sheetData/a:row", NS):
        row_index = int(row.attrib["r"])
        rows_by_idx[row_index] = row

    g2 = _as_float(_cell_value(_get_cell(rows_by_idx.get(2), "G2"), shared_strings), "G2")
    g3 = _as_float(_cell_value(_get_cell(rows_by_idx.get(3), "G3"), shared_strings), "G3")

    missing_count = 0
    for row_idx in range(2, 8):
        row = rows_by_idx.get(row_idx)
        if row is None:
            continue

        pol_value = _cell_value(_get_cell(row, f"B{row_idx}"), shared_strings)
        formula_cell = _get_cell(row, f"D{row_idx}")
        formula_value = _cell_value(formula_cell, shared_strings)

        if formula_value is not None:
            continue

        if pol_value == "M":
            _set_numeric_cell(row, f"D{row_idx}", 2 * g2)
            missing_count += 1
        elif pol_value == "Z":
            _set_numeric_cell(row, f"D{row_idx}", 2 * g3)
            missing_count += 1

    updated_sheet_xml = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)

    with zipfile.ZipFile(input_path, "r") as src, zipfile.ZipFile(output_path, "w") as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                data = updated_sheet_xml
            dst.writestr(info, data)

    return missing_count


def main() -> None:
    parser = argparse.ArgumentParser(description="Popuna missing FORMULA vrednosti u tabeli Pol.")
    parser.add_argument("--input", default="Codex_proba.xlsx", help="Ulazni Excel fajl")
    parser.add_argument("--output", default="Codex_proba_filled.xlsx", help="Izlazni Excel fajl")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)

    missing_count = fill_missing_values(input_path, output_path)
    print(f"Pronadjeno missing vrednosti u koloni FORMULA: {missing_count}")
    print(f"Sacuvan izlazni fajl: {output_path}")


if __name__ == "__main__":
    main()
