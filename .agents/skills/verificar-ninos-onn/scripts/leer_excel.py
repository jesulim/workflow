#!/usr/bin/env python3
"""Lector de listas de niños desde archivos Excel (.xlsx) sin dependencias externas."""

from __future__ import annotations

import argparse
import json
import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

NS = {"s": MAIN_NS}


def _norm_header(text: str) -> str:
    """Normaliza encabezados para detección tolerante a tildes y puntuación."""
    nfd = unicodedata.normalize("NFD", text or "")
    sin_tildes = "".join(ch for ch in nfd if unicodedata.category(ch) != "Mn")
    limpio = re.sub(r"[^a-z0-9]+", " ", sin_tildes.lower()).strip()
    return limpio


def _clean_cell_text(val: str) -> str:
    """Limpia espacios y normaliza números enteros guardados como float en Excel."""
    s = re.sub(r"\s+", " ", (val or "").strip())
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".", 1)[0]
    return s


def _col_from_ref(ref: str) -> str:
    """Extrae la letra de columna de una referencia Excel (ej. 'G53' -> 'G')."""
    m = re.match(r"([A-Z]+)", (ref or "").upper())
    return m.group(1) if m else ""


def _resolve_sheet_path(zf: zipfile.ZipFile, sheet_name: str | None = None) -> tuple[str, str]:
    """Resuelve la ruta XML interna de la hoja solicitada o de la primera hoja."""
    wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
    sheets = wb_root.findall(".//s:sheet", NS)
    if not sheets:
        raise ValueError("El archivo Excel no contiene hojas.")

    chosen = None
    if sheet_name:
        target_norm = _norm_header(sheet_name)
        for sh in sheets:
            name = sh.attrib.get("name", "")
            if name == sheet_name or _norm_header(name) == target_norm:
                chosen = sh
                break
        if chosen is None:
            disponibles = [sh.attrib.get("name", "") for sh in sheets]
            raise ValueError(
                f"No se encontró la hoja '{sheet_name}'. Hojas disponibles: {disponibles}"
            )
    else:
        chosen = sheets[0]

    chosen_name = chosen.attrib.get("name", "Hoja1")
    rid = chosen.attrib.get(f"{{{REL_NS}}}id")

    rels_path = "xl/_rels/workbook.xml.rels"
    if rid and rels_path in zf.namelist():
        rels_root = ET.fromstring(zf.read(rels_path))
        for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship"):
            if rel.attrib.get("Id") == rid:
                target = rel.attrib.get("Target", "").lstrip("/")
                if not target.startswith("xl/"):
                    target = f"xl/{target}"
                if target in zf.namelist():
                    return chosen_name, target

    fallback = "xl/worksheets/sheet1.xml"
    if fallback in zf.namelist():
        return chosen_name, fallback
    raise ValueError("No se pudo localizar el XML de la hoja en el archivo .xlsx.")


def _load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    """Carga la tabla sharedStrings.xml del .xlsx si existe."""
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for si in root.findall("s:si", NS):
        texts = [t.text or "" for t in si.findall(".//s:t", NS)]
        strings.append("".join(texts))
    return strings


def _extract_rows_by_col(
    sheet_xml: bytes, shared_strings: list[str]
) -> list[tuple[int, dict[str, str]]]:
    """Devuelve lista de (fila_excel, {col_letra: valor_texto}) preservando filas dispersas."""
    root = ET.fromstring(sheet_xml)
    rows_out: list[tuple[int, dict[str, str]]] = []
    for idx, row_el in enumerate(root.findall(".//s:sheetData/s:row", NS), start=1):
        row_num_str = row_el.attrib.get("r", "")
        row_num = int(row_num_str) if row_num_str.isdigit() else idx
        cells: dict[str, str] = {}
        for cell_el in row_el.findall("s:c", NS):
            ref = cell_el.attrib.get("r", "")
            col = _col_from_ref(ref)
            if not col:
                continue
            cell_type = cell_el.attrib.get("t", "")
            v_el = cell_el.find("s:v", NS)
            is_el = cell_el.find("s:is", NS)
            val = ""
            if cell_type == "s" and v_el is not None and v_el.text is not None:
                ss_idx = int(v_el.text)
                if 0 <= ss_idx < len(shared_strings):
                    val = shared_strings[ss_idx]
            elif cell_type == "inlineStr" and is_el is not None:
                val = "".join(t.text or "" for t in is_el.findall(".//s:t", NS))
            elif cell_type in ("str", "b", "e") and v_el is not None and v_el.text is not None:
                val = v_el.text
            elif v_el is not None and v_el.text is not None:
                val = v_el.text
            elif is_el is not None:
                val = "".join(t.text or "" for t in is_el.findall(".//s:t", NS))

            cleaned = _clean_cell_text(val)
            if cleaned != "":
                cells[col] = cleaned
        rows_out.append((row_num, cells))
    return rows_out


def _detect_header_mapping(
    rows: list[tuple[int, dict[str, str]]]
) -> tuple[int, dict[str, str]]:
    """Detecta la fila de encabezados (en filas 1..15) y mapea campos -> letra de columna."""
    for idx, (row_num, cells) in enumerate(rows):
        if row_num > 15:
            break
        col_map: dict[str, str] = {}
        for col, raw_text in cells.items():
            h = _norm_header(raw_text)
            if not h:
                continue
            if h in ("n", "no", "nro", "num", "numero") or h.startswith("n "):
                col_map.setdefault("numero", col)
            elif any(k in h for k in ("nombre", "nino", "persona", "apellido")):
                col_map.setdefault("nombre", col)
            elif any(k in h for k in ("carnet", "c i", "identidad", "cedula", "dni")) or (
                "ci" in h.split()
            ):
                col_map.setdefault("carnet", col)
            elif "edad" in h:
                col_map.setdefault("edad", col)
            elif "estado" in h or "invitacion" in h:
                col_map.setdefault("estado_invitacion", col)
            elif "visto" in h or h == "v b" or "vb" in h.split():
                col_map.setdefault("visto_bueno", col)
            elif "sistema" in h:
                col_map.setdefault("en_sistema", col)
            elif "observa" in h or "nota" in h or "detalle" in h:
                col_map.setdefault("observaciones", col)

        if "nombre" in col_map and (
            "carnet" in col_map or len(col_map) >= 2
        ):
            return idx, col_map

    raise ValueError(
        "No se pudo detectar la fila de encabezados con columna de nombre en las primeras 15 filas."
    )


def leer_personas_excel(
    xlsx_path: str | Path, sheet_name: str | None = None
) -> dict[str, Any]:
    """Lee un archivo .xlsx y devuelve metadatos y la lista estructurada de personas."""
    path = Path(xlsx_path)
    if not path.is_file():
        raise FileNotFoundError(f"No existe el archivo Excel: {path}")

    with zipfile.ZipFile(path, "r") as zf:
        resolved_sheet, sheet_xml_path = _resolve_sheet_path(zf, sheet_name)
        shared_strings = _load_shared_strings(zf)
        rows = _extract_rows_by_col(zf.read(sheet_xml_path), shared_strings)

    header_idx, col_map = _detect_header_mapping(rows)
    header_row_num = rows[header_idx][0]

    personas: list[dict[str, Any]] = []
    for row_num, cells in rows[header_idx + 1 :]:
        nombre = cells.get(col_map["nombre"], "").strip()
        carnet_raw = cells.get(col_map.get("carnet", ""), "").strip()
        if not nombre and not carnet_raw:
            continue

        carnet_digits = re.sub(r"\D+", "", carnet_raw)
        personas.append(
            {
                "fila_excel": row_num,
                "numero": cells.get(col_map.get("numero", ""), "").strip(),
                "nombre": nombre,
                "carnet_raw": carnet_raw,
                "carnet_digits": carnet_digits,
                "edad_excel": cells.get(col_map.get("edad", ""), "").strip(),
                "estado_invitacion": cells.get(
                    col_map.get("estado_invitacion", ""), ""
                ).strip(),
                "visto_bueno": cells.get(col_map.get("visto_bueno", ""), "").strip(),
                "en_sistema_excel": cells.get(
                    col_map.get("en_sistema", ""), ""
                ).strip(),
                "observaciones_excel": cells.get(
                    col_map.get("observaciones", ""), ""
                ).strip(),
            }
        )

    return {
        "archivo": str(path.resolve()),
        "nombre_archivo": path.name,
        "hoja": resolved_sheet,
        "fila_encabezado": header_row_num,
        "columnas_detectadas": col_map,
        "total_personas": len(personas),
        "personas": personas,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Lee un archivo .xlsx con lista de niños y muestra los registros detectados."
    )
    parser.add_argument("excel", help="Ruta al archivo .xlsx (ej. yaque.xlsx)")
    parser.add_argument("--sheet", default=None, help="Nombre opcional de la hoja Excel")
    parser.add_argument(
        "--json",
        action="store_true",
        help="Imprime toda la salida en formato JSON",
    )
    args = parser.parse_args()

    data = leer_personas_excel(args.excel, sheet_name=args.sheet)
    if args.json:
        print(json.dumps(data, ensure_ascii=False, indent=2))
        return 0

    print(
        f"Archivo: {data['nombre_archivo']} | Hoja: {data['hoja']} | "
        f"Encabezado fila {data['fila_encabezado']} | Total personas: {data['total_personas']}"
    )
    print(f"Columnas: {data['columnas_detectadas']}")
    for p in data["personas"]:
        print(
            f"  Fila {p['fila_excel']:>2} | N° {p['numero'] or '-':>2} | "
            f"{p['nombre']} | CI: {p['carnet_raw'] or '(sin CI)'} "
            f"[{p['carnet_digits'] or '-'}] | Excel EnSistema: {p['en_sistema_excel'] or '-'}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
