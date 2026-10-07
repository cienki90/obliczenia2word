#!/usr/bin/env python3
"""
obliczenia2word – automatyczne wstawianie tabel z obliczeniami (Excel) do projektów (Word).

Dla każdego pliku .docx (np. 181.docx) program:
  1. wybiera arkusz o tej samej nazwie w pliku Excel (np. arkusz "181"),
  2. odtwarza zakres wydruku arkusza jako tabelę Word z zachowaniem formatowania
     (scalenia komórek, obramowania, czcionki, pogrubienia, wyrównanie, wysokości
     wierszy, szerokości kolumn, wypełnienia, format liczb),
  3. wstawia tabelę w rozdziale „Zestawienie zbiorcze słupów” – w tym samym miejscu,
     co w poprawnie wykonanych plikach (112.docx, 115.docx),
  4. zapisuje wynik (domyślnie do folderu "wynik").

Jeżeli dokument ma już wstawioną tabelę, zostanie ona podmieniona na aktualną.

Użycie:
    python obliczenia2word.py                          # wszystkie .docx z bieżącego folderu
    python obliczenia2word.py 181.docx 191.docx        # wybrane pliki
    python obliczenia2word.py -x "Obliczenia.xlsx" -o gotowe 181.docx
    python obliczenia2word.py --arkusz 181 projekt.docx
    python obliczenia2word.py --nadpisz 181.docx       # zapis w miejscu oryginału
"""
from __future__ import annotations

import argparse
import glob
import math
import os
import re
import sys
from xml.sax.saxutils import escape

try:
    import openpyxl
    from openpyxl.utils import range_boundaries
    import docx
    from docx.oxml import parse_xml
except ImportError:  # pragma: no cover
    print("Brak wymaganych bibliotek. Zainstaluj je poleceniem:\n"
          "    pip install python-docx openpyxl")
    sys.exit(1)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS

HEADING_TEXT = "Zestawienie zbiorcze słupów"
DEFAULT_EMPTY_PARAGRAPHS_BEFORE = 5   # tak jak w 112.docx i 115.docx
KEEP_WITH_NEXT = re.compile(r"obw nr|\bLp\b", re.IGNORECASE)

# Excel styl obramowania -> (styl Word, grubość w 1/8 pt)
BORDER_MAP = {
    "hair": ("single", 2),
    "thin": ("single", 4),
    "medium": ("single", 12),
    "thick": ("single", 18),
    "double": ("double", 4),
    "dotted": ("dotted", 4),
    "dashed": ("dashed", 4),
    "mediumDashed": ("dashed", 12),
    "dashDot": ("dotDash", 4),
    "mediumDashDot": ("dotDash", 12),
    "dashDotDot": ("dotDotDash", 4),
    "mediumDashDotDot": ("dotDotDash", 12),
    "slantDashDot": ("dashDotStroked", 12),
}

H_ALIGN_MAP = {
    "left": "left",
    "center": "center",
    "centerContinuous": "center",
    "right": "right",
    "justify": "both",
    "distributed": "distribute",
    "fill": "left",
}

V_ALIGN_MAP = {"top": "top", "center": "center", "bottom": "bottom",
               "justify": "center", "distributed": "center"}


# --------------------------------------------------------------------------- #
#  Excel -> model tabeli
# --------------------------------------------------------------------------- #
def _rgb(color) -> str | None:
    """Zwraca kolor w formacie RRGGBB (tylko kolory zdefiniowane wprost jako RGB)."""
    if color is None:
        return None
    try:
        if color.type != "rgb":
            return None
        rgb = color.rgb
    except Exception:
        return None
    if not isinstance(rgb, str) or len(rgb) < 6:
        return None
    return rgb[-6:].upper()


def format_value(value, number_format: str) -> str:
    """Formatuje wartość komórki tak, jak wyświetla ją Excel (dla typowych formatów)."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "PRAWDA" if value else "FAŁSZ"
    if isinstance(value, (int, float)):
        fmt = (number_format or "General").split(";")[0]
        if fmt == "@":
            return str(value)
        m = re.fullmatch(r'[#,]*0(?:\.(0+))?(%?)', fmt.replace("\\", "").replace('"', ""))
        if m:
            decimals = len(m.group(1) or "")
            v = value * 100 if m.group(2) else value
            thousands = "," in fmt
            s = f"{v:,.{decimals}f}" if thousands else f"{v:.{decimals}f}"
            if thousands:
                s = s.replace(",", "\u00a0")
            return s + m.group(2)
        # General
        if isinstance(value, float):
            if math.isfinite(value) and value == int(value) and abs(value) < 1e15:
                return str(int(value))
            return f"{value:.10g}"
        return str(value)
    if hasattr(value, "strftime"):
        return value.strftime("%d.%m.%Y")
    return str(value)


def column_widths_chars(ws, min_col: int, max_col: int) -> list[float]:
    default = ws.sheet_format.defaultColWidth or ((ws.sheet_format.baseColWidth or 8) + 0.71)
    widths = {c: default for c in range(min_col, max_col + 1)}
    hidden = set()
    for dim in ws.column_dimensions.values():
        lo, hi = dim.min or 0, dim.max or 0
        for c in range(lo, hi + 1):
            if c in widths:
                if dim.width:
                    widths[c] = dim.width
                if dim.hidden:
                    hidden.add(c)
    return [0.0 if c in hidden else widths[c] for c in range(min_col, max_col + 1)]


def read_sheet(ws, ws_values):
    """Czyta arkusz i zwraca prosty model: kolumny, wiersze, komórki z formatowaniem."""
    area = None
    if ws.print_area:
        pa = ws.print_area
        pa = pa[0] if isinstance(pa, (list, tuple)) else pa
        pa = pa.split(",")[0].split("!")[-1].replace("$", "")
        area = range_boundaries(pa)
    if not area:
        area = range_boundaries(ws.calculate_dimension())
    min_col, min_row, max_col, max_row = area

    col_chars = column_widths_chars(ws, min_col, max_col)
    default_h = ws.sheet_format.defaultRowHeight or 15

    merged_at = {}
    covered = set()
    for rng in ws.merged_cells.ranges:
        if rng.max_row < min_row or rng.min_row > max_row or rng.max_col < min_col or rng.min_col > max_col:
            continue
        merged_at[(rng.min_row, rng.min_col)] = rng
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                if (r, c) != (rng.min_row, rng.min_col):
                    covered.add((r, c))

    visible_cols = [c for i, c in enumerate(range(min_col, max_col + 1)) if col_chars[i] > 0]
    rows = []
    for r in range(min_row, max_row + 1):
        dim = ws.row_dimensions.get(r)
        if dim is not None and dim.hidden:
            continue
        height = (dim.height if dim is not None and dim.height else None) or default_h
        cells = []
        for c in visible_cols:
            if (r, c) in covered:
                rng = next((m for (mr, mc), m in merged_at.items()
                            if m.min_row < r <= m.max_row and m.min_col <= c <= m.max_col
                            and m.min_col == c), None)
                if rng is not None:  # kontynuacja scalenia pionowego
                    cells.append({"vmerge_continue": True, "col": c,
                                  "span": _visible_span(visible_cols, rng.min_col, rng.max_col),
                                  "borders": _merged_borders(ws, rng)})
                continue
            rng = merged_at.get((r, c))
            cell = ws.cell(r, c)
            value = ws_values.cell(r, c).value
            if isinstance(value, str) and value.startswith("=") and isinstance(cell.value, str):
                value = None  # formuła bez zapisanego wyniku
            if rng is not None:
                span = _visible_span(visible_cols, rng.min_col, rng.max_col)
                borders = _merged_borders(ws, rng)
                vmerge = rng.max_row > rng.min_row
            else:
                span, borders, vmerge = 1, _cell_borders(cell), False
            cells.append({
                "col": c,
                "span": span,
                "vmerge_start": vmerge,
                "borders": borders,
                "text": format_value(value, cell.number_format),
                "is_number": isinstance(value, (int, float)) and not isinstance(value, bool),
                "font": cell.font,
                "alignment": cell.alignment,
                "fill": cell.fill,
            })
        rows.append({"height": height, "cells": cells})
    widths = [col_chars[c - min_col] for c in visible_cols]
    return widths, rows, ws_values.title


def _visible_span(visible_cols, c1, c2):
    return max(1, sum(1 for c in visible_cols if c1 <= c <= c2))


def _side(border_side):
    if border_side is None or not border_side.style:
        return None
    style, size = BORDER_MAP.get(border_side.style, ("single", 4))
    return style, size, _rgb(border_side.color) or "auto"


def _cell_borders(cell):
    b = cell.border
    return {"top": _side(b.top), "left": _side(b.left),
            "bottom": _side(b.bottom), "right": _side(b.right)}


def _merged_borders(ws, rng):
    def first(sides):
        return next((s for s in sides if s), None)
    rows = range(rng.min_row, rng.max_row + 1)
    cols = range(rng.min_col, rng.max_col + 1)
    return {
        "top": first(_side(ws.cell(rng.min_row, c).border.top) for c in cols),
        "bottom": first(_side(ws.cell(rng.max_row, c).border.bottom) for c in cols),
        "left": first(_side(ws.cell(r, rng.min_col).border.left) for r in rows),
        "right": first(_side(ws.cell(r, rng.max_col).border.right) for r in rows),
    }


# --------------------------------------------------------------------------- #
#  Szerokości kolumn
# --------------------------------------------------------------------------- #
def _text_twips(text, font):
    """Przybliżona szerokość tekstu (twips) dla czcionki z Excela."""
    size_pt = float(font.sz or 11)
    factor = 0.60 if font.b else 0.55
    return len(text) * size_pt * factor * 20


def compute_grid(widths_chars, rows, available_twips):
    """Szerokości kolumn liczone jak autodopasowanie Worda: każda kolumna dostaje co
    najmniej szerokość najdłuższego słowa, a pozostałe miejsce jest dzielone
    proporcjonalnie do tego, ile kolumna potrzebuje, by tekst się nie zawijał.
    Kolumny bez tekstu dostają szerokość proporcjonalną do szerokości w Excelu."""
    n = len(widths_chars)
    cell_margin = 140 + 30  # marginesy komórki (2 x 70) + zapas
    minimum = [0.0] * n
    maximum = [0.0] * n
    for row in rows:
        idx = 0
        for cell in row["cells"]:
            span = cell["span"]
            text = cell.get("text")
            if span == 1 and text and idx < n:
                font = cell["font"]
                words = re.split(r"\s+", text)
                lines = text.split("\n")
                minimum[idx] = max(minimum[idx], max(_text_twips(w, font) for w in words) + cell_margin)
                maximum[idx] = max(maximum[idx], max(_text_twips(ln, font) for ln in lines) + cell_margin)
            idx += span

    px = [int(w * 7 + 5) if w > 0 else 0 for w in widths_chars]
    excel = [p * available_twips / (sum(px) or 1) for p in px]
    # kolumny bez własnego tekstu: minimum wg Excela (ale nie więcej niż potrzeba)
    for i in range(n):
        if minimum[i] == 0 and px[i] > 0:
            minimum[i] = maximum[i] = min(excel[i], 400)

    total_min = sum(minimum)
    if total_min >= available_twips:  # nie mieści się – skalujemy minimum
        grid = [m * available_twips / total_min for m in minimum]
    else:
        extra = available_twips - total_min
        want = [max(0.0, mx - mn) for mx, mn in zip(maximum, minimum)]
        total_want = sum(want)
        if total_want <= extra:  # wszystko mieści się w jednej linii
            spare = extra - total_want
            grid = [mx + spare * e / (sum(excel) or 1) for mx, e in zip(maximum, excel)]
        else:
            grid = [mn + extra * w / total_want for mn, w in zip(minimum, want)]
    grid = [int(round(g)) for g in grid]
    grid[-1] += available_twips - sum(grid)
    return grid


# --------------------------------------------------------------------------- #
#  Model -> XML tabeli Word
# --------------------------------------------------------------------------- #
def _border_xml(name, side):
    if side is None:
        return f'<w:{name} w:val="nil"/>'
    style, size, color = side
    return f'<w:{name} w:val="{style}" w:sz="{size}" w:space="0" w:color="{color}"/>'


def _rpr_xml(font):
    name = escape(font.name or "Calibri", {'"': "&quot;"})
    parts = [f'<w:rFonts w:ascii="{name}" w:eastAsia="Times New Roman" w:hAnsi="{name}" w:cs="{name}"/>']
    if font.b:
        parts.append("<w:b/><w:bCs/>")
    if font.i:
        parts.append("<w:i/><w:iCs/>")
    if font.strike:
        parts.append("<w:strike/>")
    color = _rgb(font.color)
    if color and color != "000000":
        parts.append(f'<w:color w:val="{color}"/>')
    if font.sz:
        hp = int(round(float(font.sz) * 2))
        parts.append(f'<w:sz w:val="{hp}"/><w:szCs w:val="{hp}"/>')
    if font.u:
        parts.append('<w:u w:val="{}"/>'.format("double" if "double" in str(font.u) else "single"))
    if font.vertAlign in ("superscript", "subscript"):
        parts.append(f'<w:vertAlign w:val="{font.vertAlign}"/>')
    parts.append('<w:lang w:val="pl-PL"/>')
    return "<w:rPr>" + "".join(parts) + "</w:rPr>"


def _runs_xml(text, rpr):
    out = []
    for i, line in enumerate(text.replace("\r\n", "\n").split("\n")):
        if i:
            out.append(f"<w:r>{rpr}<w:br/></w:r>")
        if line:
            out.append(f'<w:r>{rpr}<w:t xml:space="preserve">{escape(line)}</w:t></w:r>')
    return "".join(out)


def build_table_xml(grid, rows):
    xml = [f'<w:tbl xmlns:w="{W_NS}">',
           "<w:tblPr>",
           f'<w:tblW w:w="{sum(grid)}" w:type="dxa"/>',
           '<w:jc w:val="center"/>',
           '<w:tblLayout w:type="autofit"/>',
           '<w:tblCellMar><w:left w:w="70" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tblCellMar>',
           '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" '
           'w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>',
           "</w:tblPr>",
           "<w:tblGrid>" + "".join(f'<w:gridCol w:w="{g}"/>' for g in grid) + "</w:tblGrid>"]

    for row in rows:
        height = int(round(row["height"] * 20))
        row_text = " ".join(c.get("text", "") for c in row["cells"])
        # nagłówek obwodu i nagłówek kolumn trzymamy razem z kolejnym wierszem,
        # żeby nie zostały same na dole strony
        keep_next = "<w:keepNext/>" if KEEP_WITH_NEXT.search(row_text) else ""
        xml.append(f'<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="{height}"/></w:trPr>')
        idx = 0
        for cell in row["cells"]:
            span = cell["span"]
            width = sum(grid[idx:idx + span])
            idx += span
            b = cell["borders"]
            tcpr = [f'<w:tcW w:w="{width}" w:type="dxa"/>']
            if span > 1:
                tcpr.append(f'<w:gridSpan w:val="{span}"/>')
            if cell.get("vmerge_continue"):
                tcpr.append("<w:vMerge/>")
            elif cell.get("vmerge_start"):
                tcpr.append('<w:vMerge w:val="restart"/>')
            tcpr.append("<w:tcBorders>" + "".join(_border_xml(n, b[n]) for n in
                                                  ("top", "left", "bottom", "right")) + "</w:tcBorders>")
            if cell.get("vmerge_continue"):
                xml.append("<w:tc><w:tcPr>" + "".join(tcpr) + "</w:tcPr><w:p/></w:tc>")
                continue

            fill = cell["fill"]
            if fill is not None and fill.fill_type == "solid":
                color = _rgb(fill.fgColor)
                if color:
                    tcpr.append(f'<w:shd w:val="clear" w:color="auto" w:fill="{color}"/>')
            al = cell["alignment"]
            tcpr.append(f'<w:vAlign w:val="{V_ALIGN_MAP.get(al.vertical or "bottom", "bottom")}"/>')

            if al.horizontal in H_ALIGN_MAP:
                jc = H_ALIGN_MAP[al.horizontal]
            else:  # "general": liczby do prawej, tekst do lewej
                jc = "right" if cell["is_number"] else "left"
            rpr = _rpr_xml(cell["font"])
            ppr = f'<w:pPr>{keep_next}<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>' \
                  f'<w:jc w:val="{jc}"/>{rpr}</w:pPr>'
            xml.append("<w:tc><w:tcPr>" + "".join(tcpr) + "</w:tcPr><w:p>" + ppr
                       + _runs_xml(cell["text"], rpr) + "</w:p></w:tc>")
        xml.append("</w:tr>")
    xml.append("</w:tbl>")
    return "".join(xml)


# --------------------------------------------------------------------------- #
#  Wstawianie do dokumentu
# --------------------------------------------------------------------------- #
def _text(el):
    return "".join(t.text or "" for t in el.iter(W + "t"))


def _is_empty_paragraph(el):
    return (el.tag == W + "p" and not _text(el).strip()
            and el.find(".//" + W + "drawing") is None
            and el.find(".//" + W + "pict") is None
            and el.find(".//" + W + "sectPr") is None)


def find_insertion_point(document, empty_before):
    body = document.element.body
    children = list(body.iterchildren())
    heading_idx = None
    for i, el in enumerate(children):
        if el.tag == W + "p" and _text(el).strip().rstrip(":").lower() == HEADING_TEXT.lower():
            heading_idx = i
    if heading_idx is None:
        raise ValueError(f'nie znaleziono nagłówka "{HEADING_TEXT}"')

    # istniejąca tabela (dokument już wypełniony) -> do podmiany
    anchor = children[heading_idx]
    existing = None
    empties = 0
    for el in children[heading_idx + 1:]:
        if el.tag == W + "tbl":
            existing = el
            break
        if not _is_empty_paragraph(el):
            break
        empties += 1
        if empties <= empty_before:
            anchor = el
    return anchor, existing


def section_text_width(document, anchor):
    """Szerokość obszaru tekstu (twips) sekcji, w której znajdzie się tabela."""
    sect = None
    el = anchor.getnext()
    while el is not None:
        sect = el.find(".//" + W + "sectPr") if el.tag == W + "p" else (el if el.tag == W + "sectPr" else None)
        if sect is not None:
            break
        el = el.getnext()
    if sect is None:
        sect = document.element.body.find(W + "sectPr")
    pg, mar = sect.find(W + "pgSz"), sect.find(W + "pgMar")
    width = int(pg.get(W + "w")) - int(mar.get(W + "left")) - int(mar.get(W + "right"))
    if mar.get(W + "gutter"):
        width -= int(mar.get(W + "gutter"))
    return width


def sheet_for_document(wb, docx_path, sheet_name=None):
    if sheet_name:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f'w pliku Excel nie ma arkusza "{sheet_name}"')
        return sheet_name
    stem = os.path.splitext(os.path.basename(docx_path))[0]
    if stem in wb.sheetnames:
        return stem
    numbers = re.findall(r"\d+", stem)
    for num in numbers:
        for name in wb.sheetnames:
            if re.fullmatch(r"\d+", name.strip()) and int(name) == int(num):
                return name
    raise ValueError(f'nie znaleziono arkusza pasującego do pliku "{os.path.basename(docx_path)}"')


def process(docx_path, wb, wb_values, out_path, sheet_name=None, empty_before=DEFAULT_EMPTY_PARAGRAPHS_BEFORE):
    name = sheet_for_document(wb, docx_path, sheet_name)
    document = docx.Document(docx_path)
    anchor, existing = find_insertion_point(document, empty_before)

    widths, rows, _ = read_sheet(wb[name], wb_values[name])
    grid = compute_grid(widths, rows, section_text_width(document, anchor))
    table = parse_xml(build_table_xml(grid, rows))

    if existing is not None:
        existing.addprevious(table)
        existing.getparent().remove(existing)
        action = "podmieniono tabelę"
    else:
        anchor.addnext(table)
        action = "wstawiono tabelę"
    document.save(out_path)
    return name, action, len(rows)


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Wstawia tabele obliczeń z pliku Excel do dokumentów Word "
                    "(rozdział „Zestawienie zbiorcze słupów”).")
    ap.add_argument("docx", nargs="*", help="pliki .docx lub foldery (domyślnie: bieżący folder)")
    ap.add_argument("-x", "--excel", help="plik Excel z obliczeniami (domyślnie: jedyny .xlsx w folderze)")
    ap.add_argument("-o", "--wynik", default="wynik", help='folder wyjściowy (domyślnie "wynik")')
    ap.add_argument("--nadpisz", action="store_true", help="zapisz zmiany w oryginalnych plikach")
    ap.add_argument("--arkusz", help="nazwa arkusza (gdy nazwa pliku nie odpowiada nazwie arkusza)")
    ap.add_argument("--puste-akapity", type=int, default=DEFAULT_EMPTY_PARAGRAPHS_BEFORE,
                    help="ile pustych akapitów pod nagłówkiem zostawić przed tabelą (domyślnie 5)")
    args = ap.parse_args(argv)

    base = os.path.dirname(os.path.abspath(__file__))
    inputs = args.docx or [os.getcwd()]
    files = []
    for p in inputs:
        if os.path.isdir(p):
            files += sorted(glob.glob(os.path.join(p, "*.docx")))
        else:
            files.append(p)
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    if not files:
        print("Nie znaleziono plików .docx.")
        return 1

    excel = args.excel
    if not excel:
        search_dirs = {os.path.dirname(os.path.abspath(files[0])), os.getcwd(), base}
        found = sorted({os.path.abspath(f) for d in search_dirs for f in glob.glob(os.path.join(d, "*.xlsx"))
                        if not os.path.basename(f).startswith("~$")})
        if len(found) != 1:
            print("Podaj plik Excel opcją -x (znaleziono: %s)." % (", ".join(found) or "brak"))
            return 1
        excel = found[0]
    print(f"Excel: {excel}")
    wb = openpyxl.load_workbook(excel)
    wb_values = openpyxl.load_workbook(excel, data_only=True)

    if not args.nadpisz:
        os.makedirs(args.wynik, exist_ok=True)
    ok = errors = 0
    for f in files:
        out = f if args.nadpisz else os.path.join(args.wynik, os.path.basename(f))
        try:
            sheet, action, n = process(f, wb, wb_values, out, args.arkusz, args.puste_akapity)
            print(f"  OK   {os.path.basename(f)}: {action} z arkusza \"{sheet}\" ({n} wierszy) -> {out}")
            ok += 1
        except Exception as e:  # noqa: BLE001
            print(f"  BŁĄD {os.path.basename(f)}: {e}")
            errors += 1
    print(f"Gotowe: {ok} plików, błędów: {errors}.")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
