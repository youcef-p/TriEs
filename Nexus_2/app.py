from flask import Flask, render_template, request, jsonify, send_file, Response
import os
import re
import io
import tempfile
import string
import platform
import time
import json
import copy
import queue
import threading
import shutil
import uuid
from pathlib import Path
from datetime import date, datetime
from typing import Any, Optional
import unicodedata

# Safely import external libraries
try: import pdfplumber
except ImportError: pdfplumber = None

try:
    import xlrd
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill
except ImportError:
    xlrd = None
    load_workbook = None

app = Flask(__name__)

# ==========================================
# APP 1: DRIVE LISTER LOGIC
# ==========================================
def format_size(size_bytes):
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if size_bytes < 1024: return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} PB"

def list_contents(path, show_hidden, show_sizes, recursive):
    lines = []
    path = Path(path).expanduser()
    if not path.exists() or not path.is_dir(): return [f"❌ '{path}' isn't a valid, accessible folder."]
    lines.append(f"\n📂 Listing contents of: {path}\n")
    lines.append("-" * 60)
    entry_count = 0; total_size = 0
    walker = os.walk(str(path)) if recursive else [(str(path), [], os.listdir(str(path)))]
    for root, dirs, files in walker:
        if not recursive:
            all_entries = os.listdir(str(path))
            dirs = [e for e in all_entries if (path / e).is_dir()]
            files = [e for e in all_entries if (path / e).is_file()]
        if not show_hidden:
            dirs[:] = [d for d in dirs if not d.startswith(".")]
            files = [f for f in files if not f.startswith(".")]
        rel_root = os.path.relpath(root, str(path))
        indent = "" if rel_root == "." else "  " * rel_root.count(os.sep)
        if rel_root != ".": lines.append(f"{indent}📁 {os.path.basename(root)}/")
        for d in sorted(dirs):
            if recursive: continue
            lines.append(f"{indent}📁 {d}/"); entry_count += 1
        for f in sorted(files):
            full_path = os.path.join(root, f)
            try: size = os.path.getsize(full_path)
            except OSError: size = 0
            total_size += size; entry_count += 1
            if show_sizes: lines.append(f"{indent}  📄 {f}  ({format_size(size)})")
            else: lines.append(f"{indent}  📄 {f}")
        if not recursive: break
    lines.append("-" * 60)
    lines.append(f"\n✅ Done! Found {entry_count} item(s).")
    if show_sizes: lines.append(f"📦 Total size of listed files: {format_size(total_size)}")
    return lines

def build_file_index(path, show_hidden):
    file_index = {}; dir_set = set()
    for root, dirs, files in os.walk(str(path)):
        if not show_hidden:
            dirs[:] = [d for d in dirs if not d.startswith(".")]
            files = [f for f in files if not f.startswith(".")]
        rel_root = os.path.relpath(root, str(path))
        for d in dirs: dir_set.add((d if rel_root == "." else os.path.join(rel_root, d)).replace(os.sep, "/"))
        for f in files:
            rel_file = (f if rel_root == "." else os.path.join(rel_root, f)).replace(os.sep, "/")
            try: size = os.path.getsize(os.path.join(root, f))
            except OSError: size = -1
            file_index[rel_file] = size
    return file_index, dir_set

def compare_listings(path_a, path_b, show_hidden):
    lines = []
    files_a, dirs_a = build_file_index(path_a, show_hidden)
    files_b, dirs_b = build_file_index(path_b, show_hidden)
    lines.append("             FOLDER / DRIVE COMPARISON RESULT\n" + "=" * 60 + f"\nA: {path_a}\nB: {path_b}\n" + "-" * 60)
    files_a_set, files_b_set = set(files_a.keys()), set(files_b.keys())
    only_in_a, only_in_b = sorted(files_a_set - files_b_set), sorted(files_b_set - files_a_set)
    common = sorted(files_a_set & files_b_set)
    same_size = [f for f in common if files_a[f] == files_b[f]]
    diff_size = [f for f in common if files_a[f] != files_b[f]]
    only_dirs_a, only_dirs_b = sorted(dirs_a - dirs_b), sorted(dirs_b - dirs_a)
    lines.append(f"\n📊 SUMMARY\n  Files only in A:            {len(only_in_a)}\n  Files only in B:            {len(only_in_b)}\n  Files in both, same size:   {len(same_size)}\n  Files in both, diff. size:  {len(diff_size)}\n  Folders only in A:          {len(only_dirs_a)}\n  Folders only in B:          {len(only_dirs_b)}")
    
    if only_dirs_a:
        lines.append(f"\n📁 Folders only in A ({len(only_dirs_a)}):")
        for d in only_dirs_a: lines.append(f"  + {d}/")
    if only_dirs_b:
        lines.append(f"\n📁 Folders only in B ({len(only_dirs_b)}):")
        for d in only_dirs_b: lines.append(f"  + {d}/")
    if only_in_a:
        lines.append(f"\n📄 Files only in A ({len(only_in_a)}):")
        for f in only_in_a: lines.append(f"  + {f}")
    if only_in_b:
        lines.append(f"\n📄 Files only in B ({len(only_in_b)}):")
        for f in only_in_b: lines.append(f"  + {f}")
    if diff_size:
        lines.append(f"\n⚠️  Different sizes ({len(diff_size)}):")
        for f in diff_size: lines.append(f"  ≠ {f}   A: {format_size(files_a[f])}   |   B: {format_size(files_b[f])}")
        
    lines.append(f"\n✅ Matches: {len(same_size)}\n" + "-" * 60)
    total_diffs = len(only_in_a) + len(only_in_b) + len(diff_size) + len(only_dirs_a) + len(only_dirs_b)
    lines.append("🎉 Perfect match." if total_diffs == 0 else f"🔎 Total differences: {total_diffs}")
    return lines

_HEADER_RE = re.compile(r"^📂 Listing contents of:\s*(.+)$")
_FOLDER_RE = re.compile(r"^(\s*)📁\s+(.+?)/\s*$")
_FILE_RE = re.compile(r"^(\s*)📄\s+(.+?)(?:\s{2}\(([^)]+)\))?\s*$")

def parse_listing_file(filepath):
    with open(filepath, "r", encoding="utf-8", errors="replace") as fh: raw_lines = fh.read().splitlines()
    source_label, files, dirs, stack, had_sizes = None, {}, set(), [], False
    for line in raw_lines:
        if source_label is None:
            m = _HEADER_RE.match(line.strip())
            if m: source_label = m.group(1).strip(); continue
        fm = _FOLDER_RE.match(line)
        if fm: depth = len(fm.group(1)) // 2; stack = stack[:depth] + [fm.group(2)]; dirs.add("/".join(stack)); continue
        pm = _FILE_RE.match(line)
        if pm:
            parent = "/".join(stack[:len(pm.group(1)) // 2])
            rel_path = f"{parent}/{pm.group(2)}" if parent else pm.group(2)
            if pm.group(3): had_sizes = True
            files[rel_path] = pm.group(3)
    return {"source_label": source_label or Path(filepath).name, "files": files, "dirs": dirs, "had_sizes": had_sizes}

def compare_parsed_listings(parsed_a, parsed_b):
    lines = [f"     COMPARISON OF TWO PREVIOUSLY EXPORTED LISTINGS\n{'=' * 60}\nA: {parsed_a['source_label']}\nB: {parsed_b['source_label']}\n{'-' * 60}"]
    both_sizes = parsed_a['had_sizes'] and parsed_b['had_sizes']
    files_a_set, files_b_set = set(parsed_a['files']), set(parsed_b['files'])
    only_a, only_b = sorted(files_a_set - files_b_set), sorted(files_b_set - files_a_set)
    common = sorted(files_a_set & files_b_set)
    same_size = [f for f in common if not both_sizes or parsed_a['files'][f] == parsed_b['files'][f]]
    diff_size = [f for f in common if both_sizes and parsed_a['files'][f] != parsed_b['files'][f]]
    lines.append(f"\n📊 SUMMARY\n  Files only in A: {len(only_a)}\n  Files only in B: {len(only_b)}")
    if only_a:
        lines.append(f"\n📄 Files only in A:")
        for f in only_a: lines.append(f"  + {f}")
    if only_b:
        lines.append(f"\n📄 Files only in B:")
        for f in only_b: lines.append(f"  + {f}")
    if diff_size:
        lines.append(f"\n⚠️  Different sizes:")
        for f in diff_size: lines.append(f"  ≠ {f}   A: {parsed_a['files'][f]}   |   B: {parsed_b['files'][f]}")
    lines.append(f"\n✅ Matches: {len(same_size)}\n{'-' * 60}")
    total_diffs = len(only_a) + len(only_b) + len(diff_size)
    lines.append("🎉 Perfect match." if total_diffs == 0 else f"🔎 Total differences: {total_diffs}")
    return lines

# ==========================================
# APP 3: PDF COMPRESSOR LOGIC
# ==========================================
STOPWORDS = {"the", "a", "an", "is", "are", "was", "were", "be", "been", "being", "have", "has", "had", "do", "does", "did", "will", "would", "could", "should", "may", "might", "must", "shall", "can", "need", "dare", "ought", "used", "to", "of", "in", "for", "on", "with", "at", "by", "from", "as", "into", "through", "during", "before", "after", "above", "below", "between", "under", "again", "further", "then", "once", "here", "there", "when", "where", "why", "how", "all", "each", "few", "more", "most", "other", "some", "such", "no", "nor", "not", "only", "own", "same", "so", "than", "too", "very", "just", "and", "but", "if", "or", "because", "until", "while", "about", "against", "between", "into", "through", "during", "before", "after", "above", "below", "up", "down", "out", "off", "over", "under", "again", "further", "then", "once"}
BOILERPLATE_PATTERNS = [r"page\s*\d+\s*of\s*\d+", r"\d+\s*/\s*\d+", r"confidential", r"all rights reserved", r"copyright\s*[©®™]?\s*\d{4}", r" proprietary ", r"terms of use", r"privacy policy", r"\bwww\.\S+\b", r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b", r"tel[:\.]?\s*\+?[\d\s\-\(\)]{7,}", r"fax[:\.]?\s*\+?[\d\s\-\(\)]{7,}"]
URL_PATTERN = re.compile(r"https?://\S+|www\.\S+")
WHITESPACE_PATTERN = re.compile(r"\s+")

def normalize_text(text): return WHITESPACE_PATTERN.sub(" ", unicodedata.normalize("NFKC", text)).strip()
def remove_boilerplate(text):
    text = URL_PATTERN.sub("", text)
    for pattern in BOILERPLATE_PATTERNS: text = re.sub(pattern, "", text, flags=re.IGNORECASE)
    return text
def compress_paragraph(text, aggressive=False):
    words = text.split()
    filtered = [w for w in words if w.lower() not in STOPWORDS and len(w) > 1] if aggressive else [w for w in words if w.lower() not in STOPWORDS or len(w) > 3]
    return " ".join(filtered)
def table_to_markdown(table):
    if not table or len(table) < 1: return ""
    lines = []
    for i, row in enumerate(table):
        cells = [re.sub(r"\s+", " ", str(cell).strip() if cell is not None else "") for cell in row]
        lines.append("| " + " | ".join(cells) + " |")
        if i == 0: lines.append("|" + "|".join([" --- " for _ in cells]) + "|")
    return "\n".join(lines)
def extract_page_content(page, aggressive=False):
    chunks = []; tables = page.find_tables()
    text = page.extract_text()
    if text:
        text = remove_boilerplate(normalize_text(text))
        paragraphs = [p.strip() for p in text.split("\n") if p.strip()]
        compressed_paras = [compress_paragraph(p, aggressive) for p in paragraphs]
        chunks.append("\n\n".join([p for p in compressed_paras if p and len(p) > 3]))
    for table in tables:
        md_table = table_to_markdown(table.extract())
        if md_table: chunks.append("\n\n" + md_table + "\n")
    return "\n\n".join(chunks)

# ==========================================
# APP 2: SEISMIC DATA MINER LOGIC
# ==========================================
OUTPUT_COLUMNS = ("study_name", "reception_date", "support_type", "support_number", "swath_profile", "ffid", "vp", "data_type", "centre_traitement", "OBESERVATION", "source_file")
SUPPORT_COLOURS = {"Cassette": PatternFill("solid", fgColor="FFF2CC"), "CD/DVD": PatternFill("solid", fgColor="DEEAF6"), "USB": PatternFill("solid", fgColor="E2EFDA"), "DOC": PatternFill("solid", fgColor="F2F2F2"), "Unknown": PatternFill("solid", fgColor="FFFFFF")} if load_workbook else {}
FAMILY_SIT2, FAMILY_SIT1, FAMILY_STACK, FAMILY_RECEPTION, FAMILY_TEMPLATE = "SIT2", "SIT1", "STACK", "RECEPTION", "TEMPLATE"

class SeismicRecord:
    __slots__ = OUTPUT_COLUMNS
    def __init__(self):
        for col in OUTPUT_COLUMNS: setattr(self, col, None)

def make_record(study_name, reception_date, support_type, support_number, swath_profile="", data_type="", centre_traitement="", observation="", source_file=""):
    r = SeismicRecord()
    r.study_name = _strip(study_name) or None; r.reception_date = reception_date; r.support_type = support_type
    r.support_number = max(support_number, 1); r.swath_profile = _strip(swath_profile) or None; r.ffid = None; r.vp = None
    r.data_type = _strip(data_type) or None; r.centre_traitement = _strip(centre_traitement) or None
    r.OBESERVATION = _strip(observation) or None; r.source_file = source_file
    return r

def _emit_supports(n_cass, n_cd, n_usb, n_doc, study, rec_date, swath, data_type, centre, observation, fname):
    rows = []
    for i in range(n_cass): rows.append(make_record(study, rec_date, "Cassette", i + 1, swath, data_type, centre, observation, fname))
    for i in range(n_cd): rows.append(make_record(study, rec_date, "CD/DVD", i + 1, swath, data_type, centre, observation, fname))
    for i in range(n_usb): rows.append(make_record(study, rec_date, "USB", i + 1, swath, data_type, centre, observation, fname))
    for i in range(n_doc): rows.append(make_record(study, rec_date, "DOC", i + 1, swath, data_type, centre, observation, fname))
    return rows

def _strip(val):
    if val is None: return ""
    s = str(val).strip()
    return "" if s in ("None", "nan") else s

def _to_int(val):
    try: return int(float(val))
    except: return 0

def _xldate(serial, book):
    if not serial or not isinstance(serial, (int, float)): return None
    try:
        y, m, d, *_ = xlrd.xldate_as_tuple(serial, book.datemode)
        return date(y, m, d) if y else None
    except: return None

def _to_date(raw, book=None):
    if raw is None: return None
    if isinstance(raw, datetime): return raw.date()
    if isinstance(raw, date): return raw
    if isinstance(raw, (int, float)) and book is not None: return _xldate(raw, book)
    return None

def _normalise_support(raw):
    r = raw.upper()
    if any(k in r for k in ("CASS", "3592", "CASSETTE", "CART")): return "Cassette"
    if any(k in r for k in ("CD", "DVD", "DISC", "DISQUE")): return "CD/DVD"
    if any(k in r for k in ("USB", "CLÉ", "CLE", "FLASH")): return "USB"
    if "DOC" in r: return "DOC"
    return "Unknown"

def _join_obs(*parts): return " / ".join(p.strip() for p in parts if p and p.strip())

def _xls_obs(structure, trailing, sheet_name):
    parts = []
    s = _strip(structure)
    if s and len(s) > 1: parts.append(f"Structure : {s}")
    if sheet_name.strip().upper() == "ASS": parts.append("ASSOCIATION")
    t = _strip(trailing)
    if t: parts.append(t)
    return _join_obs(*parts)

class SeismicConsolidator:
    def __init__(self, log_cb): self.log_cb = log_cb

    def _is_template_xls(self, filepath: Path) -> bool:
        if filepath.suffix.lower() != ".xls": return False
        try:
            book = xlrd.open_workbook(str(filepath)); sheet = book.sheets()[0]
            for r in range(min(10, sheet.nrows)):
                if "directions" in " ".join(_strip(sheet.cell_value(r, c)).lower() for c in range(sheet.ncols)) and "date de rcep" in " ".join(_strip(sheet.cell_value(r, c)).lower() for c in range(sheet.ncols)): return True
        except: pass
        return False

    def _is_output_template(self, filepath: Path) -> bool:
        try:
            wb = load_workbook(str(filepath), read_only=True, data_only=True); val = wb.active.cell(1, 1).value; wb.close()
            return bool(val and _strip(val).lower() == "study_name")
        except: return False

    def detect_family(self, filepath: Path) -> Optional[str]:
        suffix = filepath.suffix.lower()
        if suffix == ".xlsx": return None if self._is_output_template(filepath) else FAMILY_RECEPTION
        if suffix == ".xls":
            if self._is_template_xls(filepath): return FAMILY_TEMPLATE
            stem = filepath.stem.lower()
            if "stack" in stem: return FAMILY_STACK
            if stem.startswith("sit1"): return FAMILY_SIT1
            return FAMILY_SIT2
        return None

    def discover_sources(self, directory: Path, output_path: Path) -> list:
        return [f for f in sorted(directory.glob("*.*")) if f.resolve() != output_path.resolve() and self.detect_family(f) is not None]

    def find_reception_template(self, directory: Path, output_path: Path) -> Path:
        candidates = [f for f in sorted(directory.glob("*.xlsx")) if f.resolve() != output_path.resolve() and self._is_output_template(f)]
        if not candidates: raise FileNotFoundError("No .xlsx output template found. Expected A1 == 'study_name'.")
        return candidates[0]

    def parse_sit2(self, filepath: Path) -> list:
        records = []; fname = filepath.name
        try: book = xlrd.open_workbook(str(filepath))
        except Exception as e: self.log_cb(f"⚠️ Cannot open {fname}: {e}"); return records
        for sheet_name in [s for s in ("SH", "ASS") if s in book.sheet_names()]:
            sheet = book.sheet_by_name(sheet_name); hdr_row = 0
            for r in range(min(10, sheet.nrows)):
                if any("campagne" in _strip(sheet.cell_value(r, c)).lower() for c in range(sheet.ncols)): hdr_row = r; break
            hdr = [_strip(sheet.cell_value(hdr_row, c)).lower() for c in range(sheet.ncols)]
            try: camp_col = next(i for i, h in enumerate(hdr) if "campagne" in h)
            except StopIteration: continue
            has_structure = camp_col > 0; trailing_start = camp_col + 6
            for row_idx in range(hdr_row + 1, sheet.nrows):
                row = [sheet.cell_value(row_idx, c) for c in range(sheet.ncols)]; study = _strip(row[camp_col])
                if not study or study.upper() == "TOTAL": continue
                swath = _strip(row[camp_col + 1]) if camp_col + 1 < len(row) else ""; n_raw = row[camp_col + 2] if camp_col + 2 < len(row) else ""
                data_fmt = _strip(row[camp_col + 3]) if camp_col + 3 < len(row) else ""; date_raw = row[camp_col + 4] if camp_col + 4 < len(row) else None
                nat_raw = _strip(row[camp_col + 5]) if camp_col + 5 < len(row) else ""; structure = _strip(row[0]) if has_structure else ""
                trailing = _join_obs(*[_strip(row[c]) for c in range(trailing_start, len(row)) if _strip(row[c])])
                obs = _xls_obs(structure, trailing, sheet_name); rec_date = _xldate(date_raw, book) if isinstance(date_raw, (int, float)) else None
                supp_type = ("Cassette" if nat_raw in ("3592.0", "3592", "3592,0") else _normalise_support(nat_raw or data_fmt))
                n_supp = max(_to_int(n_raw), 1) if n_raw != "" else 1
                for i in range(n_supp): records.append(make_record(study, rec_date, supp_type, i + 1, swath, data_fmt, "", obs, fname))
        return records

    def parse_sit1(self, filepath: Path) -> list:
        records = []; fname = filepath.name
        try: book = xlrd.open_workbook(str(filepath))
        except Exception as e: self.log_cb(f"⚠️ Cannot open {fname}: {e}"); return records
        for sheet_name in ("ASS", "SH"):
            if sheet_name not in book.sheet_names(): continue
            sheet = book.sheet_by_name(sheet_name); hdr1, hdr2 = 1, 2
            for r in range(min(10, sheet.nrows)):
                if any("structure" in _strip(sheet.cell_value(r, c)).lower() for c in range(sheet.ncols)): hdr1, hdr2 = r, r + 1; break
            sub_hdr = [_strip(sheet.cell_value(hdr2, c)).lower() for c in range(sheet.ncols)]
            cass_col = next((i for i, h in enumerate(sub_hdr) if "cass" in h), 6); cd_col = next((i for i, h in enumerate(sub_hdr) if "cd" in h), 7)
            doc_col = next((i for i, h in enumerate(sub_hdr) if "doc" in h), 8); nature_col = next((i for i, h in enumerate(sub_hdr) if "nature" in h or "supp" in h), 9)
            trailing_start = nature_col + 1
            for row_idx in range(hdr2 + 1, sheet.nrows):
                row = [sheet.cell_value(row_idx, c) for c in range(sheet.ncols)]; study = _strip(row[1]) if len(row) > 1 else ""; swath = _strip(row[2]) if len(row) > 2 else ""
                if not study or study.upper() in ("STRUCTURE", "ETUDE"): continue
                if "total" in _strip(row[3] if len(row) > 3 else "").lower(): continue
                n_cass = _to_int(row[cass_col]) if cass_col < len(row) else 0; n_cd = _to_int(row[cd_col]) if cd_col < len(row) else 0; n_doc = _to_int(row[doc_col]) if doc_col < len(row) else 0
                if n_cass == 0 and n_cd == 0 and n_doc == 0: continue
                structure = _strip(row[0]); trailing = _join_obs(*[_strip(row[c]) for c in range(trailing_start, len(row)) if _strip(row[c])])
                obs = _xls_obs(structure, trailing, sheet_name); records.extend(_emit_supports(n_cass, n_cd, 0, n_doc, study, None, swath, "", "", obs, fname))
            if records: break
        return records

    def parse_stack(self, filepath: Path) -> list:
        records = []; fname = filepath.name
        try: book = xlrd.open_workbook(str(filepath))
        except Exception as e: self.log_cb(f"⚠️ Cannot open {fname}: {e}"); return records
        for sheet_name in [s for s in ("SH", "ASS") if s in book.sheet_names()]:
            sheet = book.sheet_by_name(sheet_name); hdr1, hdr2 = 4, 5
            for r in range(min(10, sheet.nrows)):
                if any("structure" in _strip(sheet.cell_value(r, c)).lower() for c in range(sheet.ncols)): hdr1, hdr2 = r, r + 1; break
            main_hdr = [_strip(sheet.cell_value(hdr1, c)).lower() for c in range(sheet.ncols)]; sub_hdr = [_strip(sheet.cell_value(hdr2, c)).lower() for c in range(sheet.ncols)]
            cass_col = next((i for i, h in enumerate(sub_hdr) if "cass" in h), 4); cd_col = next((i for i, h in enumerate(sub_hdr) if "cd" in h), 5)
            usb_col = next((i for i, h in enumerate(sub_hdr) if "usb" in h), None); doc_col = next((i for i, h in enumerate(sub_hdr) if "doc" in h), None)
            centre_col = next((i for i, h in enumerate(main_hdr) if "centre" in h), None); date_col = next((i for i, h in enumerate(main_hdr) if "date" in h), 3)
            if sheet_name == "ASS":
                for c in range(sheet.ncols - 1, -1, -1):
                    if "date" in _strip(sheet.cell_value(hdr1, c)).lower(): date_col = c; break
            last_named = max([i for i, h in enumerate(main_hdr) if h] + [i for i, h in enumerate(sub_hdr) if h]); trailing_start = last_named + 1
            for row_idx in range(hdr2 + 1, sheet.nrows):
                row = [sheet.cell_value(row_idx, c) for c in range(sheet.ncols)]; study = _strip(row[1]) if len(row) > 1 else ""; swath = _strip(row[2]) if len(row) > 2 else ""
                if not study or study.upper() in ("STRUCTURE", "ETUDE"): continue
                if "total" in _strip(row[3] if len(row) > 3 else "").lower(): continue
                rec_date = _xldate(row[date_col], book) if date_col < len(row) and isinstance(row[date_col], (int, float)) else None
                centre = _strip(row[centre_col]) if centre_col and centre_col < len(row) else ""
                n_cass = _to_int(row[cass_col]) if cass_col < len(row) else 0; n_cd = _to_int(row[cd_col]) if cd_col < len(row) else 0
                n_usb = _to_int(row[usb_col]) if usb_col is not None and usb_col < len(row) else 0; n_doc = _to_int(row[doc_col]) if doc_col is not None and doc_col < len(row) else 0
                if n_cass == 0 and n_cd == 0 and n_usb == 0 and n_doc == 0: continue
                structure = _strip(row[0]); trailing = _join_obs(*[_strip(row[c]) for c in range(trailing_start, len(row)) if _strip(row[c])])
                obs = _xls_obs(structure, trailing, sheet_name); records.extend(_emit_supports(n_cass, n_cd, n_usb, n_doc, study, rec_date, swath, "stack", centre, obs, fname))
        return records

    def parse_reception(self, filepath: Path) -> list:
        records = []; fname = filepath.name
        try: wb = load_workbook(str(filepath), read_only=True, data_only=True)
        except Exception as e: self.log_cb(f"⚠️ Cannot open {fname}: {e}"); return records
        ws = wb.active; header = [ws.cell(1, c).value for c in range(1, (ws.max_column or 0) + 1)]
        kw_map = {"date": "date", "type": "data_type", "compagnie": "centre", "etude": "study", "profil": "swath", "cassette": "cass", "cd": "cd", "doc": "doc", "usb": "usb", "observ": "obs"}
        col_map = {}
        for idx, val in enumerate(header):
            h = _strip(val).lower()
            for kw, field in kw_map.items():
                if kw in h and field not in col_map: col_map[field] = idx; break
        missing = {"study", "cass", "cd"} - col_map.keys()
        if missing: self.log_cb(f"⚠️ {fname}: Missing {missing}"); wb.close(); return records
        def _get(row_vals, key, default=None): idx = col_map.get(key); return row_vals[idx] if (idx is not None and idx < len(row_vals)) else default
        for row in ws.iter_rows(min_row=2, values_only=True):
            row = list(row); study = _strip(_get(row, "study"))
            if not study: continue
            rec_date = _to_date(_get(row, "date")); swath = _strip(_get(row, "swath")); data_type = _strip(_get(row, "data_type"))
            centre = _strip(_get(row, "centre")); obs = _strip(_get(row, "obs"))
            n_cass = _to_int(_get(row, "cass", 0)); n_cd = _to_int(_get(row, "cd", 0)); n_usb = _to_int(_get(row, "usb", 0)); n_doc = _to_int(_get(row, "doc", 0))
            if n_cass == 0 and n_cd == 0 and n_usb == 0 and n_doc == 0: continue
            records.extend(_emit_supports(n_cass, n_cd, n_usb, n_doc, study, rec_date, swath, data_type, centre, obs, fname))
        wb.close(); return records

    def parse_template(self, filepath: Path) -> list:
        records = []; fname = filepath.name
        try: book = xlrd.open_workbook(str(filepath))
        except Exception as e: self.log_cb(f"⚠️ Cannot open {fname}: {e}"); return records
        sheet = book.sheets()[0]
        section_rows = [r for r in range(sheet.nrows) if "directions" in " ".join(_strip(sheet.cell_value(r, c)).lower() for c in range(sheet.ncols))]
        if not section_rows: return records
        for idx, hdr_row in enumerate(section_rows):
            next_hdr = section_rows[idx + 1] if idx + 1 < len(section_rows) else sheet.nrows
            row_text = " ".join(_strip(sheet.cell_value(hdr_row, c)).lower() for c in range(sheet.ncols))
            if "date de rcep" in row_text:
                for row_idx in range(hdr_row + 2, next_hdr):
                    row = [sheet.cell_value(row_idx, c) for c in range(sheet.ncols)]; study = _strip(row[1]) if len(row) > 1 else ""
                    if not study or any("total" in v.lower() for v in [_strip(v) for v in row if _strip(v)]): continue
                    n_cass = _to_int(row[6]) if len(row) > 6 else 0; n_cd = _to_int(row[7]) if len(row) > 7 else 0; n_doc = _to_int(row[8]) if len(row) > 8 else 0
                    if n_cass == 0 and n_cd == 0 and n_doc == 0: continue
                    records.extend(_emit_supports(n_cass, n_cd, 0, n_doc, study, _to_date(row[9] if len(row) > 9 else None, book), _strip(row[2]) if len(row) > 2 else "", "terrain", "", _join_obs(f"Structure : {_strip(row[0])}" if len(row) > 0 and _strip(row[0]) and len(_strip(row[0])) > 1 else "", _strip(row[11]) if len(row) > 11 else ""), fname))
            elif any(k in row_text for k in ("date de r", "réception", "reception")):
                main_hdr = [_strip(sheet.cell_value(hdr_row, c)).lower() for c in range(sheet.ncols)]; sub_hdr = [_strip(sheet.cell_value(hdr_row + 1, c)).lower() for c in range(sheet.ncols)]
                date_col = next((i for i, h in enumerate(main_hdr) if "date" in h), 3); cass_col = next((i for i, h in enumerate(sub_hdr) if "cass" in h), 4)
                cd_col = next((i for i, h in enumerate(sub_hdr) if "cd" in h), 5); doc_col = next((i for i, h in enumerate(sub_hdr) if "doc" in h), 6)
                usb_col = next((i for i, h in enumerate(sub_hdr) if "usb" in h), 7); centre_col = next((i for i, h in enumerate(main_hdr) if "centre" in h), 9)
                for row_idx in range(hdr_row + 2, next_hdr):
                    row = [sheet.cell_value(row_idx, c) for c in range(sheet.ncols)]; study = _strip(row[1]) if len(row) > 1 else ""
                    if not study or any("total" in v.lower() for v in [_strip(v) for v in row if _strip(v)]): continue
                    n_cass = _to_int(row[cass_col]) if cass_col < len(row) else 0; n_cd = _to_int(row[cd_col]) if cd_col < len(row) else 0
                    n_doc = _to_int(row[doc_col]) if doc_col < len(row) else 0; n_usb = _to_int(row[usb_col]) if usb_col < len(row) else 0
                    if n_cass == 0 and n_cd == 0 and n_doc == 0 and n_usb == 0: continue
                    records.extend(_emit_supports(n_cass, n_cd, n_usb, n_doc, study, _to_date(row[date_col] if date_col < len(row) else None, book), _strip(row[2]) if len(row) > 2 else "", "stack", _strip(row[centre_col]) if centre_col < len(row) else "", _join_obs(f"Structure : {_strip(row[0])}" if len(row) > 0 and _strip(row[0]) and len(_strip(row[0])) > 1 else ""), fname))
        return records

    def parse_file(self, filepath: Path) -> list:
        family = self.detect_family(filepath)
        if family == FAMILY_SIT2: return self.parse_sit2(filepath)
        if family == FAMILY_SIT1: return self.parse_sit1(filepath)
        if family == FAMILY_STACK: return self.parse_stack(filepath)
        if family == FAMILY_RECEPTION: return self.parse_reception(filepath)
        if family == FAMILY_TEMPLATE: return self.parse_template(filepath)
        return []

    def _capture_row_style(self, ws, row_num: int) -> dict:
        styles = {}
        for cell in next(ws.iter_rows(min_row=row_num, max_row=row_num)):
            styles[cell.column] = {"font": copy.copy(cell.font) if cell.font else None, "border": copy.copy(cell.border) if cell.border else None, "alignment": copy.copy(cell.alignment) if cell.alignment else None, "number_format": cell.number_format}
        return styles

    def write_output(self, records: list, template_path: Path, output_path: Path) -> None:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb = load_workbook(str(template_path)); ws = wb.active
        try: row_style = self._capture_row_style(ws, 3)
        except: row_style = {}
        write_row = 2
        for row_idx in range(ws.max_row, 1, -1):
            if ws.cell(row=row_idx, column=1).value is not None: write_row = row_idx + 1; break
        col_map = {name: idx + 1 for idx, name in enumerate(OUTPUT_COLUMNS)}
        k_col = col_map["source_file"]
        if _strip(ws.cell(1, k_col).value).lower() != "source_file":
            hdr_cell = ws.cell(1, k_col); hdr_cell.value = "source_file"; ref = ws.cell(1, 1)
            hdr_cell.font = copy.copy(ref.font); hdr_cell.fill = copy.copy(ref.fill); hdr_cell.border = copy.copy(ref.border); hdr_cell.alignment = copy.copy(ref.alignment)
        self.log_cb(f"> Writing {len(records)} rows starting at row {write_row}...")
        for rec in records:
            fill = SUPPORT_COLOURS.get(_strip(rec.support_type), SUPPORT_COLOURS["Unknown"])
            for col_name, col_idx in col_map.items():
                cell = ws.cell(row=write_row, column=col_idx); cell.value = getattr(rec, col_name); cell.fill = fill
                if col_idx in row_style:
                    st = row_style[col_idx]
                    if st["font"]: cell.font = copy.copy(st["font"])
                    if st["border"]: cell.border = copy.copy(st["border"])
                    if st["alignment"]: cell.alignment = copy.copy(st["alignment"])
                    if st["number_format"]: cell.number_format = st["number_format"]
                if col_name == "reception_date" and isinstance(cell.value, date): cell.number_format = "DD/MM/YYYY"
            write_row += 1
        wb.save(str(output_path))
        self.log_cb(f"> Saved master workbook to {output_path.name}")

# ==========================================
# FLASK ROUTES
# ==========================================
@app.route('/')
def index(): return render_template('index.html')

@app.route('/synthese')
def synthese(): return render_template('synthese.html')

@app.route('/api/drive/list', methods=['POST'])
def api_list():
    data = request.get_json(); lines = list_contents(data.get('path'), data.get('show_hidden', False), data.get('show_sizes', True), data.get('recursive', False))
    if data.get('export'):
        mem = io.BytesIO("\n".join(lines).encode('utf-8')); mem.seek(0)
        return send_file(mem, as_attachment=True, download_name=f"listing_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt", mimetype='text/plain')
    return jsonify({"lines": lines})

@app.route('/api/drive/compare', methods=['POST'])
def api_compare():
    data = request.get_json(); lines = compare_listings(data.get('path_a'), data.get('path_b'), data.get('show_hidden', False))
    if data.get('export'):
        mem = io.BytesIO("\n".join(lines).encode('utf-8')); mem.seek(0)
        return send_file(mem, as_attachment=True, download_name=f"compare_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt", mimetype='text/plain')
    return jsonify({"lines": lines})

@app.route('/api/drive/compare-files', methods=['POST'])
def api_compare_files():
    f_a, f_b = request.files.get('file_a'), request.files.get('file_b')
    if not f_a or not f_b: return jsonify({"error": "Both files required"}), 400
    with tempfile.NamedTemporaryFile(delete=False, suffix=".txt") as t_a, tempfile.NamedTemporaryFile(delete=False, suffix=".txt") as t_b:
        f_a.save(t_a.name); f_b.save(t_b.name)
        p_a, p_b = parse_listing_file(t_a.name), parse_listing_file(t_b.name)
        os.unlink(t_a.name); os.unlink(t_b.name)
    lines = compare_parsed_listings(p_a, p_b)
    if request.form.get('export') == 'true':
        mem = io.BytesIO("\n".join(lines).encode('utf-8')); mem.seek(0)
        return send_file(mem, as_attachment=True, download_name=f"txt_compare_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt", mimetype='text/plain')
    return jsonify({"lines": lines})

@app.route('/api/drive/roots', methods=['GET'])
def get_roots():
    roots = []
    if platform.system() == 'Windows':
        for letter in string.ascii_uppercase:
            if os.path.exists(f"{letter}:\\"): roots.append(f"{letter}:\\")
    else:
        roots.append('/')
        for p in ['/media', '/mnt']:
            if os.path.exists(p): roots.extend([f"{p}/{d}" for d in os.listdir(p)])
    return jsonify({"roots": roots})

@app.route('/api/drive/folders', methods=['POST'])
def get_folders():
    data = request.get_json(); target_path = data.get('path', '').strip()
    if not target_path or not os.path.exists(target_path) or not os.path.isdir(target_path): return jsonify({"error": "Invalid path"}), 400
    try:
        dirs = [{"name": entry.name, "full_path": entry.path} for entry in os.scandir(target_path) if entry.is_dir() and entry.name not in ['$RECYCLE.BIN', 'System Volume Information']]
        dirs.sort(key=lambda x: x['name'].lower()); parent = os.path.dirname(target_path)
        return jsonify({"current": target_path, "parent": parent if parent != target_path else None, "dirs": dirs})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/api/pdf/compress', methods=['POST'])
def api_compress_pdf():
    if pdfplumber is None: return jsonify({"error": "pdfplumber not installed"}), 500
    file = request.files['file']; aggressive = request.form.get('aggressive') == 'true'
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        file.save(temp_pdf.name); temp_path = temp_pdf.name
    try:
        original_size = os.path.getsize(temp_path); all_pages = []
        with pdfplumber.open(temp_path) as pdf:
            for i, page in enumerate(pdf.pages, 1):
                page_md = extract_page_content(page, aggressive)
                if page_md.strip(): all_pages.append(f"\n\n<!-- Page {i} -->\n\n{page_md}")
        full_md = re.sub(r"\n{3,}", "\n\n", WHITESPACE_PATTERN.sub(" ", "\n".join(all_pages).strip()))
        full_md = f"# {os.path.splitext(file.filename)[0]}\n> **Mode:** {'Aggressive' if aggressive else 'Standard'}\n---\n" + full_md
        compressed_size = len(full_md.encode('utf-8'))
        return jsonify({"markdown": full_md, "original_kb": round(original_size / 1024, 2), "compressed_kb": round(compressed_size / 1024, 2), "reduction_percent": round(((original_size - compressed_size) / original_size) * 100, 1)})
    except Exception as e: return jsonify({"error": str(e)}), 500
    finally: os.unlink(temp_path)

# ==========================================
# APP 2: SEISMIC MINER ROUTES
# ==========================================
@app.route('/api/miner/stream', methods=['POST'])
def miner_stream():
    template_file = request.files.get('template')
    source_files = request.files.getlist('sources')
    
    temp_dir = Path(tempfile.mkdtemp())
    output_id = str(uuid.uuid4())
    output_filename = f"receptions_filled_{output_id}.xlsx"
    output_dir = Path("outputs")
    output_dir.mkdir(exist_ok=True)
    output_path = output_dir / output_filename

    template_filename = os.path.basename(template_file.filename)
    template_path = temp_dir / template_filename
    template_file.save(str(template_path))
    
    for f in source_files:
        source_filename = os.path.basename(f.filename)
        f.save(str(temp_dir / source_filename))

    log_queue = queue.Queue()

    def run_task():
        try:
            log_queue.put(f"> Template loaded: {template_path.name}")
            log_queue.put(f"> {len(source_files)} source files staged in temp environment.")

            consolidator = SeismicConsolidator(log_queue.put)
            
            try:
                rec_template = consolidator.find_reception_template(temp_dir, output_path)
                log_queue.put(f"> Output template identified: {rec_template.name}")
            except FileNotFoundError as e:
                log_queue.put(f"❌ Error: {str(e)}")
                log_queue.put("DONE_ERROR"); return

            sources = consolidator.discover_sources(temp_dir, output_path)
            log_queue.put(f"> Discovered {len(sources)} processable source files.")
            if not sources: 
                log_queue.put("❌ Error: No valid source files found.")
                log_queue.put("DONE_ERROR"); return

            all_records = []
            for filepath in sources:
                family = consolidator.detect_family(filepath)
                log_queue.put(f"> Parsing {filepath.name} ({family})...")
                try: 
                    all_records.extend(consolidator.parse_file(filepath))
                except Exception as exc: 
                    log_queue.put(f"⚠️ Error parsing {filepath.name}: {exc}")

            log_queue.put(f"> Total records extracted: {len(all_records)}")
            if not all_records: 
                log_queue.put("❌ Error: No records extracted. Aborting.")
                log_queue.put("DONE_ERROR"); return

            consolidator.write_output(all_records, rec_template, output_path)
            log_queue.put(f"✅ Pipeline completed successfully.")
            log_queue.put(f"DOWNLOAD_URL:/api/miner/download/{output_id}")
            log_queue.put("DONE_SUCCESS")
        except Exception as e:
            log_queue.put(f"❌ FATAL: {str(e)}")
            log_queue.put("DONE_ERROR")
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    threading.Thread(target=run_task).start()

    def generate():
        while True:
            try:
                msg = log_queue.get(timeout=30)
                if msg == "DONE_SUCCESS": 
                    yield f"data: {json.dumps({'log': '✅ Done', 'done': True, 'success': True})}\n\n"
                    break
                elif msg == "DONE_ERROR": 
                    yield f"data: {json.dumps({'log': '❌ Failed', 'done': True, 'success': False})}\n\n"
                    break
                elif str(msg).startswith("DOWNLOAD_URL:"): 
                    yield f"data: {json.dumps({'log': 'Generating download link...', 'download_url': str(msg).split(':', 1)[1]})}\n\n"
                else: 
                    yield f"data: {json.dumps({'log': str(msg)})}\n\n"
            except queue.Empty: 
                yield f"data: {json.dumps({'log': '...'})}\n\n"

    return Response(generate(), mimetype='text/event-stream')

@app.route('/api/miner/download/<file_id>')
def miner_download(file_id):
    filepath = Path("outputs") / f"receptions_filled_{file_id}.xlsx"
    if filepath.exists(): return send_file(filepath, as_attachment=True, download_name="receptions_filled.xlsx")
    return "File not found", 404

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)