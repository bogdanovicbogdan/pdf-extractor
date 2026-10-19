import os
import sys
import io
import re
import argparse
from collections import Counter
import fitz
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image

sys.stdout.reconfigure(encoding='utf-8')


DEFAULT_EXCLUDE_TEXTS = [
]

DEFAULT_REMOVE_LOGOS = True
DEFAULT_AUTO_FOOTER = True


PUA_BULLET_CHARS = {
    '\uf07d', '\uf0e3', '\uf06c', '\uf0d8', '\uf0a7', '\uf0b7', '\uf0fc',
    '\uf076', '\uf09f', '\uf077', '\uf0a8', '\uf0fe', '\uf02a', '\uf06e',
    '\uf075', '\uf0d9',  
}

PUA_MATH_MAP = {
    '\uf02b': '+',
    '\uf02d': '−',
    '\uf0b1': '±',
    '\uf0b4': '×',
    '\uf0b8': '÷',
    '\uf0b9': '≠',
    '\uf0a3': '≤',
    '\uf0b3': '≥',
    '\uf0bb': '≈',
    '\uf0a5': '∞',
    '\uf0ae': '→',
    '\uf0ac': '←',
    '\uf0ab': '↔',
    '\uf0de': '⇒',
    '\uf0db': '⇔',
    '\uf0ce': '∈',
    '\uf0cf': '∉',
    '\uf0cc': '⊂',
    '\uf0cd': '⊆',
    '\uf0c7': '∩',
    '\uf0c8': '∪',
    '\uf0d9': '∧',
    '\uf0da': '∨',
    '\uf0d8': '¬',
    '\uf022': '∀',
    '\uf024': '∃',
    '\uf066': '∅',
    '\uf049': '⋂',
    '\uf055': '⋃',
    '\uf0e5': '∑',
    '\uf0d5': '∏',
    '\uf0be': '', 
}

SYSTEM_FONT_CANDIDATES = {
    'calibri': r'C:\Windows\Fonts\calibri.ttf',
    'calibrib': r'C:\Windows\Fonts\calibrib.ttf',
    'calibrii': r'C:\Windows\Fonts\calibrii.ttf',
    'arial': r'C:\Windows\Fonts\arial.ttf',
    'arialbd': r'C:\Windows\Fonts\arialbd.ttf',
    'times': r'C:\Windows\Fonts\times.ttf',
    'consola': r'C:\Windows\Fonts\consola.ttf',
    'tahoma': r'C:\Windows\Fonts\tahoma.ttf',
    'verdana': r'C:\Windows\Fonts\verdana.ttf',
    'cambria': r'C:\Windows\Fonts\cambria.ttc',
}


# POPRAVKA FONTOVA I SRPSKIH SLOVA (ToUnicode CMap Repair)
def _parse_tounicode_cmap(cmap_str):
    mapping = {}
    for block in re.findall(r'beginbfchar(.*?)endbfchar', cmap_str, re.S):
        for m in re.finditer(r'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', block):
            cid = int(m.group(1), 16)
            uni_hex = m.group(2)
            if len(uni_hex) <= 4:
                mapping[cid] = int(uni_hex, 16)
    for block in re.findall(r'beginbfrange(.*?)endbfrange', cmap_str, re.S):
        for line in block.strip().splitlines():
            line = line.strip()
            if not line:
                continue
            m_arr = re.match(r'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*\[(.*)\]', line)
            if m_arr:
                start = int(m_arr.group(1), 16)
                end = int(m_arr.group(2), 16)
                vals = [int(x, 16) for x in re.findall(r'<([0-9A-Fa-f]+)>', m_arr.group(3))]
                for i, cid in enumerate(range(start, end + 1)):
                    if i < len(vals):
                        mapping[cid] = vals[i]
                continue
            m_rng = re.match(r'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', line)
            if m_rng:
                start = int(m_rng.group(1), 16)
                end = int(m_rng.group(2), 16)
                base = int(m_rng.group(3), 16)
                for i, cid in enumerate(range(start, end + 1)):
                    mapping[cid] = base + i
    return mapping


def _build_tounicode_stream(mapping):
    bfchars = [f"<{cid:04X}> <{uni:04X}>" for cid, uni in sorted(mapping.items()) if 0 <= uni <= 0xFFFF]
    chunks = []
    for i in range(0, len(bfchars), 100):
        chunk = bfchars[i:i + 100]
        chunks.append(f"{len(chunk)} beginbfchar\n" + "\n".join(chunk) + "\nendbfchar")
    return (
        "/CIDInit /ProcSet findresource begin\n"
        "12 dict begin\n"
        "begincmap\n"
        "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n"
        "/CMapName /Adobe-Identity-UCS def\n"
        "/CMapType 2 def\n"
        "1 begincodespacerange\n"
        "<0000> <FFFF>\n"
        "endcodespacerange\n" +
        "\n".join(chunks) + "\n"
        "endcmap\n"
        "CMapName currentdict /CMap defineresource pop\n"
        "end\n"
        "end\n"
    ).encode('latin1')


def fix_pdf_tounicode(doc):
    scan_codepoints = (
        list(range(0x0020, 0x0250)) +  # Osnovna latinica + srpska slova (č, ć, ž, š, đ, Č, Ć, Ž, Š, Đ)
        list(range(0x0370, 0x04FF)) +  # Grčki alfabet i ćirilica
        list(range(0x2000, 0x2300)) +  # Interpunkcija, strelice, matematički operatori
        list(range(0xF000, 0xF100))    # Private Use Area (Wingdings / Symbol)
    )

    for xref in range(1, doc.xref_length()):
        try:
            obj = doc.xref_object(xref)
        except Exception:
            continue
        if '/ToUnicode' not in obj:
            continue
        tu_m = re.search(r'/ToUnicode\s+(\d+)\s+0\s+R', obj)
        if not tu_m:
            continue
        tu_xref = int(tu_m.group(1))

        cid_xref = xref
        df_match = re.search(r'/DescendantFonts\s+(\d+)\s+0\s+R', obj)
        if df_match:
            try:
                df_obj = doc.xref_object(int(df_match.group(1)))
                m_cid = re.search(r'(\d+)\s+0\s+R', df_obj)
                if m_cid:
                    cid_xref = int(m_cid.group(1))
            except Exception:
                pass
        else:
            m2 = re.search(r'/DescendantFonts\s*\[\s*(\d+)\s+0\s+R', obj)
            if m2:
                cid_xref = int(m2.group(1))

        try:
            cid_obj = doc.xref_object(cid_xref)
        except Exception:
            continue

        ff2_xref = None
        fd_match = re.search(r'/FontDescriptor\s+(\d+)\s+0\s+R', cid_obj)
        if fd_match:
            try:
                fd_obj = doc.xref_object(int(fd_match.group(1)))
                ff2_m = re.search(r'/FontFile2\s+(\d+)\s+0\s+R', fd_obj)
                if ff2_m:
                    ff2_xref = int(ff2_m.group(1))
            except Exception:
                pass

        if not ff2_xref:
            continue

        try:
            ttf_bytes = doc.xref_stream(ff2_xref)
            f_emb = fitz.Font(fontbuffer=ttf_bytes)
        except Exception:
            continue

        gid_to_uni = {}
        for cp in scan_codepoints:
            try:
                g = f_emb.has_glyph(cp)
                if g > 0 and g not in gid_to_uni:
                    gid_to_uni[g] = cp
            except Exception:
                pass

        if not gid_to_uni and '/CIDToGIDMap /Identity' in cid_obj:
            base_m = re.search(r'/BaseFont\s+/([A-Za-z0-9+,\-_]+)', obj)
            if base_m:
                clean_font_name = re.sub(r'^[A-Z]{6}\+', '', base_m.group(1)).lower()
                for key, sys_path in SYSTEM_FONT_CANDIDATES.items():
                    if key in clean_font_name and os.path.exists(sys_path):
                        try:
                            f_sys = fitz.Font(fontfile=sys_path)
                            for cp in scan_codepoints:
                                g = f_sys.has_glyph(cp)
                                if g > 0 and g not in gid_to_uni:
                                    gid_to_uni[g] = cp
                        except Exception:
                            pass
                        break

        if not gid_to_uni:
            continue

        try:
            cmap_str = doc.xref_stream(tu_xref).decode('latin1', errors='ignore')
            tu_map = _parse_tounicode_cmap(cmap_str)
        except Exception:
            continue

        needs_patch = False
        merged = dict(tu_map)
        for cid, true_uni in gid_to_uni.items():
            if cid in merged and merged[cid] != true_uni:
                needs_patch = True
            merged[cid] = true_uni

        if needs_patch:
            try:
                doc.update_stream(tu_xref, _build_tounicode_stream(merged))
            except Exception:
                pass


def clean_line_characters(line):
    if not line:
        return ""

    line = line.replace('\uf0be\uf0ae', '→').replace('\uf0be', '')

    stripped = line.lstrip()
    leading_ws = line[:len(line) - len(stripped)]

    if stripped and stripped[0] in PUA_BULLET_CHARS:
        rest = stripped[1:].lstrip()
        if not rest:
            return ""
        stripped = "• " + rest
        line = leading_ws + stripped

    out_chars = []
    for ch in line:
        if ch in PUA_MATH_MAP:
            out_chars.append(PUA_MATH_MAP[ch])
        elif ch in PUA_BULLET_CHARS:
            out_chars.append('• ')
        elif 0xF000 <= ord(ch) <= 0xF0FF:
            out_chars.append('• ')
        else:
            out_chars.append(ch)

    result = "".join(out_chars)
    result = re.sub(r'^(•\s*)+', '• ', result.strip()) if result.strip().startswith('•') else result
    return result


def _normalize_for_match(text):
    t = text.lower().strip()
    replacements = {
        'č': 'c', 'ć': 'c', 'ž': 'z', 'š': 's', 'đ': 'dj',
    }
    t_ascii = "".join(replacements.get(c, c) for c in t)
    t_clean = re.sub(r'[^\w\s]', ' ', t)
    t_clean = re.sub(r'\s+', ' ', t_clean).strip()
    t_ascii_clean = re.sub(r'[^\w\s]', ' ', t_ascii)
    t_ascii_clean = re.sub(r'\s+', ' ', t_ascii_clean).strip()
    return t_clean, t_ascii_clean


def detect_repeated_slide_texts(doc):
    total_pages = len(doc)
    if total_pages < 3:
        return set()

    line_page_counts = Counter()
    for pno, page in enumerate(doc):
        ph = page.rect.height
        seen_on_page = set()
        for b in page.get_text("blocks"):
            y0, y1 = b[1], b[3]
            in_margin = (y0 <= 0.18 * ph) or (y1 >= 0.82 * ph)
            if not in_margin:
                continue
            for raw_line in b[4].splitlines():
                line = clean_line_characters(raw_line).strip()
                if not line or line.isdigit():
                    continue
                norm, norm_ascii = _normalize_for_match(line)
                if len(norm) >= 3:
                    seen_on_page.add(norm_ascii)
        for norm_ascii in seen_on_page:
            line_page_counts[norm_ascii] += 1

    threshold = max(3, int(total_pages * 0.30))
    repeated = {txt for txt, cnt in line_page_counts.items() if cnt >= threshold}
    return repeated


def should_exclude_line(line, y0, y1, page_height, pno, user_patterns, auto_repeated_set, pres_titles=None):
    t = line.strip()
    if not t:
        return True

    in_top_margin = (y0 <= 0.16 * page_height)
    in_bottom_margin = (y1 >= 0.82 * page_height) or (y0 >= 0.80 * page_height)
    in_margin_or_title_slide = in_top_margin or in_bottom_margin or (pno == 0)

    if (in_top_margin or in_bottom_margin) and re.match(r'^\d+(\s*/\s*\d+)?$', t):
        return True

    norm, norm_ascii = _normalize_for_match(t)
    if not norm_ascii:
        return True

    if norm_ascii in auto_repeated_set and in_margin_or_title_slide:
        return True

    if pno == 0 and auto_repeated_set:
        line_words = set(w for w in norm_ascii.split() if len(w) > 2)
        for rep in auto_repeated_set:
            rep_words = set(w for w in rep.split() if len(w) > 2)
            if len(rep_words) >= 2 and rep_words.issubset(line_words) and len(line_words) <= len(rep_words) + 3:
                return True

    if pres_titles and in_bottom_margin:
        for pt in pres_titles:
            _, pt_ascii = _normalize_for_match(pt)
            if pt_ascii and pt_ascii in norm_ascii and len(norm_ascii) <= len(pt_ascii) + 15:
                return True

    for pat in user_patterns:
        _, pat_ascii = _normalize_for_match(pat)
        if not pat_ascii:
            continue

        if norm_ascii == pat_ascii:
            return True

        if pat_ascii in norm_ascii:
            if in_margin_or_title_slide or len(norm_ascii) <= len(pat_ascii) + 25:
                return True

        if in_margin_or_title_slide:
            pat_words = [w for w in pat_ascii.split() if len(w) > 2]
            line_words = set(w for w in norm_ascii.split() if len(w) > 2)
            if len(pat_words) >= 2:
                matched = sum(1 for w in pat_words if w in line_words)
                if matched == len(pat_words) and len(line_words) <= len(pat_words) + 3:
                    return True

    return False


def detect_repeating_images(doc):
    if len(doc) < 2:
        return set(), set()
    xref_counts = Counter()
    sig_counts = Counter()
    for page in doc:
        seen_xrefs = set()
        seen_sigs = set()
        for im in page.get_images():
            xref = im[0]
            seen_xrefs.add(xref)
            try:
                base = doc.extract_image(xref)
                sig = (base['width'], base['height'], len(base['image']))
                seen_sigs.add(sig)
            except Exception:
                pass
        for x in seen_xrefs:
            xref_counts[x] += 1
        for s in seen_sigs:
            sig_counts[s] += 1

    threshold = 2 if len(doc) >= 3 else 2
    rep_xrefs = {x for x, c in xref_counts.items() if c >= threshold}
    rep_sigs = {s for s, c in sig_counts.items() if c >= threshold}
    return rep_xrefs, rep_sigs


def is_logo_or_decoration(rect, img_w, img_h, page_w, page_h, remove_logos=True):
    if (img_w <= 35 and img_h <= 35) or (rect.width <= 25 and rect.height <= 25):
        return True
    if (rect.height <= 10 and rect.width > 150) or (img_h <= 12 and img_w > 180):
        return True
    if (rect.width <= 10 and rect.height > 150) or (img_w <= 12 and img_h > 180):
        return True
    if rect.y0 >= 0.88 * page_h:
        return True

    if not remove_logos:
        return False

    if rect.width >= 0.85 * page_w and rect.height >= 0.85 * page_h:
        return True

    in_top_band = (rect.y1 <= 0.27 * page_h) and (rect.y0 <= 0.15 * page_h)
    in_left_corner = (rect.x1 <= 0.25 * page_w)
    in_right_corner = (rect.x0 >= 0.75 * page_w)
    is_small_emblem = (rect.width <= 0.22 * page_w) and (rect.height <= 0.25 * page_h)

    if in_top_band and (in_left_corner or in_right_corner) and is_small_emblem:
        return True

    in_bottom_band = (rect.y0 >= 0.78 * page_h)
    if in_bottom_band and (in_left_corner or in_right_corner) and is_small_emblem:
        return True

    return False


def deduplicate_images(img_entries):
    if len(img_entries) <= 1:
        return img_entries
    kept = []
    for item in img_entries:
        r1 = item['rect']
        overlap_found = False
        for k_idx, k_item in enumerate(kept):
            r2 = k_item['rect']
            ix0 = max(r1.x0, r2.x0)
            iy0 = max(r1.y0, r2.y0)
            ix1 = min(r1.x1, r2.x1)
            iy1 = min(r1.y1, r2.y1)
            if ix1 > ix0 and iy1 > iy0:
                inter_area = (ix1 - ix0) * (iy1 - iy0)
                a1 = max(1e-5, r1.width * r1.height)
                a2 = max(1e-5, r2.width * r2.height)
                if inter_area / min(a1, a2) > 0.6:
                    overlap_found = True
                    if item['w'] * item['h'] > k_item['w'] * k_item['h']:
                        kept[k_idx] = item
                    break
        if not overlap_found:
            kept.append(item)
    return kept


def detect_template_drawings(doc):
    if len(doc) < 3:
        return set()
    rect_counts = Counter()
    for page in doc:
        seen = set()
        for d in page.get_drawings():
            r = d['rect']
            seen.add((round(r.x0, 1), round(r.y0, 1), round(r.x1, 1), round(r.y1, 1)))
        for k in seen:
            rect_counts[k] += 1
    return {k for k, c in rect_counts.items() if c >= 3}


def get_vector_diagram_clip(page, template_drawing_rects):
    pw, ph = page.rect.width, page.rect.height
    non_tmpl = []
    for d in page.get_drawings():
        r = d['rect']
        if r.width >= 0.9 * pw and r.height >= 0.9 * ph:
            continue
        if r.y0 >= 0.88 * ph or r.y1 <= 0.10 * ph:
            continue
        key = (round(r.x0, 1), round(r.y0, 1), round(r.x1, 1), round(r.y1, 1))
        if key not in template_drawing_rects:
            non_tmpl.append(r)

    if len(non_tmpl) < 6:
        return None

    u = fitz.Rect(non_tmpl[0])
    for r in non_tmpl[1:]:
        u |= r

    if u.width < 80 or u.height < 40:
        return None

    expanded = fitz.Rect(u.x0 - 15, u.y0 - 15, u.x1 + 15, u.y1 + 15)
    for b in page.get_text("blocks"):
        br = fitz.Rect(b[0], b[1], b[2], b[3])
        if br.y0 < 0.12 * ph or br.y1 > 0.88 * ph:
            continue
        if br.intersects(expanded) and len(b[4].strip()) < 120:
            expanded |= br

    clip = fitz.Rect(
        max(10, expanded.x0 - 8),
        max(0.10 * ph, expanded.y0 - 8),
        min(pw - 10, expanded.x1 + 8),
        min(0.88 * ph, expanded.y1 + 8)
    )
    return clip


def insert_image(doc, img_bytes, w, h):
    try:
        img_stream = io.BytesIO(img_bytes)
        pil_img = Image.open(img_stream)
        pw, ph = pil_img.size
        img_stream.seek(0)

        max_w = 5.6
        max_h = 4.5
        aspect = pw / max(1, ph)

        if pw < 250 and ph < 250:
            disp_w = min(2.5, pw / 96.0)
        elif aspect > 1.2:
            disp_w = min(max_w, pw / 120.0)
            if disp_w < 3.5:
                disp_w = 3.5
        else:
            disp_w = min(max_w, max_h * aspect)
            if disp_w < 3.0:
                disp_w = 3.0

        disp_w = min(disp_w, max_w)

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True

        run = p.add_run()
        run.add_picture(img_stream, width=Inches(disp_w))
    except Exception as e:
        print(f"      [Upozorenje] Greška pri umetanju slike: {e}")


def convert_pdf_to_docx(
    pdf_path,
    out_docx_path=None,
    exclude_texts=None,
    remove_logos=True,
    auto_footer=True,
    master_doc=None
):
    if not out_docx_path:
        out_docx_path = os.path.splitext(pdf_path)[0] + ".docx"

    user_patterns = list(DEFAULT_EXCLUDE_TEXTS)
    if exclude_texts:
        for item in exclude_texts:
            if item and item.strip():
                user_patterns.append(item.strip())

    doc_docx = docx.Document() if master_doc is None else master_doc
    if master_doc is None:
        for s in doc_docx.sections:
            s.top_margin = Inches(1)
            s.bottom_margin = Inches(1)
            s.left_margin = Inches(1)
            s.right_margin = Inches(1)

    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    pres_titles = [
        base_name.strip(),
        re.sub(r'^(\d+\.|\w+\s*-\s*\d+\s*-\s*|[A-Za-z]+\s+\d+)\s*', '', base_name).strip()
    ]

    title_p = doc_docx.add_paragraph()
    title_p.paragraph_format.space_before = Pt(14)
    title_p.paragraph_format.space_after = Pt(8)
    t_run = title_p.add_run(base_name)
    t_run.font.name = "Arial"
    t_run.font.size = Pt(16)
    t_run.font.bold = True
    t_run.font.color.rgb = RGBColor(15, 44, 89)

    pdf_doc = fitz.open(pdf_path)

    fix_pdf_tounicode(pdf_doc)

    auto_repeated_set = detect_repeated_slide_texts(pdf_doc) if auto_footer else set()
    rep_xrefs, rep_sigs = detect_repeating_images(pdf_doc) if remove_logos else (set(), set())
    template_drawing_rects = detect_template_drawings(pdf_doc)

    total_pages = len(pdf_doc)
    images_count = 0

    for pno in range(total_pages):
        page = pdf_doc[pno]
        pw, ph = page.rect.width, page.rect.height

        slide_h = doc_docx.add_paragraph()
        slide_h.paragraph_format.space_before = Pt(16)
        slide_h.paragraph_format.space_after = Pt(4)
        slide_h.paragraph_format.keep_with_next = True
        h_run = slide_h.add_run(f"--- Slajd {pno + 1} ---")
        h_run.font.name = "Arial"
        h_run.font.size = Pt(12)
        h_run.font.bold = True
        h_run.font.color.rgb = RGBColor(31, 78, 121)

        blocks = page.get_text("blocks")
        blocks_sorted = sorted(blocks, key=lambda b: (1 if b[0] >= 0.55 * pw else 0, round(b[1], 1), round(b[0], 1)))

        for b in blocks_sorted:
            if len(b) >= 7 and b[6] != 0:
                continue
            y0, y1 = b[1], b[3]
            raw_text = b[4]
            if not raw_text.strip():
                continue

            for raw_line in raw_text.splitlines():
                cleaned_line = clean_line_characters(raw_line).rstrip()
                if not cleaned_line.strip():
                    continue

                if should_exclude_line(
                    cleaned_line, y0, y1, ph, pno,
                    user_patterns, auto_repeated_set, pres_titles
                ):
                    continue

                p = doc_docx.add_paragraph()
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.15
                r = p.add_run(cleaned_line)
                r.font.name = "Arial"
                r.font.size = Pt(10.5)
                r.font.color.rgb = RGBColor(34, 34, 34)

        raw_imgs = page.get_images()
        valid_entries = []
        for im in raw_imgs:
            xref = im[0]
            if remove_logos and xref in rep_xrefs:
                continue
            try:
                base = pdf_doc.extract_image(xref)
            except Exception:
                continue
            w, h = base['width'], base['height']
            sig = (w, h, len(base['image']))
            if remove_logos and sig in rep_sigs:
                continue

            rects = page.get_image_rects(xref)
            for r in rects:
                if is_logo_or_decoration(r, w, h, pw, ph, remove_logos=remove_logos):
                    continue
                valid_entries.append({
                    'xref': xref,
                    'w': w,
                    'h': h,
                    'rect': r,
                    'bytes': base['image'],
                    'ext': base['ext']
                })

        vector_clip = get_vector_diagram_clip(page, template_drawing_rects)
        if len(valid_entries) >= 4 and all(ve['w'] < 350 and ve['h'] < 250 for ve in valid_entries):
            cluster_rect = fitz.Rect(valid_entries[0]['rect'])
            for ve in valid_entries[1:]:
                cluster_rect |= ve['rect']
            if vector_clip:
                cluster_rect |= vector_clip
            clip_r = fitz.Rect(
                max(10, cluster_rect.x0 - 10),
                max(0.10 * ph, cluster_rect.y0 - 10),
                min(pw - 10, cluster_rect.x1 + 10),
                min(0.88 * ph, cluster_rect.y1 + 10)
            )
            pix = page.get_pixmap(clip=clip_r, dpi=180)
            insert_image(doc_docx, pix.tobytes("png"), pix.width, pix.height)
            images_count += 1
        elif vector_clip and not valid_entries:
            pix = page.get_pixmap(clip=vector_clip, dpi=180)
            insert_image(doc_docx, pix.tobytes("png"), pix.width, pix.height)
            images_count += 1
        else:
            clean_entries = deduplicate_images(valid_entries)
            clean_entries.sort(key=lambda item: item['rect'].y0)
            for entry in clean_entries:
                insert_image(doc_docx, entry['bytes'], entry['w'], entry['h'])
                images_count += 1
            if vector_clip:
                overlaps_img = any(vector_clip.intersects(e['rect']) for e in clean_entries)
                if not overlaps_img:
                    pix = page.get_pixmap(clip=vector_clip, dpi=180)
                    insert_image(doc_docx, pix.tobytes("png"), pix.width, pix.height)
                    images_count += 1

    if master_doc is None:
        doc_docx.save(out_docx_path)
        print(f"Konvertovano: {os.path.basename(pdf_path)} -> {os.path.basename(out_docx_path)} ({total_pages} slajdova, {images_count} slika)")
    return out_docx_path, total_pages, images_count


def main():
    parser = argparse.ArgumentParser(
        description="Konverter .pdf u .docx sa filtriranjem futera i logoa."
    )
    parser.add_argument("files", nargs="*", help="PDF fajlovi ili folderi za obradu (podrazumevano: trenutni folder).")
    parser.add_argument(
        "-e", "--exclude",
        nargs="+",
        default=[],
        help="Tekstovi koje treba izbaciti sa slajdova"
    )
    parser.add_argument(
        "--keep-logos",
        action="store_true",
        help="Zadrži logoe (podrazumevano se izbacuju)"
    )
    parser.add_argument(
        "--no-auto-footer",
        action="store_true",
        help="Isključi automatsku detekciju ponavljajućeg teksta u hederu/futeru."
    )
    parser.add_argument(
        "--interactive", "-i",
        action="store_true",
        help="Interaktivno pitaj za tekstove koje treba izbaciti pre pokretanja."
    )
    parser.add_argument("--all", action="store_true", help="Obradi sve PDF fajlove u folderu.")
    parser.add_argument("--master", action="store_true", help="Kreiraj i objedinjeni Word fajl.")
    args = parser.parse_args()

    exclude_list = list(DEFAULT_EXCLUDE_TEXTS)
    for item in args.exclude:
        for sub in re.split(r'[;,]', item):
            if sub.strip():
                exclude_list.append(sub.strip())

    remove_logos = DEFAULT_REMOVE_LOGOS and (not args.keep_logos)
    auto_footer = DEFAULT_AUTO_FOOTER and (not args.no_auto_footer)

    if args.interactive:
        print("=== PODEŠAVANJA KONVERZIJE PDF -> DOCX ===")
        raw_ex = input("Unesite tekstove za izbegavanje (odvojene zarezom, Enter za preskakanje): ").strip()
        if raw_ex:
            for sub in re.split(r'[;,]', raw_ex):
                if sub.strip():
                    exclude_list.append(sub.strip())
        raw_logo = input("Izbaciti logoe? [y/n]: ").strip().lower()
        if raw_logo in ('n', 'ne', 'no'):
            remove_logos = False

    target_files = []
    search_dir = "."
    if not args.files or args.all:
        pdfs = [f for f in os.listdir('.') if f.lower().endswith('.pdf')]
        target_files = pdfs
    else:
        for f in args.files:
            if os.path.isdir(f):
                search_dir = f
                for entry in os.listdir(f):
                    if entry.lower().endswith('.pdf'):
                        target_files.append(os.path.join(f, entry))
            elif os.path.exists(f) and f.lower().endswith('.pdf'):
                target_files.append(f)

    def sort_key(s):
        base = os.path.basename(s)
        m = re.search(r'(\d+)', base)
        return (int(m.group(1)) if m else 999, base.lower())

    target_files = sorted(target_files, key=sort_key)

    if not target_files:
        print("Nisu pronađeni PDF fajlovi za obradu.")
        return

    print(f"Pronađeno {len(target_files)} PDF fajlova za obradu.")
    if exclude_list:
        print(f"Izbacivanje zadatih tekstova: {exclude_list}")
    print(f"Automatsko čišćenje ponavljajućih futera/hedera: {'DA' if auto_footer else 'NE'}")
    print(f"Izbacivanje logoa i ukrasnih pozadina: {'DA' if remove_logos else 'NE'}\n")

    for f in target_files:
        convert_pdf_to_docx(
            f,
            exclude_texts=exclude_list,
            remove_logos=remove_logos,
            auto_footer=auto_footer
        )

    if args.master and len(target_files) > 1:
        print("\nKreiranje objedinjenog fajla...")
        master_doc = docx.Document()
        for s in master_doc.sections:
            s.top_margin = Inches(1)
            s.bottom_margin = Inches(1)
            s.left_margin = Inches(1)
            s.right_margin = Inches(1)
        for idx, f in enumerate(target_files):
            if idx > 0:
                master_doc.add_page_break()
            convert_pdf_to_docx(
                f,
                exclude_texts=exclude_list,
                remove_logos=remove_logos,
                auto_footer=auto_footer,
                master_doc=master_doc
            )
        m_name = os.path.join(search_dir, "Objedinjeno.docx")
        master_doc.save(m_name)
        print(f"Kreiran objedinjeni fajl: {m_name}")


if __name__ == '__main__':
    main()