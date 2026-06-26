import csv
import json
import re
from collections import Counter
from pathlib import Path

from docx import Document


BASE_DIR = Path(__file__).resolve().parents[1]
OUTPUT_DIR = BASE_DIR / "output"
DOCX_PATH = next(BASE_DIR.glob("*.docx"), None)
PDF_PATH = next(BASE_DIR.glob("*.pdf"), None)

UNIT_TERMS = [
    "Formasyonu",
    "Uyesi",
    "Üyesi",
    "Gurubu",
    "Grubu",
    "Kompleksi",
    "Volkaniti",
    "Graniti",
    "Granodiyoriti",
    "Kirectasi",
    "Kireçtaşı",
]

FALSE_PREFIXES = (
    "Sekil",
    "Şekil",
    "Cizelge",
    "Çizelge",
    "Tablo",
    "Levha",
)

SECTION_RE = re.compile(r"^(?P<section>(?:[IVX]+\.)?(?:\d+\.)+\d*\.?)\s*")
CODE_RE = re.compile(r"\((?P<code>[A-Za-z0-9ÇĞİÖŞÜçğıöşü]+)\)")

# Üye alt bölümleri kaynakta "IV.4.2.1. Çukurçeşme Üyesi (Tçç): açıklama..." gibi
# bölüm no + ad + (kod) + açıklama TEK paragrafta ve uzundur; normal başlık filtreleri
# (uzunluk > 140, nokta ile biter) bunları eler. Bu kalıp, gövdesi satır içinde başlayan
# üye başlığını yakalar (kod ZORUNLU; satır içi açıklama `rest` olarak alınır).
MEMBER_HEADING_RE = re.compile(
    r"^(?P<section>(?:[IVX]+\.)?(?:\d+\.)+\d*\.?)\s*"
    r"(?P<name>[A-ZÇĞİÖŞÜ][^()]*?(?:Üyesi|Kireçtaşı Üyesi))\s*"
    r"\((?P<code>[A-Za-z0-9ÇĞİÖŞÜçğıöşü]+)\)\s*[:\.]\s*(?P<rest>.*)$"
)

# Üye bloğunun içinde, formasyon kapsamına dönüldüğünü gösteren alt başlıklar; üye
# gövdesi bunlarda kesilir (aksi halde son üye, formasyonun kapanış bölümlerini yutar).
MEMBER_BODY_STOP_RE = re.compile(
    r"^(?:Dokanak ili[şs]kileri|Fosil [Kk]apsam[ıi] ve Ya[şs]|"
    r"Çökelme ortam[ıi]|Kal[ıi]nl[ıi]k ve Yay[ıi]l[ıi]m|Dene[şs]tirme)\b"
)
TOC_DOTS_RE = re.compile(r"\.{5,}\s*\d+\s*$")
MULTISPACE_RE = re.compile(r"[ \t]+")

FIELD_PATTERNS = {
    "age": [
        re.compile(r"(?:Fosil Kapsam[ıi] ve Ya[şs]|Ya[şs])\.\s*(.+?)(?=\n[A-ZÇĞİÖŞÜ][^\n]{0,80}\.|$)", re.S),
    ],
    "lithology": [
        re.compile(r"(?:Kaya ?t[üu]r[üu] [ÖO]zellikleri|Kaya t[üu]r[üu]|Litoloji)\.\s*(.+?)(?=\n[A-ZÇĞİÖŞÜ][^\n]{0,80}\.|$)", re.S),
    ],
    "thickness": [
        re.compile(r"(?:Dokanak ili[şs]kileri ve Kal[ıi]nl[ıi]k|Kal[ıi]nl[ıi]k)\.\s*(.+?)(?=\n[A-ZÇĞİÖŞÜ][^\n]{0,80}\.|$)", re.S),
    ],
    "contacts": [
        re.compile(r"(?:Dokanak ili[şs]kileri(?: ve Kal[ıi]nl[ıi]k)?)\.\s*(.+?)(?=\n[A-ZÇĞİÖŞÜ][^\n]{0,80}\.|$)", re.S),
    ],
    "distribution": [
        re.compile(r"(?:Da[ğg][ıi]l[ıi]m[ıi]|Yay[ıi]l[ıi]m[ıi]|Y[üu]zeyleme)\.\s*(.+?)(?=\n[A-ZÇĞİÖŞÜ][^\n]{0,80}\.|$)", re.S),
    ],
    "fossils": [
        re.compile(r"(?:Fosil Kapsam[ıi])\.\s*(.+?)(?=\n[A-ZÇĞİÖŞÜ][^\n]{0,80}\.|$)", re.S),
    ],
    "correlation": [
        re.compile(r"(?:Dene[şs]tirme|Korelasyon)\.\s*(.+?)(?=\n[A-ZÇĞİÖŞÜ][^\n]{0,80}\.|$)", re.S),
    ],
}

KNOWN_CODES = {
    "Kapaklı Formasyonu": "PTRk",
    "Bakırlıkıran Formasyonu": "TRgbkr",
}

SUBHEADING_PATTERNS = [
    ("definition", re.compile(r"^(Tanım ve [Aa]d)\.\s*(.*)$")),
    ("type_location", re.compile(r"^(Tip yer(?:, Başvuru yeri)?|Başvuru yeri)\.\s*(.*)$")),
    ("lithology", re.compile(r"^(Kayatürü özellikleri|Kaya türü özellikleri|Litoloji)\.\s*(.*)$")),
    ("contacts_thickness", re.compile(r"^(Dokanak ilişkileri ve Kalınlık)\.\s*(.*)$")),
    ("contacts", re.compile(r"^(Dokanak ilişkileri)\.\s*(.*)$")),
    ("fossils_age", re.compile(r"^(Fosil kapsamı ve yaş|Fosil Kapsamı ve Yaş)\.\s*(.*)$")),
    ("depositional_environment", re.compile(r"^(Çökelme ortamı)\.\s*(.*)$")),
    ("correlation", re.compile(r"^(Deneştirme|Korelasyon)\.\s*(.*)$")),
    ("distribution", re.compile(r"^(Dağılımı|Yayılımı|Yüzeyleme)\.\s*(.*)$")),
]


CHAR_TRANSLATION = str.maketrans(
    {
        "�": "",
        "’": "'",
        "‘": "'",
        "“": '"',
        "”": '"',
        "\u00a0": " ",
    }
)


COMMON_FIXES = {
    "STANBUL": "İSTANBUL",
    "stanbul": "İstanbul",
    "Jeolojisi": "Jeolojisi",
    "Knalada": "Kınalıada",
    "Kinaliada": "Kınalıada",
    "Kocatngel": "Kocatöngel",
    "Kurtky": "Kurtköy",
    "Denizli Ky": "Denizli Köyü",
    "Polonezky": "Polonezköy",
    "Gurubu": "Grubu",
    "kirectasi": "kireçtaşı",
    "Kirectasi": "Kireçtaşı",
    "eyil": "şeyl",
    "cakltasi": "çakıltaşı",
}


def normalize_text(text):
    clean = text.translate(CHAR_TRANSLATION)
    clean = MULTISPACE_RE.sub(" ", clean)
    clean = "\n".join(line.strip() for line in clean.splitlines())
    clean = clean.strip()
    for wrong, right in COMMON_FIXES.items():
        clean = clean.replace(wrong, right)
    return clean


def compact_text(text, limit=900):
    text = re.sub(r"\s+", " ", text).strip()
    return text[:limit].rstrip()


def normalize_for_lookup(text):
    return re.sub(r"\s+", " ", text or "").strip()


def read_docx_paragraphs(path):
    doc = Document(str(path))
    paragraphs = []
    for idx, para in enumerate(doc.paragraphs):
        raw = para.text.strip()
        if not raw:
            continue
        paragraphs.append(
            {
                "index": idx,
                "raw": raw,
                "clean": normalize_text(raw),
                "style": para.style.name if para.style else None,
            }
        )
    return paragraphs


def looks_like_front_matter_toc(text):
    return bool(TOC_DOTS_RE.search(text)) or text.count(".") > 12


def has_unit_term(text):
    folded = text.replace("�", "")
    return any(term in folded for term in UNIT_TERMS)


def is_false_positive(text):
    stripped = text.strip()
    if not stripped:
        return True
    if stripped.startswith(FALSE_PREFIXES):
        return True
    if looks_like_front_matter_toc(stripped):
        return True
    if len(stripped) > 140:
        return True
    if "Formasyonunun" in stripped or "Formasyonu'nun" in stripped:
        return True
    if stripped.endswith(".") and not SECTION_RE.match(stripped) and not CODE_RE.search(stripped):
        return True
    return False


def parse_heading(raw_text, clean_text):
    text = clean_text
    section = None
    section_match = SECTION_RE.match(text)
    if section_match:
        section = section_match.group("section").rstrip(".")
        text = text[section_match.end() :].strip()

    code = None
    code_match = CODE_RE.search(text)
    if code_match:
        code = code_match.group("code")
        text = text[: code_match.start()].strip()

    text = re.sub(r"\s+", " ", text)
    raw_no_page = re.sub(r"\s+\d+$", "", text).strip()
    name = raw_no_page.strip(" .")

    if not name:
        return None

    return {
        "section_number": section,
        "name": name,
        "code": code,
        "raw_heading": raw_text,
        "clean_heading": clean_text,
    }


def classify_rank(name):
    if "Üyesi" in name or "Uyesi" in name:
        return "üye"
    if "Formasyonu" in name:
        return "formasyon"
    if "Grubu" in name or "Gurubu" in name:
        return "grup"
    if "Kompleksi" in name:
        return "kompleks"
    if "Volkaniti" in name:
        return "volkanit"
    if "Graniti" in name or "Granodiyoriti" in name:
        return "magmatik_birim"
    if "Kireçtaşı" in name or "Kirectasi" in name:
        return "kireçtaşı"
    return "diğer"


def confidence_for(header, body):
    score = 0
    if header["section_number"]:
        score += 2
    if header["code"]:
        score += 2
    if classify_rank(header["name"]) != "diğer":
        score += 1
    if len(body) > 700:
        score += 1
    if "�" in header["raw_heading"]:
        score -= 1
    if score >= 5:
        return "high"
    if score >= 3:
        return "medium"
    return "low"


def extract_field(block, field_name):
    for pattern in FIELD_PATTERNS[field_name]:
        match = pattern.search(block)
        if match:
            return compact_text(match.group(1), 700)
    return None


def split_subheading(text):
    for key, pattern in SUBHEADING_PATTERNS:
        match = pattern.match(text)
        if match:
            return key, match.group(1), match.group(2).strip()
    return None, None, text


def extract_sections(body_paras):
    sections = {}
    current_key = "definition"
    current_lines = []

    def flush():
        if current_lines:
            sections.setdefault(current_key, [])
            sections[current_key].extend(current_lines)

    for para in body_paras:
        text = para["clean"].strip()
        if not text:
            continue
        key, _label, remainder = split_subheading(text)
        if key:
            flush()
            current_key = key
            current_lines = [remainder] if remainder else []
        else:
            current_lines.append(text)
    flush()

    return {key: compact_text("\n".join(lines), 1200) for key, lines in sections.items() if compact_text("\n".join(lines), 1200)}


def pick_field_from_sections(sections, field_name, clean_block):
    if field_name == "lithology":
        return sections.get("lithology") or sections.get("definition") or extract_field(clean_block, field_name)
    if field_name == "contacts":
        return sections.get("contacts_thickness") or sections.get("contacts") or extract_field(clean_block, field_name)
    if field_name == "thickness":
        return sections.get("contacts_thickness") or extract_field(clean_block, field_name)
    if field_name == "age":
        return sections.get("fossils_age") or extract_field(clean_block, field_name)
    if field_name == "fossils":
        return sections.get("fossils_age") or extract_field(clean_block, field_name)
    if field_name == "correlation":
        return sections.get("correlation") or extract_field(clean_block, field_name)
    if field_name == "distribution":
        return sections.get("distribution") or infer_distribution(sections.get("definition", ""))
    return extract_field(clean_block, field_name)


def infer_distribution(definition_text):
    if not definition_text:
        return None
    sentences = re.split(r"(?<=[.!?])\s+", definition_text)
    keywords = ("yüzey", "alan", "kaplar", "yaygın", "dolay", "köy", "semt", "mahalle", "İstanbul")
    selected = [sentence for sentence in sentences if any(keyword in sentence for keyword in keywords)]
    if not selected:
        return None
    return compact_text(" ".join(selected[:3]), 700)


def find_candidate_headers(paragraphs):
    candidates = []
    filtered = Counter()
    body_started = False

    for pos, para in enumerate(paragraphs):
        raw = para["raw"]
        clean = para["clean"]

        if clean.startswith("IV.1.1.") or clean.startswith("IV.1.2."):
            body_started = True

        # Üye alt bölümü: satır içi başlık+gövde; normal uzunluk/nokta filtrelerini atlar.
        if body_started and not clean.startswith(FALSE_PREFIXES) \
                and not looks_like_front_matter_toc(clean):
            member = MEMBER_HEADING_RE.match(clean)
            if member:
                name = re.sub(r"\s+", " ", member.group("name").strip(" ."))
                candidates.append({
                    "section_number": member.group("section").rstrip("."),
                    "name": name,
                    "code": member.group("code"),
                    "raw_heading": raw,
                    "clean_heading": clean,
                    "inline_body": member.group("rest").strip(),
                    "paragraph_pos": pos,
                    "docx_paragraph_start": para["index"],
                })
                continue

        if not has_unit_term(raw) and not has_unit_term(clean):
            continue

        if not body_started:
            filtered["front_matter_or_toc"] += 1
            continue

        if is_false_positive(clean):
            if clean.startswith(FALSE_PREFIXES):
                filtered["figure_or_table"] += 1
            elif looks_like_front_matter_toc(clean):
                filtered["toc_dots"] += 1
            else:
                filtered["paragraph_or_long_line"] += 1
            continue

        parsed = parse_heading(raw, clean)
        if not parsed:
            filtered["parse_failed"] += 1
            continue
        if not parsed["section_number"] and not parsed["code"]:
            filtered["no_section_or_code"] += 1
            continue
        if classify_rank(parsed["name"]) == "diğer":
            filtered["unclassified_heading"] += 1
            continue

        parsed["paragraph_pos"] = pos
        parsed["docx_paragraph_start"] = para["index"]
        candidates.append(parsed)

    return candidates, filtered


def parent_for(header, stack):
    section = header["section_number"]
    if not section:
        return None
    parts = section.split(".")
    if len(parts) <= 2:
        return None
    parent_section = ".".join(parts[:-1])
    return stack.get(parent_section)


def build_units(paragraphs, headers):
    units = []
    section_stack = {}

    for i, header in enumerate(headers):
        start_pos = header["paragraph_pos"]
        end_pos = headers[i + 1]["paragraph_pos"] if i + 1 < len(headers) else len(paragraphs)
        body_paras = paragraphs[start_pos + 1 : end_pos]

        inline = header.get("inline_body")
        if inline:
            # Üye gövdesi başlık paragrafında (kod ":" sonrası) başlar; öne ekle.
            body_paras = [{
                "index": header["docx_paragraph_start"],
                "raw": inline,
                "clean": normalize_text(inline),
                "style": None,
            }] + body_paras
            # Formasyon kapanış alt başlığında (Dokanak/Fosil/Çökelme...) üye gövdesini kes.
            for cut, bp in enumerate(body_paras):
                if cut > 0 and MEMBER_BODY_STOP_RE.match(bp["clean"].strip()):
                    body_paras = body_paras[:cut]
                    break

        raw_block = "\n".join(p["raw"] for p in body_paras)
        clean_block = "\n".join(p["clean"] for p in body_paras)
        sections = extract_sections(body_paras)

        rank = classify_rank(header["name"])
        parent_unit = parent_for(header, section_stack)
        unit_id = f"unit_{i + 1:04d}"
        code = header["code"] or KNOWN_CODES.get(header["name"])
        available_sections = sorted(sections.keys())
        expected_sections = ["lithology", "contacts_thickness", "fossils_age", "correlation"]
        missing_sections = [section for section in expected_sections if section not in sections]
        integration_ready = bool(code and (sections.get("lithology") or sections.get("definition")) and (sections.get("contacts_thickness") or sections.get("contacts")) and sections.get("fossils_age"))

        if header["section_number"]:
            section_stack[header["section_number"]] = unit_id

        units.append(
            {
                "id": unit_id,
                "name": header["name"],
                "code": code,
                "rank": rank,
                "parent_unit": parent_unit,
                "section_number": header["section_number"],
                "docx_paragraph_start": header["docx_paragraph_start"],
                "docx_paragraph_end": paragraphs[end_pos - 1]["index"] if end_pos > start_pos else header["docx_paragraph_start"],
                "source_pages": [],
                "age": pick_field_from_sections(sections, "age", clean_block),
                "lithology": pick_field_from_sections(sections, "lithology", clean_block),
                "thickness": pick_field_from_sections(sections, "thickness", clean_block),
                "contacts": pick_field_from_sections(sections, "contacts", clean_block),
                "distribution": pick_field_from_sections(sections, "distribution", clean_block),
                "fossils": pick_field_from_sections(sections, "fossils", clean_block),
                "correlation": pick_field_from_sections(sections, "correlation", clean_block),
                "available_sections": available_sections,
                "missing_sections": missing_sections,
                "integration_ready": integration_ready,
                "raw_heading": header["raw_heading"],
                "clean_heading": header["clean_heading"],
                "raw_text_excerpt": compact_text(raw_block),
                "clean_text_excerpt": compact_text(clean_block),
                "extraction_confidence": confidence_for(header, clean_block),
            }
        )

    return units


def write_json(path, units):
    with path.open("w", encoding="utf-8") as f:
        json.dump(units, f, ensure_ascii=False, indent=2)


def write_csv(path, units):
    if not units:
        return
    fieldnames = list(units[0].keys())
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for unit in units:
            row = dict(unit)
            for key, value in row.items():
                if isinstance(value, (list, dict)):
                    row[key] = json.dumps(value, ensure_ascii=False)
            writer.writerow(row)


def report_markdown(paragraph_count, units, filtered):
    rank_counts = Counter(unit["rank"] for unit in units)
    confidence_counts = Counter(unit["extraction_confidence"] for unit in units)
    tracked_fields = ["code", "lithology", "distribution", "fossils", "contacts", "thickness", "correlation", "age", "parent_unit"]
    field_counts = {field: sum(1 for unit in units if unit.get(field)) for field in tracked_fields}
    integration_ready = [u for u in units if u.get("integration_ready")]
    missing_section_counts = Counter(section for unit in units for section in unit.get("missing_sections", []))
    no_code = [u for u in units if not u["code"]]
    short_blocks = [u for u in units if len(u["clean_text_excerpt"] or "") < 180]
    bad_chars = [u for u in units if "�" in u["raw_heading"] or "�" in (u["raw_text_excerpt"] or "")]
    low_conf = [u for u in units if u["extraction_confidence"] == "low"]
    missing_core = [
        u for u in units
        if not u.get("code") or not u.get("lithology") or not u.get("contacts") or not u.get("fossils")
    ]

    def unit_line(unit):
        return f"- {unit['id']}: {unit['name']} ({unit['code'] or 'kod yok'}) - {unit['rank']} - p:{unit['docx_paragraph_start']}"

    lines = [
        "# Jeoloji Birimleri Çıkarım Raporu",
        "",
        "## Kaynaklar",
        f"- DOCX: {DOCX_PATH.name if DOCX_PATH else 'bulunamadı'}",
        f"- PDF: {PDF_PATH.name if PDF_PATH else 'bulunamadı'}",
        "",
        "## Özet",
        f"- Toplam metin paragrafı: {paragraph_count}",
        f"- Yakalanan birim sayısı: {len(units)}",
        f"- Rank dağılımı: {dict(rank_counts)}",
        f"- Güven dağılımı: {dict(confidence_counts)}",
        f"- Filtrelenen adaylar: {dict(filtered)}",
        f"- Entegrasyona hazır kayıt sayısı: {len(integration_ready)}",
        "",
        "## Alan Doluluk Özeti",
    ]
    lines.extend(f"- {field}: {count}/{len(units)}" for field, count in field_counts.items())
    lines.extend(
        [
            "",
            "## Alt Başlık Eksikliği Özeti",
        ]
    )
    lines.extend(f"- {section}: {count}/{len(units)}" for section, count in sorted(missing_section_counts.items()))
    lines.extend(
        [
            "",
        "## Karakter Düzeltme Kontrolü",
        "- DOCX dönüşümünde Türkçe karakterlerin bir kısmı ham metinde bozuk kalmış durumda.",
        "- Betik ham başlığı ve temiz başlığı birlikte saklar; mühendislik kullanımı öncesinde adlar manuel örneklemle kontrol edilmelidir.",
        "- Özellikle `�` karakteri içeren kayıtlar aşağıdaki manuel kontrol listesinde işaretlenmiştir.",
        "",
        "## İlk 20 Birim",
    ]
    )
    lines.extend(unit_line(unit) for unit in units[:20])

    lines.extend(
        [
            "",
            "## Manuel Kontrol Önerilen Kayıtlar",
            "",
            "### Kodu Olmayan Başlıklar",
        ]
    )
    lines.extend(unit_line(unit) for unit in no_code[:30])
    if not no_code:
        lines.append("- Yok")

    lines.extend(["", "### Çok Kısa Açıklama Bloğu Olanlar"])
    lines.extend(unit_line(unit) for unit in short_blocks[:30])
    if not short_blocks:
        lines.append("- Yok")

    lines.extend(["", "### Düşük Güvenli Kayıtlar"])
    lines.extend(unit_line(unit) for unit in low_conf[:30])
    if not low_conf:
        lines.append("- Yok")

    lines.extend(["", "### Eksik Temel Alan Nedeniyle Kontrol Gerektirenler"])
    lines.extend(unit_line(unit) for unit in missing_core[:40])
    if not missing_core:
        lines.append("- Yok")

    lines.extend(["", "### Ham Metninde Bozuk Karakter Bulunanlar"])
    lines.extend(unit_line(unit) for unit in bad_chars[:30])
    if not bad_chars:
        lines.append("- Yok")

    lines.extend(["", "## Entegrasyona Hazır İlk Kayıtlar"])
    lines.extend(unit_line(unit) for unit in integration_ready[:30])
    if not integration_ready:
        lines.append("- Yok")

    lines.extend(
        [
            "",
            "## Test Sonuçları",
            f"- JSON/CSV üretimi: {'başarılı' if units else 'başarısız'}",
            f"- JSON/CSV kayıt sayısı: {len(units)}",
            "- `Şekil` ve `Çizelge` ile başlayan adaylar katalogdan filtrelendi.",
            "- İçindekilerdeki noktalı sayfa satırları katalogdan filtrelendi.",
            f"- Litoloji doluluğu: {field_counts['lithology']}/{len(units)}",
            f"- Yayılım/dağılım doluluğu: {field_counts['distribution']}/{len(units)}",
            f"- Fosil/yaş doluluğu: {field_counts['fossils']}/{len(units)}",
            "- `source_pages` ilk sürümde boş bırakıldı; kesin izlenebilirlik DOCX paragraf indeksleriyle sağlanır.",
            "",
            "## Sonraki Adımlar",
            "1. `geology_units.csv` dosyasında manuel kontrol listesine giren kayıtları gözden geçirin.",
            "2. Karakter düzeltme sözlüğünü gerçek jeolojik adlara göre genişletin.",
            "3. Gerekirse PDF sayfa numaralarıyla DOCX paragraf aralıklarını eşleştiren ikinci bir doğrulama adımı ekleyin.",
        ]
    )
    return "\n".join(lines) + "\n"


def main():
    if DOCX_PATH is None:
        raise FileNotFoundError(f"DOCX bulunamadı: {BASE_DIR}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    paragraphs = read_docx_paragraphs(DOCX_PATH)
    headers, filtered = find_candidate_headers(paragraphs)
    units = build_units(paragraphs, headers)

    write_json(OUTPUT_DIR / "geology_units.json", units)
    write_csv(OUTPUT_DIR / "geology_units.csv", units)
    (OUTPUT_DIR / "extraction_report.md").write_text(
        report_markdown(len(paragraphs), units, filtered),
        encoding="utf-8",
    )

    print(f"DOCX: {DOCX_PATH.name}")
    print(f"Paragraphs: {len(paragraphs)}")
    print(f"Units: {len(units)}")
    print(f"Rank counts: {dict(Counter(unit['rank'] for unit in units))}")
    print(f"Confidence counts: {dict(Counter(unit['extraction_confidence'] for unit in units))}")
    print(f"Output: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
