import argparse
import json
import re
import sys
import unicodedata
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parents[1]
DEFAULT_DATA_PATH = BASE_DIR / "output" / "geology_units.json"

DISPLAY_FIELDS = [
    ("lithology", "Litoloji"),
    ("distribution", "Yayılım"),
    ("contacts", "Dokanak İlişkileri"),
    ("thickness", "Kalınlık"),
    ("fossils", "Fosil Kapsamı"),
    ("age", "Yaş"),
    ("correlation", "Korelasyon"),
]

MISSING_MESSAGES = {
    "lithology": "Bu birim için litoloji bilgisi otomatik çıkarılamadı.",
    "distribution": "Bu birim için yayılım bilgisi otomatik çıkarılamadı.",
    "contacts": "Dokanak ilişkileri kaynak blokta ayrı başlık olarak bulunamadı.",
    "thickness": "Kalınlık bilgisi kaynak blokta ayrı başlık olarak bulunamadı.",
    "fossils": "Bu birim için fosil kapsamı bölümü otomatik çıkarılamadı.",
    "age": "Bu birim için yaş bölümü otomatik çıkarılamadı.",
    "correlation": "Bu birim için korelasyon/deneştirme bölümü otomatik çıkarılamadı.",
}


def load_units(path=DEFAULT_DATA_PATH):
    with Path(path).open("r", encoding="utf-8") as f:
        return json.load(f)


def normalize_lookup(text):
    text = clean_display_text(text or "").casefold()
    replacements = str.maketrans(
        {
            "ı": "i",
            "İ": "i",
            "ö": "o",
            "Ö": "o",
            "ü": "u",
            "Ü": "u",
            "ğ": "g",
            "Ğ": "g",
            "ş": "s",
            "Ş": "s",
            "ç": "c",
            "Ç": "c",
        }
    )
    text = text.translate(replacements)
    text = "".join(ch for ch in unicodedata.normalize("NFKD", text) if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", text).strip()


def clean_display_text(text):
    if text is None:
        return ""
    replacements = {
        "İİstanbul": "İstanbul",
        "��stanbul": "İstanbul",
        "Formsyonu": "Formasyonu",
        "Ortdovisiyen": "Ordovisiyen",
        "belirtmişitir": "belirtmiştir",
        "şşeyl": "şeyl",
        "Şşeyl": "Şeyl",
        "��eyl": "şeyl",
        "�eyl": "şeyl",
        "��": "",
        "�": "",
        "\u00a0": " ",
        "’": "'",
        "‘": "'",
        "“": '"',
        "”": '"',
    }
    cleaned = str(text)
    for wrong, right in replacements.items():
        cleaned = cleaned.replace(wrong, right)
    cleaned = re.sub(r"\(Şekil\s*(\d+)\)\s*Şekil\s*\1\.", r"(Şekil \1).", cleaned)
    cleaned = re.sub(r"\bŞekil\s*\d+\.\s*[^.]{0,180}\.\s*", "", cleaned)
    cleaned = re.sub(r"[ \t]+", " ", cleaned)
    cleaned = re.sub(r"\s+\n", "\n", cleaned)
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    return cleaned.strip()


def split_paragraphs(text, max_chars=650):
    text = clean_display_text(text)
    if not text:
        return []
    sentences = re.split(r"(?<=[.!?])\s+", text)
    paragraphs = []
    current = ""
    for sentence in sentences:
        sentence = sentence.strip()
        if not sentence:
            continue
        if current and len(current) + 1 + len(sentence) > max_chars:
            paragraphs.append(current)
            current = sentence
        else:
            current = f"{current} {sentence}".strip()
    if current:
        paragraphs.append(current)
    return paragraphs


def truncate_text(text, limit):
    text = clean_display_text(text)
    if len(text) <= limit:
        return text
    cut = text[:limit].rsplit(" ", 1)[0].rstrip(" ,;:")
    return f"{cut}..."


def field_value(unit, field, mode):
    value = clean_display_text(unit.get(field))
    if not value:
        return None
    if mode == "short":
        return truncate_text(value, 700)
    return value


def same_text(left, right):
    return normalize_lookup(left) == normalize_lookup(right)


def find_unit_by_name(name, units=None, data_path=DEFAULT_DATA_PATH):
    units = units if units is not None else load_units(data_path)
    query = normalize_lookup(name)
    code_matches = [unit for unit in units if normalize_lookup(unit.get("code")) == query]
    if code_matches:
        return code_matches[0]
    exact = [unit for unit in units if normalize_lookup(unit.get("name")) == query]
    if exact:
        return exact[0]
    if "?" in query:
        wildcard = re.compile("^" + re.escape(query).replace(r"\?", ".") + "$")
        wildcard_matches = [unit for unit in units if wildcard.match(normalize_lookup(unit.get("name")))]
        if wildcard_matches:
            return wildcard_matches[0]
    contains = [unit for unit in units if query in normalize_lookup(unit.get("name"))]
    if len(contains) == 1:
        return contains[0]
    if contains:
        return contains[0]
    return None


def format_missing_fields(unit):
    missing = []
    for field, _label in DISPLAY_FIELDS:
        if not clean_display_text(unit.get(field)):
            missing.append(MISSING_MESSAGES[field])
    if unit.get("missing_sections"):
        sections = ", ".join(unit["missing_sections"])
        missing.append(f"Kaynak blokta eksik görünen alt başlıklar: {sections}.")
    if unit.get("extraction_confidence") != "high":
        missing.append(f"Çıkarım güven seviyesi: {unit.get('extraction_confidence')}.")
    if not unit.get("integration_ready"):
        missing.append("Bu kayıt otomatik kalite ölçütlerine göre entegrasyona tam hazır işaretlenmemiştir.")
    return missing


def add_section(lines, title, text, mode):
    value = clean_display_text(text)
    if not value:
        return
    lines.append(f"## {title}")
    for paragraph in split_paragraphs(value, max_chars=650 if mode == "short" else 900):
        lines.append(paragraph)
    lines.append("")


def format_unit_for_display(unit, mode="short"):
    if mode not in {"short", "full"}:
        raise ValueError("mode must be 'short' or 'full'")

    name = clean_display_text(unit.get("name"))
    code = clean_display_text(unit.get("code")) or "kod yok"
    rank = clean_display_text(unit.get("rank")) or "birim"

    lines = [
        f"# {name} ({code})",
        "",
        f"**Birim türü:** {rank}",
        f"**Çıkarım güveni:** {unit.get('extraction_confidence', 'bilinmiyor')}",
        f"**Entegrasyon durumu:** {'hazır' if unit.get('integration_ready') else 'kontrol gerekli'}",
        "",
    ]

    summary = unit.get("distribution") or unit.get("lithology") or unit.get("clean_text_excerpt")
    add_section(lines, "Kısa Özet", field_value({"summary": summary}, "summary", "short"), "short")

    add_section(lines, "Litoloji", field_value(unit, "lithology", mode), mode)
    add_section(lines, "Yayılım", field_value(unit, "distribution", mode), mode)

    contacts = field_value(unit, "contacts", mode)
    thickness = field_value(unit, "thickness", mode)
    if contacts and thickness and same_text(contacts, thickness):
        add_section(lines, "Dokanak ve Kalınlık", contacts, mode)
    else:
        add_section(lines, "Dokanak İlişkileri", contacts, mode)
        add_section(lines, "Kalınlık", thickness, mode)

    fossils = field_value(unit, "fossils", mode)
    age = field_value(unit, "age", mode)
    if fossils and age and same_text(fossils, age):
        add_section(lines, "Fosil Kapsamı ve Yaş", fossils, mode)
    else:
        add_section(lines, "Fosil Kapsamı", fossils, mode)
        add_section(lines, "Yaş", age, mode)

    add_section(lines, "Korelasyon", field_value(unit, "correlation", mode), mode)

    lines.extend(
        [
            "## Kaynak İzleri",
            f"- DOCX paragraf aralığı: {unit.get('docx_paragraph_start')} - {unit.get('docx_paragraph_end')}",
            f"- Ham başlık: {clean_display_text(unit.get('raw_heading'))}",
            f"- Bulunan alt başlıklar: {', '.join(unit.get('available_sections') or []) or 'yok'}",
            "",
        ]
    )

    missing = format_missing_fields(unit)
    if missing:
        lines.append("## Eksik / Kalite Notları")
        lines.extend(f"- {message}" for message in missing)
        lines.append("")

    return "\n".join(lines).strip() + "\n"


def main(argv=None):
    parser = argparse.ArgumentParser(description="Jeoloji birimi için düzenli gösterim metni üretir.")
    parser.add_argument("name", help="Formasyon/birim adı")
    parser.add_argument("--mode", choices=["short", "full"], default="short", help="Gösterim uzunluğu")
    parser.add_argument("--data", default=str(DEFAULT_DATA_PATH), help="geology_units.json yolu")
    args = parser.parse_args(argv)

    units = load_units(args.data)
    unit = find_unit_by_name(args.name, units=units)
    if not unit:
        print(f"Birim bulunamadı: {args.name}", file=sys.stderr)
        return 1
    sys.stdout.reconfigure(encoding="utf-8")
    print(format_unit_for_display(unit, mode=args.mode), end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
