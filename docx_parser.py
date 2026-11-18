"""Word (.docx) jadvalidan test savollarini o'qish.

Jadval ustunlari: savol | to'g'ri javob | noto'g'ri javob x3 | [muallif].
Formulalar LaTeX'ga o'girilib, matn ichiga \\( ... \\) ko'rinishida qo'yiladi.
"""

from __future__ import annotations

import base64
import io
import posixpath
import re
import struct
import xml.etree.ElementTree as ET
import zipfile

from PIL import Image

from mathconv import M, MtefError, OmmlConverter, ole_equation_to_latex

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
V = "{urn:schemas-microsoft-com:vml}"
O = "{urn:schemas-microsoft-com:office:office}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

# O'girib bo'lmagan formula o'rniga qo'yiladigan belgi (ko'rib chiqish sahifasida qizil chiqadi).
UNKNOWN_FORMULA = "[FORMULA?]"

MAX_IMAGE_BYTES = 2 * 1024 * 1024
MAX_IMAGE_SIDE = 1600
SLOTS = ("question", "correct", "wrong1", "wrong2", "wrong3")
SLOT_NAMES = {
    "question": "savol", "correct": "to'g'ri javob",
    "wrong1": "1-noto'g'ri javob", "wrong2": "2-noto'g'ri javob", "wrong3": "3-noto'g'ri javob",
}

_TRAILING_BASE = re.compile(r"([0-9A-Za-z]+|\))$")


class DocxError(ValueError):
    """Fayl Word hujjati emas yoki buzilgan."""


class _Cell:
    """Bitta katak mazmuni: matn (formulalar bilan) va rasmlar."""

    def __init__(self) -> None:
        self.parts: list[str] = []
        self.image_refs: list[tuple[str, dict | None]] = []
        self.warnings: list[str] = []
        self._script: tuple[str, str] | None = None  # (kind, text)

    # Yuqori/quyi indeksli ketma-ket bo'laklar bitta formulaga yig'iladi.
    def add_script(self, kind: str, text: str) -> None:
        if self._script and self._script[0] == kind:
            self._script = (kind, self._script[1] + text)
        else:
            self.flush()
            self._script = (kind, text)

    def flush(self) -> None:
        if not self._script:
            return
        kind, text = self._script
        self._script = None
        text = text.replace(" ", " ").strip()
        if not text:
            return
        base = ""
        if self.parts and not self.parts[-1].endswith(r"\)"):
            match = _TRAILING_BASE.search(self.parts[-1])
            if match:
                base = match.group(1)
                self.parts[-1] = self.parts[-1][: match.start()]
        mark = "^" if kind == "sup" else "_"
        self.parts.append(r"\(%s%s{%s}\)" % (_tex_plain(base), mark, _tex_plain(text)))

    def add_text(self, text: str) -> None:
        self.flush()
        self.parts.append(text)

    def add_math(self, latex: str) -> None:
        self.flush()
        if latex:
            self.parts.append(r"\(" + latex + r"\)")

    def text(self) -> str:
        self.flush()
        text = "".join(self.parts).replace(" ", " ")
        text = text.replace(r"\)\(", " ")
        lines = [re.sub(r"[ \t]+", " ", line).strip() for line in text.split("\n")]
        return "\n".join(line for line in lines if line)


def _tex_plain(text: str) -> str:
    """Oddiy matnni (indeks ichidagi) LaTeX uchun xavfsiz qiladi."""
    out = []
    for ch in text:
        if ch in "%#&$_{}":
            out.append("\\" + ch)
        elif ch == "\\":
            out.append(r"\backslash ")
        else:
            out.append(ch)
    return "".join(out)


class DocxQuestionParser:
    def __init__(self, content: bytes) -> None:
        try:
            self.zip = zipfile.ZipFile(io.BytesIO(content))
            document = self.zip.read("word/document.xml")
        except (zipfile.BadZipFile, KeyError) as error:
            raise DocxError("Fayl Word (.docx) hujjati emas yoki buzilgan.") from error

        try:
            self.root = ET.fromstring(document)
        except ET.ParseError as error:
            raise DocxError("Word hujjati ichini o'qib bo'lmadi.") from error

        self.rels = self._read_rels()
        self.omml = OmmlConverter()
        self.warnings: list[str] = []
        self._ole_cache: dict[str, str | None] = {}

    def _read_rels(self) -> dict[str, str]:
        try:
            root = ET.fromstring(self.zip.read("word/_rels/document.xml.rels"))
        except (KeyError, ET.ParseError):
            return {}
        rels = {}
        for item in root.findall(REL + "Relationship"):
            if item.get("TargetMode") == "External":
                continue
            target = item.get("Target", "")
            path = target.lstrip("/") if target.startswith("/") else posixpath.normpath("word/" + target)
            rels[item.get("Id", "")] = path
        return rels

    # ------------------------------------------------------------------ jadval

    def parse(self) -> dict:
        body = self.root.find(W + "body")
        tables = body.findall(W + "tbl") if body is not None else []
        if not tables:
            raise DocxError("Hujjatda jadval topilmadi. Savollar jadval ichida bo'lishi kerak.")

        questions: list[dict] = []
        author = ""
        short_rows: list[int] = []
        number = 0

        for table in tables:
            for row in table.findall(W + "tr"):
                number += 1
                cells = row.findall(W + "tc")
                if len(cells) < 5:
                    if any(self._plain_text(cell) for cell in cells):
                        short_rows.append(number)
                    continue

                parsed = [self._cell(cell) for cell in cells[:5]]
                texts = [cell.text() for cell in parsed]

                if not questions and _is_header(texts):
                    continue
                if not any(texts) and not any(cell.image_refs for cell in parsed):
                    continue

                if len(cells) > 5 and not author:
                    author = self._plain_text(cells[5])[:255]

                questions.append(self._question(number, parsed, texts))

        if short_rows:
            self.warnings.append(
                "5 tadan kam ustunli qatorlar o'tkazib yuborildi: "
                + ", ".join(str(n) for n in short_rows[:15])
                + ("..." if len(short_rows) > 15 else "")
            )
        if self.omml.unknown:
            self.warnings.append(
                "Tanilmagan formula elementlari: " + ", ".join(sorted(self.omml.unknown))
                + ". Shu formulalarni ko'zdan kechiring."
            )

        return {"author": author, "questions": questions, "warnings": self.warnings}

    def _question(self, number: int, cells: list[_Cell], texts: list[str]) -> dict:
        warnings: list[str] = []
        errors: list[str] = []
        item: dict = {"row": number}

        for slot, cell, text in zip(SLOTS, cells, texts):
            image = self._image(cell, warnings, SLOT_NAMES[slot])
            item[slot] = {"text": text, "image": image}
            for note in cell.warnings:
                warnings.append(f"{SLOT_NAMES[slot]}: {note}")
            if UNKNOWN_FORMULA in text:
                errors.append(f"{SLOT_NAMES[slot]}da o'girilmagan formula bor - qo'lda yozing.")

        def empty(slot: str) -> bool:
            return not item[slot]["text"] and not item[slot]["image"]

        if empty("question"):
            errors.append("Savol matni bo'sh.")
        if empty("correct"):
            errors.append("To'g'ri javob bo'sh.")
        wrong = [slot for slot in SLOTS[2:] if not empty(slot)]
        if not wrong:
            errors.append("Noto'g'ri javoblar yo'q.")
        elif len(wrong) < 3:
            warnings.append(f"Faqat {len(wrong)} ta noto'g'ri javob bor.")

        seen: dict[tuple, str] = {}
        for slot in SLOTS[1:]:
            if empty(slot):
                continue
            key = (_normalize(item[slot]["text"]), item[slot]["image"])
            if key in seen:
                warnings.append(f"{SLOT_NAMES[slot]} va {seen[key]} bir xil.")
            else:
                seen[key] = SLOT_NAMES[slot]

        item["warnings"] = warnings
        item["errors"] = errors
        return item

    # ------------------------------------------------------------------ katak

    def _plain_text(self, cell: ET.Element) -> str:
        text = "".join(node.text or "" for node in cell.iter(W + "t"))
        return re.sub(r"\s+", " ", text.replace(" ", " ")).strip()

    def _cell(self, element: ET.Element) -> _Cell:
        cell = _Cell()
        first = True
        for paragraph in element.iter(W + "p"):
            if not first:
                cell.add_text("\n")
            first = False
            self._walk(paragraph, cell)
        cell.flush()
        return cell

    def _walk(self, node: ET.Element, cell: _Cell) -> None:
        for child in node:
            tag = child.tag
            if tag == M + "oMath":
                cell.add_math(self.omml.convert(child))
            elif tag == M + "oMathPara":
                self._walk(child, cell)
            elif tag == W + "r":
                self._run(child, cell)
            elif tag in (W + "pPr", W + "rPr", W + "del", W + "moveFrom", MC + "Fallback"):
                continue
            else:
                self._walk(child, cell)

    def _run(self, run: ET.Element, cell: _Cell) -> None:
        script = None
        align = run.find(W + "rPr/" + W + "vertAlign")
        if align is not None:
            script = {"superscript": "sup", "subscript": "sub"}.get(align.get(W + "val", ""))

        for child in run:
            tag = child.tag
            if tag == W + "t":
                text = child.text or ""
                if script:
                    cell.add_script(script, text)
                else:
                    cell.add_text(text)
            elif tag == W + "tab":
                cell.add_text(" ")
            elif tag in (W + "br", W + "cr"):
                cell.add_text("\n")
            elif tag == W + "noBreakHyphen":
                cell.add_text("-")
            elif tag == W + "sym":
                cell.add_text(_symbol(child.get(W + "char", "")))
            elif tag in (W + "object", W + "pict"):
                self._object(child, cell)
            elif tag == W + "drawing":
                self._drawing(child, cell)
            elif tag == MC + "AlternateContent":
                choice = child.find(MC + "Choice")
                if choice is not None:
                    self._run(choice, cell)

    def _object(self, element: ET.Element, cell: _Cell) -> None:
        """OLE obyekt: formula bo'lsa LaTeX, aks holda uning rasmi."""
        ole = element.find(O + "OLEObject")
        if ole is not None and ole.get("ProgID", "").lower().startswith("equation"):
            latex = self._ole_latex(ole.get(R + "id", ""))
            if latex:
                cell.add_math(latex)
            else:
                cell.add_text(UNKNOWN_FORMULA)
            return

        for image in element.iter(V + "imagedata"):
            ref = image.get(R + "id")
            if ref:
                cell.image_refs.append((ref, None))

    def _ole_latex(self, ref: str) -> str | None:
        if ref not in self._ole_cache:
            latex = None
            path = self.rels.get(ref)
            if path:
                try:
                    latex = ole_equation_to_latex(self.zip.read(path))
                except (MtefError, KeyError):
                    latex = None
            self._ole_cache[ref] = latex
        return self._ole_cache[ref]

    def _drawing(self, element: ET.Element, cell: _Cell) -> None:
        for holder in element.iter():
            blip = holder.find(A + "blip")
            if blip is None:
                continue
            ref = blip.get(R + "embed")
            if ref:
                cell.image_refs.append((ref, _crop_info(holder.find(A + "srcRect"))))

    # ------------------------------------------------------------------ rasm

    def _image(self, cell: _Cell, warnings: list[str], where: str) -> str | None:
        images: list[Image.Image] = []
        raw: bytes | None = None

        for ref, crop in cell.image_refs:
            path = self.rels.get(ref)
            if not path:
                continue
            try:
                data = self.zip.read(path)
            except KeyError:
                continue
            picture = _open_picture(data)
            if picture is None:
                ext = posixpath.splitext(path)[1].lstrip(".").upper() or "noma'lum"
                warnings.append(f"{where}dagi {ext} rasmni o'girib bo'lmadi - rasm tashlab ketildi.")
                continue

            keep_raw = crop is None and picture.format in ("PNG", "JPEG", "GIF")
            if crop:
                picture = _crop(picture, crop)
            images.append(picture)
            raw = data if keep_raw else None

        if not images:
            return None
        if len(images) == 1 and raw is not None and len(raw) <= MAX_IMAGE_BYTES \
                and max(images[0].size) <= MAX_IMAGE_SIDE:
            mime = {"PNG": "image/png", "JPEG": "image/jpeg", "GIF": "image/gif"}[images[0].format]
            return f"data:{mime};base64," + base64.b64encode(raw).decode()

        if len(images) > 1:
            warnings.append(f"{where}da {len(images)} ta rasm bor - bittaga birlashtirildi.")
        return _encode(_stack(images))


# ---------------------------------------------------------------------- yordamchilar


def _is_header(texts: list[str]) -> bool:
    first = _normalize(texts[0])
    second = _normalize(texts[1])
    if len(first) > 25:
        return False
    if first.startswith("savol"):
        return "javob" in second or "togri" in second
    return first.startswith("вопрос") and "ответ" in second


def _normalize(text: str) -> str:
    text = re.sub(r"[’‘`'ʻʼ]", "", text.lower())
    return re.sub(r"\s+", " ", text).strip()


def _symbol(code: str) -> str:
    """<w:sym> belgisi (Symbol shrifti) - eng ko'p uchraydiganlari."""
    try:
        value = int(code, 16)
    except ValueError:
        return ""
    if value >= 0xF000:
        value -= 0xF000
    table = {0xB0: "°", 0xB4: "×", 0xB8: "÷", 0xB1: "±", 0xA3: "≤", 0xB3: "≥", 0xB9: "≠",
             0xD7: "·", 0xA5: "∞", 0x70: "π", 0x61: "α", 0x62: "β", 0x67: "γ", 0xD0: "∠",
             0x44: "Δ", 0xBB: "≈", 0xAE: "→", 0xD6: "√"}
    return table.get(value, "")


def _open_picture(data: bytes) -> Image.Image | None:
    """Rasmni ochadi. EMF/WMF bo'lsa, ichiga joylangan bitmapni chiqarib oladi."""
    try:
        picture = Image.open(io.BytesIO(data))
        picture.load()
        return picture
    except Exception:  # noqa: BLE001 - EMF/WMF ni Pillow faqat Windows'da chiza oladi
        pass

    best: Image.Image | None = None
    for dib in _metafile_bitmaps(data):
        try:
            # DIB oldiga BMP fayl sarlavhasi qo'yilsa, oddiy BMP rasm hosil bo'ladi.
            header_size = struct.unpack_from("<I", dib, 0)[0]
            colors = struct.unpack_from("<I", dib, 32)[0] if header_size >= 36 else 0
            bits = struct.unpack_from("<H", dib, 14)[0]
            palette = (colors or (1 << bits if bits <= 8 else 0)) * 4
            offset = 14 + header_size + palette
            bmp = b"BM" + struct.pack("<IHHI", 14 + len(dib), 0, 0, offset) + dib
            picture = Image.open(io.BytesIO(bmp))
            picture.load()
        except Exception:  # noqa: BLE001 - buzilgan bitmap o'tkazib yuboriladi
            continue
        if best is None or picture.width * picture.height > best.width * best.height:
            best = picture
    return best.convert("RGB") if best is not None else None


def _metafile_bitmaps(data: bytes) -> list[bytes]:
    """EMF yoki WMF ichidagi bitmaplarni (DIB) qaytaradi. Faqat chizmadan iborat fayllarda bo'sh."""
    found: list[bytes] = []
    try:
        if len(data) > 44 and data[40:44] == b" EMF":
            # yozuv turi -> BITMAPINFO manzili turgan joy (yozuv boshidan)
            layouts = {80: 48, 81: 48, 76: 84, 77: 84, 114: 84}
            position = 0
            while position + 8 <= len(data):
                kind, size = struct.unpack_from("<II", data, position)
                if size < 8:
                    break
                if kind in layouts and position + layouts[kind] + 16 <= len(data):
                    info_at, info_size, bits_at, bits_size = struct.unpack_from("<IIII", data, position + layouts[kind])
                    if info_size and bits_size and bits_at >= info_at + info_size:
                        info = data[position + info_at: position + info_at + info_size]
                        bits = data[position + bits_at: position + bits_at + bits_size]
                        found.append(info + bits)
                position += size
        else:
            position = 22 if data[:4] == b"\xd7\xcd\xc6\x9a" else 0
            position += 18
            while position + 6 <= len(data):
                size, function = struct.unpack_from("<IH", data, position)
                if size < 3:
                    break
                start = {0x0F43: 28, 0x0B41: 26}.get(function)
                if start:
                    found.append(data[position + start: position + size * 2])
                position += size * 2
    except struct.error:
        pass
    return found


def _crop_info(rect: ET.Element | None) -> dict | None:
    if rect is None:
        return None
    try:
        sides = {name: float(rect.get(name, "0")) / 100000.0 for name in ("l", "t", "r", "b")}
    except ValueError:
        return None
    return sides if any(value > 0 for value in sides.values()) else None


def _crop(picture: Image.Image, crop: dict) -> Image.Image:
    width, height = picture.size
    left = max(0, min(int(width * crop["l"]), width - 1))
    top = max(0, min(int(height * crop["t"]), height - 1))
    right = max(left + 1, min(int(width * (1 - crop["r"])), width))
    bottom = max(top + 1, min(int(height * (1 - crop["b"])), height))
    return picture.crop((left, top, right, bottom))


def _stack(images: list[Image.Image]) -> Image.Image:
    if len(images) == 1:
        return images[0]
    images = [image.convert("RGBA") for image in images]
    gap = 8
    width = max(image.width for image in images)
    height = sum(image.height for image in images) + gap * (len(images) - 1)
    sheet = Image.new("RGBA", (width, height), (255, 255, 255, 255))
    y = 0
    for image in images:
        sheet.paste(image, (0, y), image)
        y += image.height + gap
    return sheet


def _encode(picture: Image.Image) -> str:
    """Rasmni API chegarasiga (2 MB) sig'adigan PNG yoki JPEG qilib kodlaydi."""
    if max(picture.size) > MAX_IMAGE_SIDE:
        picture.thumbnail((MAX_IMAGE_SIDE, MAX_IMAGE_SIDE))
    if picture.mode not in ("RGB", "RGBA", "L"):
        picture = picture.convert("RGBA")

    buffer = io.BytesIO()
    picture.save(buffer, format="PNG", optimize=True)
    mime = "image/png"

    quality = 88
    while buffer.tell() > MAX_IMAGE_BYTES and quality >= 40:
        flat = Image.new("RGB", picture.size, (255, 255, 255))
        flat.paste(picture, mask=picture.split()[-1] if picture.mode == "RGBA" else None)
        buffer = io.BytesIO()
        flat.save(buffer, format="JPEG", quality=quality)
        mime = "image/jpeg"
        quality -= 12

    return f"data:{mime};base64," + base64.b64encode(buffer.getvalue()).decode()


def parse_docx(content: bytes) -> dict:
    """Word fayl mazmunidan savollar ro'yxatini qaytaradi."""
    return DocxQuestionParser(content).parse()
