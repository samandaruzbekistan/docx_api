import base64
import io
import struct
import zipfile

import pytest

from docx_parser import UNKNOWN_FORMULA, DocxError, parse_docx

NAMESPACES = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:o="urn:schemas-microsoft-com:office:office"'
)
PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


def t(text: str, script: str | None = None) -> str:
    props = f'<w:rPr><w:vertAlign w:val="{script}"/></w:rPr>' if script else ""
    return f'<w:r>{props}<w:t xml:space="preserve">{text}</w:t></w:r>'


def fraction(num: str, den: str) -> str:
    return (
        "<m:oMath><m:f>"
        f"<m:num><m:r><m:t>{num}</m:t></m:r></m:num>"
        f"<m:den><m:r><m:t>{den}</m:t></m:r></m:den>"
        "</m:f></m:oMath>"
    )


def image(ref: str) -> str:
    return f'<w:r><w:drawing><a:graphic><a:blipFill><a:blip r:embed="{ref}"/></a:blipFill></a:graphic></w:drawing></w:r>'


def equation(ref: str) -> str:
    return (
        '<w:r><w:object><v:shape><v:imagedata r:id="rIdPreview"/></v:shape>'
        f'<o:OLEObject ProgID="Equation.3" r:id="{ref}"/></w:object></w:r>'
    )


def cell(content: str) -> str:
    inner = content if content.startswith("<") else t(content)
    return f"<w:tc><w:p>{inner}</w:p></w:tc>"


def row(*cells: str) -> str:
    return "<w:tr>" + "".join(cell(item) if item else "<w:tc><w:p/></w:tc>" for item in cells) + "</w:tr>"


def docx(*rows: str, files: dict[str, bytes] | None = None, body: str | None = None) -> bytes:
    """Xotirada kichik Word hujjati yasaydi. `files`: {rel id: (yo'l, baytlar)}."""
    content = body if body is not None else "<w:tbl>" + "".join(rows) + "</w:tbl>"
    document = f'<w:document {NAMESPACES}><w:body>{content}</w:body></w:document>'
    relations = "".join(
        f'<Relationship Id="{ref}" Type="x" Target="{path}"/>' for ref, (path, _) in (files or {}).items()
    )
    rels = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{relations}</Relationships>"
    )
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("word/document.xml", document)
        archive.writestr("word/_rels/document.xml.rels", rels)
        for path, data in (files or {}).values():
            archive.writestr("word/" + path, data)
    return buffer.getvalue()


def texts(question: dict) -> list[str]:
    return [question[slot]["text"] for slot in ("question", "correct", "wrong1", "wrong2", "wrong3")]


def test_skips_the_header_row_and_reads_the_author_column():
    result = parse_docx(docx(
        row("Savol", "Togri javob", "Notogri javob", "Notogri javob", "Notogri javob", "Test tuzuvchi"),
        row("2+2 necha?", "4", "3", "5", "6", "Temirov Maxmud"),
        row("3+3 necha?", "6", "1", "2", "3", "Temirov Maxmud"),
    ))

    assert result["author"] == "Temirov Maxmud"
    assert [q["row"] for q in result["questions"]] == [2, 3]
    assert texts(result["questions"][0]) == ["2+2 necha?", "4", "3", "5", "6"]
    assert result["questions"][0]["errors"] == []


def test_five_column_tables_without_a_header_work():
    result = parse_docx(docx(row("Poytaxt qayer?", "Toshkent", "Samarqand", "Buxoro", "Xiva")))

    assert result["author"] == ""
    assert len(result["questions"]) == 1


def test_a_question_that_only_starts_with_the_word_savol_is_not_a_header():
    result = parse_docx(docx(row("Savol: 2+2 nechaga teng bo'ladi, hisoblang?", "4", "3", "5", "6")))

    assert len(result["questions"]) == 1


def test_empty_rows_are_dropped():
    result = parse_docx(docx(
        row("Savol 1", "a", "b", "c", "d"),
        row("", "", "", "", ""),
        row(" ", " ", " ", " ", " "),
    ))

    assert len(result["questions"]) == 1


def test_word_formulas_become_latex_inside_the_text():
    result = parse_docx(docx(row(
        t("Hisoblang: ") + fraction("1", "2") + t(" va ") + fraction("3", "4"),
        fraction("5", "4"), "1", "2", "3",
    )))

    question = result["questions"][0]
    assert question["question"]["text"] == r"Hisoblang: \(\frac{1}{2}\) va \(\frac{3}{4}\)"
    assert question["correct"]["text"] == r"\(\frac{5}{4}\)"


def test_superscripts_typed_as_plain_text_are_attached_to_their_base():
    result = parse_docx(docx(row(
        t("Ajrating: a") + t("4", "superscript") + t(" -3a") + t("3", "superscript") + t("+2a-6="),
        t("2") + t("n", "superscript") + t("+1", "superscript"),
        t("H") + t("2", "subscript") + t("O"), "b", "c",
    )))

    question = result["questions"][0]
    assert question["question"]["text"] == r"Ajrating: \(a^{4}\) -\(3a^{3}\)+2a-6="
    assert question["correct"]["text"] == r"\(2^{n+1}\)"
    assert question["wrong1"]["text"] == r"\(H_{2}\)O"


def test_empty_superscript_leftovers_are_ignored():
    result = parse_docx(docx(row(t("1440 ") + t(" ", "superscript") + t("soni"), "a", "b", "c", "d")))

    assert result["questions"][0]["question"]["text"] == "1440 soni"


def test_missing_answers_are_reported():
    result = parse_docx(docx(
        row("Javobsiz", "", "b", "c", "d"),
        row("Faqat to'g'ri", "a", "", "", ""),
        row("Bitta kam", "a", "b", "c", ""),
    ))

    first, second, third = result["questions"]
    assert first["errors"] == ["To'g'ri javob bo'sh."]
    assert second["errors"] == ["Noto'g'ri javoblar yo'q."]
    assert third["errors"] == [] and third["warnings"] == ["Faqat 2 ta noto'g'ri javob bor."]


def test_identical_answers_are_flagged():
    result = parse_docx(docx(row("Savol", "5", "5 ", "6", "7")))

    assert result["questions"][0]["warnings"] == ["1-noto'g'ri javob va to'g'ri javob bir xil."]


def test_rows_with_too_few_columns_are_reported():
    result = parse_docx(docx(row("Savol", "a", "b", "c", "d"), row("Chala qator", "a", "b")))

    assert len(result["questions"]) == 1
    assert "5 tadan kam ustunli" in result["warnings"][0] and "2" in result["warnings"][0]


def test_images_are_returned_as_data_uris():
    result = parse_docx(docx(
        row(t("Rasmga qarang") + image("rId7"), image("rId7"), "b", "c", "d"),
        files={"rId7": ("media/image1.png", PNG)},
    ))

    question = result["questions"][0]
    assert question["question"]["image"].startswith("data:image/png;base64,")
    assert question["correct"] == {"text": "", "image": question["question"]["image"]}
    assert question["errors"] == []


def emf_with_bitmap(width: int, height: int) -> bytes:
    """Word ko'pincha qo'yilgan rasmni shunday saqlaydi: EMF ichida bitta STRETCHDIBITS yozuvi."""
    from PIL import Image

    buffer = io.BytesIO()
    Image.new("RGB", (width, height), (200, 30, 30)).save(buffer, format="BMP")
    dib = buffer.getvalue()[14:]
    info, bits = dib[:40], dib[40:]
    header = struct.pack("<II", 1, 88) + bytes(32) + b" EMF" + bytes(44)
    record = struct.pack("<II", 81, 80 + len(dib)) + bytes(40) + struct.pack("<IIII", 80, 40, 120, len(bits))
    record += bytes(16) + info + bits
    return header + record + struct.pack("<IIIII", 14, 20, 0, 0, 20)


def test_bitmaps_wrapped_in_emf_files_are_extracted():
    result = parse_docx(docx(
        row(t("Soddalashtiring") + image("rId3"), "a", "b", "c", "d"),
        files={"rId3": ("media/image1.emf", emf_with_bitmap(40, 12))},
    ))

    question = result["questions"][0]
    assert question["warnings"] == []
    picture = question["question"]["image"]
    assert picture.startswith("data:image/png;base64,")
    from PIL import Image
    opened = Image.open(io.BytesIO(base64.b64decode(picture.split(",", 1)[1])))
    assert opened.size == (40, 12) and opened.convert("RGB").getpixel((5, 5)) == (200, 30, 30)


def test_vector_only_metafiles_are_dropped_with_a_warning():
    empty_emf = struct.pack("<II", 1, 88) + bytes(32) + b" EMF" + bytes(44) + struct.pack("<IIIII", 14, 20, 0, 0, 20)
    result = parse_docx(docx(
        row(t("Chizma") + image("rId3"), "a", "b", "c", "d"),
        files={"rId3": ("media/image1.emf", empty_emf)},
    ))

    question = result["questions"][0]
    assert question["question"]["image"] is None
    assert question["warnings"] == ["savoldagi EMF rasmni o'girib bo'lmadi - rasm tashlab ketildi."]


def test_unreadable_equation_objects_are_marked_for_manual_entry():
    result = parse_docx(docx(
        row(t("Agar ") + equation("rId9") + t(" bo'lsa"), "a", "b", "c", "d"),
        files={"rId9": ("embeddings/oleObject1.bin", b"buzilgan fayl")},
    ))

    question = result["questions"][0]
    assert question["question"]["text"] == f"Agar {UNKNOWN_FORMULA} bo'lsa"
    assert question["question"]["image"] is None
    assert question["errors"] == ["savolda o'girilmagan formula bor - qo'lda yozing."]


def test_files_that_are_not_word_documents_are_rejected():
    with pytest.raises(DocxError):
        parse_docx(b"bu docx emas")
    with pytest.raises(DocxError):
        parse_docx(docx(body="<w:p><w:r><w:t>Jadval yo'q</w:t></w:r></w:p>"))
