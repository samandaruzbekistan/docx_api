import xml.etree.ElementTree as ET

import pytest

from mathconv import MtefError, OmmlConverter, mtef_to_latex

NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'


def omml(inner: str) -> str:
    return OmmlConverter().convert(ET.fromstring(f"<m:oMath {NS}>{inner}</m:oMath>"))


def run(text: str) -> str:
    return f"<m:r><m:t>{text}</m:t></m:r>"


def test_fraction():
    xml = f"<m:f><m:num>{run('3x-7')}</m:num><m:den>{run('4')}</m:den></m:f>"
    assert omml(xml) == r"\frac{3x-7}{4}"


def test_superscript_and_subscript():
    assert omml(f"<m:sSup><m:e>{run('x')}</m:e><m:sup>{run('2')}</m:sup></m:sSup>") == "{x}^{2}"
    assert omml(f"<m:sSub><m:e>{run('a')}</m:e><m:sub>{run('n')}</m:sub></m:sSub>") == "{a}_{n}"


def test_square_root_hides_empty_degree():
    xml = (
        '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>'
        f"<m:e>{run('27')}</m:e></m:rad>"
    )
    assert omml(xml) == r"\sqrt{27}"


def test_root_with_degree():
    xml = f"<m:rad><m:deg>{run('3')}</m:deg><m:e>{run('8')}</m:e></m:rad>"
    assert omml(xml) == r"\sqrt[3]{8}"


def test_brackets():
    assert omml(f"<m:d><m:e>{run('x-105')}</m:e></m:d>") == r"\left( x-105 \right)"
    xml = (
        '<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr>'
        f"<m:e>{run('1;2')}</m:e></m:d>"
    )
    assert omml(xml) == r"\left[ 1;2 \right]"


def test_symbols_are_translated():
    assert omml(run("2∙3≤x±1")) == r"2\cdot 3\le x\pm 1"


def test_function_names():
    assert omml(run("sinx+cos2x")) == r"\sin x+\cos 2x"


def test_plain_words_inside_formula_stay_readable():
    xml = f"<m:sSup><m:e>{run('65')}</m:e><m:sup>{run('65')}</m:sup></m:sSup>" + run(" ") + run("ni") + run(" 8")
    assert omml(xml) == r"{65}^{65} \text{ ni } 8"


def test_normal_text_run():
    xml = '<m:r><m:rPr><m:nor/></m:rPr><m:t>agar</m:t></m:r>' + run("x=1")
    assert omml(xml) == r"\text{agar}x=1"


def test_unknown_element_is_reported_but_content_kept():
    converter = OmmlConverter()
    element = ET.fromstring(f"<m:oMath {NS}><m:mystery><m:e>{run('7')}</m:e></m:mystery></m:oMath>")
    assert converter.convert(element) == "7"
    assert converter.unknown == {"mystery"}


# MathType (MTEF v5) va eski Equation 3.0 (MTEF v3) obyektlaridan olingan haqiqiy baytlar.
MTEF_V5_HEADER = (
    "050100050244534d543500011357696e416c6c4261736963436f6465506167657300110554696d6573204e"
    "657720526f6d616e00110353796d626f6c001105436f7572696572204e65770011044d5420457874726100"
    "120008212f458f442f4150f4100f475f4150f21f1e4150f4150f4100f445f425f48f425f4100f4100f435f"
    "4100f48f45f42a5f48f48f4100f4100f40f48f417f48f4100f412a5f445f45f45f45f45f410f0c01000100"
    "010202020200020001010100030001000400000a0100"
)
MTEF_V5_FRACTION = MTEF_V5_HEADER + "03000b000001000200883300000100020088340000000000"
MTEF_V5_MIXED = MTEF_V5_HEADER + (
    "02008378000204863d003d020088330003000b000001000200883100000100020088350000000000"
)
MTEF_V3_FRACTION = "030101030a0a01030e0000010288330000010288340000000000"
MTEF_V3_POWERS = (
    "03010103010a01030e00000112837000030f00000b110102883100028833000000000a0112837000030f00"
    "000b11010288310002883100000000000a02861222030e00000112837800030f00000b1101028831000288"
    "37000000000a0112837800030f00000b11010288310002883400000000000a02862b00030e000001128379"
    "00030f00000b110102883100028835000000000a0112837900030f00000b11010288310002883300000000"
    "000a02823b000000"
)


@pytest.mark.parametrize(
    ("data", "expected"),
    [
        (MTEF_V5_FRACTION, r"\frac{3}{4}"),
        (MTEF_V5_MIXED, r"x=3\frac{1}{5}"),
        (MTEF_V3_FRACTION, r"\frac{3}{4}"),
        (MTEF_V3_POWERS, r"\frac{p^{13}}{p^{11}}-\frac{x^{17}}{x^{14}}+\frac{y^{15}}{y^{13}};"),
    ],
)
def test_mathtype_and_equation_objects(data, expected):
    assert mtef_to_latex(bytes.fromhex(data)) == expected


def test_broken_mtef_raises_a_clear_error():
    with pytest.raises(MtefError):
        mtef_to_latex(bytes.fromhex("09ffff"))
    with pytest.raises(MtefError):
        mtef_to_latex(bytes.fromhex(MTEF_V3_FRACTION[:30]))
