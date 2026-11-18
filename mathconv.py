"""Word formulalarini LaTeX'ga o'girish.

Ikki manba qo'llab-quvvatlanadi:
  * OMML  - Word'ning o'z formula muharriri (<m:oMath>);
  * MTEF  - MathType va eski Equation 3.0 obyektlari (OLE ichidagi "Equation Native").

Natija delimitersiz LaTeX satri; uni chaqiruvchi tomon \\( ... \\) ichiga oladi.
"""

from __future__ import annotations

import io
import re
import struct
import xml.etree.ElementTree as ET

M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"

# Matematik rejimda maxsus yozilishi kerak bo'lgan belgilar.
SYMBOLS = {
    "∙": r"\cdot ", "⋅": r"\cdot ", "·": r"\cdot ", "•": r"\cdot ",
    "×": r"\times ", "÷": r"\div ", "±": r"\pm ", "∓": r"\mp ",
    "≤": r"\le ", "≥": r"\ge ", "≠": r"\ne ", "≈": r"\approx ", "≡": r"\equiv ",
    "∞": r"\infty ", "∈": r"\in ", "∉": r"\notin ", "∪": r"\cup ", "∩": r"\cap ",
    "⊂": r"\subset ", "⊃": r"\supset ", "∅": r"\emptyset ", "∠": r"\angle ",
    "△": r"\triangle ", "∆": r"\Delta ", "⊥": r"\perp ", "∥": r"\parallel ",
    "→": r"\to ", "←": r"\leftarrow ", "↔": r"\leftrightarrow ",
    "⇒": r"\Rightarrow ", "⇔": r"\Leftrightarrow ",
    "°": r"^\circ ", "′": "'", "″": "''", "−": "-", "–": "-", "…": r"\ldots ",
    "√": r"\surd ", "∑": r"\sum ", "∏": r"\prod ", "∫": r"\int ",
    "α": r"\alpha ", "β": r"\beta ", "γ": r"\gamma ", "δ": r"\delta ",
    "ε": r"\varepsilon ", "θ": r"\theta ", "λ": r"\lambda ", "μ": r"\mu ",
    "π": r"\pi ", "ρ": r"\rho ", "σ": r"\sigma ", "τ": r"\tau ",
    "φ": r"\varphi ", "ϕ": r"\varphi ", "ω": r"\omega ", "Ω": r"\Omega ",
    "Δ": r"\Delta ", "Σ": r"\Sigma ", "Π": r"\Pi ",
    "%": r"\%", "#": r"\#", "&": r"\&", "$": r"\$", "_": r"\_",
    "{": r"\{", "}": r"\}", "~": r"\sim ", "\\": r"\backslash ",
    " ": " ", "⁡": "", "​": "", "⁢": "", "⁣": "",
}

FUNCTIONS = (
    "arcsin", "arccos", "arctg", "arcctg", "arctan", "arccot",
    "sin", "cos", "ctg", "tg", "tan", "cot", "log", "ln", "lg", "lim", "max", "min",
)
_FUNC_RE = re.compile(r"(?<![A-Za-z\\])(" + "|".join(FUNCTIONS) + r")")
# Formula ichiga yozib yuborilgan oddiy so'zlar ("ni", "va", "bo'lsa").
_WORD_RE = re.compile(r"(?:(?<=\s)|^)([A-Za-z'’`‘ʻ]{2,})(?=\s|$)")


def _escape(text: str) -> str:
    return "".join(SYMBOLS.get(ch, ch) for ch in text)


def _math_text(text: str) -> str:
    """Formula ichidagi oddiy matn bo'lagini LaTeX'ga moslaydi."""
    text = text.replace("⁡", "")

    def word(match: re.Match) -> str:
        token = match.group(1)
        if token in FUNCTIONS:
            return token
        return "\x00" + token + "\x01"

    if " " in text or " " in text:
        text = _WORD_RE.sub(word, text.replace(" ", " "))

    out = []
    for part in re.split(r"(\x00[^\x01]*\x01)", text):
        if part.startswith("\x00"):
            out.append(r"\text{ " + part[1:-1] + " }")
        else:
            out.append(_FUNC_RE.sub(lambda m: "\\" + m.group(1) + " ", _escape(part)))
    return "".join(out)


# --------------------------------------------------------------------------- OMML

_LEFT = {"(": "(", "[": "[", "{": r"\{", "|": "|", "‖": r"\|", "⌊": r"\lfloor ",
         "⌈": r"\lceil ", "〈": r"\langle ", "⟨": r"\langle ", "": "."}
_RIGHT = {")": ")", "]": "]", "}": r"\}", "|": "|", "‖": r"\|", "⌋": r"\rfloor ",
          "⌉": r"\rceil ", "〉": r"\rangle ", "⟩": r"\rangle ", "": "."}
_ACCENTS = {"̂": r"\hat", "̄": r"\bar", "̅": r"\overline", "¯": r"\bar",
            "⃗": r"\vec", "̇": r"\dot", "̈": r"\ddot", "̃": r"\tilde",
            "̌": r"\check", "̆": r"\breve"}
_NARY = {"∑": r"\sum", "∏": r"\prod", "∫": r"\int", "∬": r"\iint", "∭": r"\iiint",
         "∮": r"\oint", "⋃": r"\bigcup", "⋂": r"\bigcap"}


class OmmlConverter:
    """<m:oMath> elementini LaTeX'ga o'giradi. Tanilmagan teglar `unknown` ga yoziladi."""

    def __init__(self) -> None:
        self.unknown: set[str] = set()

    def convert(self, element: ET.Element) -> str:
        return re.sub(r"\s+", " ", self._node(element)).strip()

    def _children(self, element: ET.Element | None) -> str:
        if element is None:
            return ""
        out: list[str] = []
        pending = ""  # ketma-ket oddiy bo'laklar birga ko'riladi: so'zlar bo'linib ketmasin
        for child in element:
            if child.tag == M + "r" and not self._is_normal_text(child):
                pending += "".join(t.text or "" for t in child.findall(M + "t"))
                continue
            if pending:
                out.append(_math_text(pending))
                pending = ""
            out.append(self._node(child))
        if pending:
            out.append(_math_text(pending))
        return "".join(out)

    @staticmethod
    def _is_normal_text(run: ET.Element) -> bool:
        props = run.find(M + "rPr")
        return props is not None and props.find(M + "nor") is not None

    def _slot(self, element: ET.Element, name: str) -> str:
        return self._children(element.find(M + name)).strip()

    @staticmethod
    def _prop(element: ET.Element, props: str, name: str, default: str | None = None) -> str | None:
        holder = element.find(M + props)
        if holder is None:
            return default
        item = holder.find(M + name)
        if item is None:
            return default
        return item.get(M + "val", default)

    def _node(self, el: ET.Element) -> str:
        if not el.tag.startswith(M):
            return ""
        tag = el.tag[len(M):]

        if tag in ("oMath", "oMathPara", "e", "num", "den", "sub", "sup", "deg", "fName", "lim"):
            return self._children(el)
        if tag == "r":
            text = "".join(t.text or "" for t in el.findall(M + "t"))
            props = el.find(M + "rPr")
            if props is not None and props.find(M + "nor") is not None and text.strip():
                return r"\text{" + text.replace(" ", " ") + "}"
            return _math_text(text)
        if tag == "f":
            num, den = self._slot(el, "num"), self._slot(el, "den")
            kind = self._prop(el, "fPr", "type")
            if kind == "lin":
                return f"{num}/{den}"
            if kind == "noBar":
                return r"\binom{%s}{%s}" % (num, den)
            return r"\frac{%s}{%s}" % (num, den)
        if tag == "sSup":
            return "{%s}^{%s}" % (self._slot(el, "e"), self._slot(el, "sup"))
        if tag == "sSub":
            return "{%s}_{%s}" % (self._slot(el, "e"), self._slot(el, "sub"))
        if tag == "sSubSup":
            return "{%s}_{%s}^{%s}" % (self._slot(el, "e"), self._slot(el, "sub"), self._slot(el, "sup"))
        if tag == "sPre":
            return "{}_{%s}^{%s}{%s}" % (self._slot(el, "sub"), self._slot(el, "sup"), self._slot(el, "e"))
        if tag == "rad":
            degree = self._slot(el, "deg")
            if degree and self._prop(el, "radPr", "degHide") not in ("1", "on", "true"):
                return r"\sqrt[%s]{%s}" % (degree, self._slot(el, "e"))
            return r"\sqrt{%s}" % self._slot(el, "e")
        if tag == "d":
            left = self._prop(el, "dPr", "begChr", "(")
            right = self._prop(el, "dPr", "endChr", ")")
            sep = self._prop(el, "dPr", "sepChr", "|")
            inner = (" " + _escape(sep or "") + " ").join(
                self._children(item).strip() for item in el.findall(M + "e")
            )
            return r"\left%s %s \right%s" % (
                _LEFT.get(left or "", _escape(left or "")), inner,
                _RIGHT.get(right or "", _escape(right or "")),
            )
        if tag == "func":
            return "%s{%s}" % (self._slot(el, "fName"), self._slot(el, "e"))
        if tag == "nary":
            char = self._prop(el, "naryPr", "chr", "∫")
            out = _NARY.get(char or "", _escape(char or ""))
            sub, sup = self._slot(el, "sub"), self._slot(el, "sup")
            if sub:
                out += "_{%s}" % sub
            if sup:
                out += "^{%s}" % sup
            return "%s {%s}" % (out, self._slot(el, "e"))
        if tag == "acc":
            char = self._prop(el, "accPr", "chr", "̂")
            return "%s{%s}" % (_ACCENTS.get(char or "", r"\hat"), self._slot(el, "e"))
        if tag == "bar":
            pos = self._prop(el, "barPr", "pos", "bot")
            return (r"\overline{%s}" if pos == "top" else r"\underline{%s}") % self._slot(el, "e")
        if tag == "limLow":
            return r"%s_{%s}" % (self._slot(el, "e"), self._slot(el, "lim"))
        if tag == "limUpp":
            return r"\overset{%s}{%s}" % (self._slot(el, "lim"), self._slot(el, "e"))
        if tag == "groupChr":
            char = self._prop(el, "groupChrPr", "chr", "⏟")
            name = r"\overbrace" if char == "⏞" else r"\underbrace"
            return "%s{%s}" % (name, self._slot(el, "e"))
        if tag == "m":
            rows = [
                " & ".join(self._children(cell).strip() for cell in row.findall(M + "e"))
                for row in el.findall(M + "mr")
            ]
            return r"\begin{matrix}%s\end{matrix}" % r" \\ ".join(rows)
        if tag == "eqArr":
            rows = [self._children(item).strip() for item in el.findall(M + "e")]
            return r"\begin{matrix}%s\end{matrix}" % r" \\ ".join(rows)
        if tag in ("box", "borderBox", "phant"):
            return self._slot(el, "e")
        if tag.endswith("Pr") or tag in ("t", "ctrlPr"):
            return ""

        self.unknown.add(tag)
        return self._children(el)


# --------------------------------------------------------------------------- MTEF


class MtefError(ValueError):
    """MathType/Equation obyektini o'qib bo'lmadi."""


_END, _LINE, _CHAR, _TMPL, _PILE, _MATRIX, _EMBELL, _RULER, _FONT_STYLE, _SIZE = range(10)

# Shablon turlari (MTEF v5 nomlari). v3 raqamlari _V3_TEMPLATES orqali shunga keltiriladi.
_FENCE = {
    "angle": (r"\left\langle ", r"\right\rangle "), "paren": (r"\left( ", r"\right) "),
    "brace": (r"\left\{ ", r"\right\} "), "brack": (r"\left[ ", r"\right] "),
    "bar": (r"\left| ", r"\right| "), "dbar": (r"\left\| ", r"\right\| "),
    "floor": (r"\left\lfloor ", r"\right\rfloor "), "ceil": (r"\left\lceil ", r"\right\rceil "),
}
_V5_TEMPLATES = {
    0: "angle", 1: "paren", 2: "brace", 3: "brack", 4: "bar", 5: "dbar", 6: "floor", 7: "ceil",
    8: "obrack", 9: "interval", 10: "root", 11: "frac", 12: "ubar", 13: "obar", 14: "arrow",
    15: "integ", 16: "sum", 17: "prod", 18: "coprod", 19: "union", 20: "inter", 21: "intop",
    22: "sumop", 23: "lim", 24: "hbrace", 25: "hbrack", 26: "ldiv", 27: "sub", 28: "sup",
    29: "subsup", 30: "dirac", 31: "vec", 32: "tilde", 33: "hat", 34: "arc", 36: "strike", 37: "box",
}
_V3_TEMPLATES = {
    0: "angle", 1: "paren", 2: "brace", 3: "brack", 4: "bar", 5: "dbar", 6: "floor", 7: "ceil",
    8: "interval", 9: "interval", 10: "interval", 11: "interval", 12: "interval",
    13: "root", 14: "frac", 15: "script", 16: "ubar", 17: "obar", 18: "arrow", 19: "arrow", 20: "arrow",
    21: "integ", 22: "integ", 23: "integ", 24: "integ", 25: "integ", 26: "integ", 27: "integ", 28: "integ",
    29: "sum", 30: "sum", 31: "prod", 32: "prod", 33: "coprod", 34: "coprod",
    35: "union", 36: "union", 37: "inter", 38: "inter", 39: "lim", 40: "ldiv", 41: "frac", 42: "intop",
    43: "sumop", 44: "script", 45: "dirac", 46: "dirac", 47: "dirac", 48: "vec", 49: "tilde", 50: "hat",
}
_BIG_OP = {"integ": r"\int", "sum": r"\sum", "prod": r"\prod", "coprod": r"\coprod",
           "union": r"\bigcup", "inter": r"\bigcap", "intop": r"\int", "sumop": r"\sum"}
_EMBELLS = {2: r"\dot", 3: r"\ddot", 4: r"\dddot", 8: r"\tilde", 9: r"\hat", 11: r"\vec",
            17: r"\bar", 5: "'", 6: "''", 7: "'''"}


class _Reader:
    def __init__(self, data: bytes) -> None:
        self.data = data
        self.pos = 0

    def byte(self) -> int:
        if self.pos >= len(self.data):
            raise MtefError("ma'lumot kutilmaganda tugadi")
        value = self.data[self.pos]
        self.pos += 1
        return value

    def word(self) -> int:
        low = self.byte()
        return low | (self.byte() << 8)

    def uint(self) -> int:
        value = self.byte()
        return self.word() if value == 255 else value

    def cstring(self) -> str:
        start = self.pos
        while self.byte() != 0:
            pass
        return self.data[start:self.pos - 1].decode("latin-1")

    def nudge(self) -> None:
        dx, dy = self.byte(), self.byte()
        if dx == 128 and dy == 128:
            self.word()
            self.word()

    def nibble_array(self) -> None:
        """EQN_PREFS ichidagi o'lcham massivini tashlab o'tadi."""
        count = self.byte()
        nibbles: list[int] = []

        def nxt() -> int:
            if not nibbles:
                value = self.byte()
                nibbles.extend((value >> 4, value & 0x0F))
            return nibbles.pop(0)

        for _ in range(count):
            nxt()  # o'lchov birligi
            while nxt() != 0x0F:
                pass


class MtefParser:
    """MTEF (v2-v5) oqimini LaTeX'ga o'giradi."""

    def __init__(self, data: bytes) -> None:
        self.r = _Reader(data)
        self.version = self.r.byte()
        if self.version not in (2, 3, 4, 5):
            raise MtefError(f"MTEF versiyasi noma'lum: {self.version}")
        self.r.byte()  # platforma
        self.r.byte()  # mahsulot
        self.r.byte()  # versiya
        self.r.byte()  # kichik versiya
        if self.version == 5:
            self.r.cstring()  # ilova kaliti
            self.r.byte()  # formula sozlamalari
        self.v5 = self.version == 5

    def convert(self) -> str:
        parts = self._objects(top=True)
        return re.sub(r"\s+", " ", "".join(parts)).strip()

    # -- yozuvlar ----------------------------------------------------------

    def _tag(self) -> tuple[int, int]:
        value = self.r.byte()
        if self.v5:
            if value in (_LINE, _CHAR, _TMPL, _PILE, _MATRIX, _EMBELL):
                return value, self.r.byte()
            return value, 0
        return value & 0x0F, value >> 4

    def _objects(self, top: bool = False) -> list[str]:
        """END gacha bo'lgan obyektlar ro'yxatini o'qiydi."""
        out: list[str] = []
        while True:
            if self.r.pos >= len(self.r.data):
                if top:
                    return out
                raise MtefError("END yozuvi topilmadi")
            kind, options = self._tag()
            if kind == _END:
                return out
            piece = self._record(kind, options, out)
            if piece:
                out.append(piece)

    def _record(self, kind: int, options: int, before: list[str]) -> str:
        r = self.r
        if kind == _LINE:
            if options & 0x08:
                r.nudge()
            if options & 0x04:
                r.word()
            if options & 0x02:
                self._ruler_record()
            if options & 0x01:
                return ""
            return "".join(self._objects())
        if kind == _CHAR:
            return self._char(options)
        if kind == _TMPL:
            return self._template(options, before)
        if kind == _PILE:
            if options & 0x08:
                r.nudge()
            r.byte()
            r.byte()
            if options & 0x02:
                self._ruler_record()
            rows = [row for row in self._objects()]
            return r"\begin{matrix}%s\end{matrix}" % r" \\ ".join(rows)
        if kind == _MATRIX:
            return self._matrix(options)
        if kind == _EMBELL:
            if options & 0x08:
                r.nudge()
            r.byte()
            return ""
        if kind == _RULER:
            self._ruler_body()
            return ""
        if kind == _FONT_STYLE:
            if self.v5:
                r.uint()
                r.byte()
            else:  # v3 FONT: shrift turi, uslub, nom
                r.byte()
                r.byte()
                r.cstring()
            return ""
        if kind == _SIZE:
            first = r.byte()
            if first == 101:
                r.byte()
                r.byte()
            elif first == 100:
                r.byte()
                r.word()
            else:
                r.byte()
            return ""
        if 10 <= kind <= 14:
            return ""
        if self.v5:
            if kind == 15:
                r.uint()
                return ""
            if kind == 16:
                opts = r.byte()
                for _ in range(4 if opts & 0x01 else 3):
                    r.word()
                if opts & 0x04:
                    r.cstring()
                return ""
            if kind == 17:
                r.uint()
                r.cstring()
                return ""
            if kind == 18:
                r.byte()
                r.nibble_array()
                r.nibble_array()
                for _ in range(r.byte()):
                    if r.byte() != 0:
                        r.byte()
                return ""
            if kind == 19:
                r.cstring()
                return ""
            if kind >= 100:
                r.pos += r.uint()
                return ""
        raise MtefError(f"noma'lum yozuv turi: {kind}")

    def _ruler_record(self) -> None:
        if self.r.byte() != _RULER:
            raise MtefError("RULER yozuvi kutilgan edi")
        self._ruler_body()

    def _ruler_body(self) -> None:
        for _ in range(self.r.byte()):
            self.r.byte()
            self.r.word()

    def _char(self, options: int) -> str:
        r = self.r
        if options & 0x08:
            r.nudge()
        typeface = r.byte()
        code = 0
        if self.v5:
            if not options & 0x20:
                code = r.word()
            if options & 0x04:
                position = r.byte()
                code = code or position
            if options & 0x10:
                position = r.word()
                code = code or position
        else:
            code = r.word()

        accent = ""
        if options & 0x02:
            accent = self._embellishments()

        text = self._char_text(code, typeface)
        if accent and text:
            if accent.startswith("'"):
                return text + accent
            return "%s{%s}" % (accent, text)
        return text

    def _embellishments(self) -> str:
        accent = ""
        while True:
            kind, options = self._tag()
            if kind == _END:
                return accent
            if kind != _EMBELL:
                raise MtefError("bezak yozuvi kutilgan edi")
            if options & 0x08:
                self.r.nudge()
            accent = _EMBELLS.get(self.r.byte(), accent)

    @staticmethod
    def _char_text(code: int, typeface: int) -> str:
        if code in (0, 0x200B) or 0xEF00 <= code <= 0xEF0F:
            return " " if code in (0xEF02, 0xEF04, 0xEF05, 0xEF08, 0xEF09) else ""
        if 0xE000 <= code <= 0xF8FF:
            if 0xF020 <= code <= 0xF0FF:  # Symbol shrifti xususiy sohasi
                code -= 0xF000
            else:
                return ""
        char = chr(code)
        style = typeface - 128 if typeface >= 128 else typeface
        if style == 1 and char.isalpha():
            return r"\text{" + char + "}"
        return SYMBOLS.get(char, char)

    def _matrix(self, options: int) -> str:
        r = self.r
        if options & 0x08:
            r.nudge()
        r.byte()
        r.byte()
        r.byte()
        rows, cols = r.byte(), r.byte()
        r.pos += ((rows + 1) * 2 + 7) // 8
        r.pos += ((cols + 1) * 2 + 7) // 8
        cells = self._objects()
        lines = []
        for index in range(0, len(cells), max(cols, 1)):
            lines.append(" & ".join(cells[index:index + cols]))
        return r"\begin{matrix}%s\end{matrix}" % r" \\ ".join(lines)

    def _template(self, options: int, before: list[str]) -> str:
        r = self.r
        if options & 0x08:
            r.nudge()
        selector = r.byte()
        variation = r.byte()
        if self.v5 and variation & 0x80:
            variation = (variation & 0x7F) | (r.byte() << 7)
        r.byte()  # shablon sozlamalari

        slots = self._template_slots()
        name = (_V5_TEMPLATES if self.v5 else _V3_TEMPLATES).get(selector)
        if name is None:
            raise MtefError(f"noma'lum shablon: {selector}")

        def slot(index: int) -> str:
            return slots[index] if index < len(slots) and slots[index] is not None else ""

        filled = [item for item in slots if item]

        if name in _FENCE:
            left, right = _FENCE[name]
            if self.v5:
                if variation & 0x03 == 1:
                    right = r"\right. "
                elif variation & 0x03 == 2:
                    left = r"\left. "
            return left + slot(0) + right
        if name in ("interval", "obrack", "dirac"):
            return r"\left[ " + slot(0) + r"\right] " if name != "dirac" else "".join(filled)
        if name == "frac":
            return r"\frac{%s}{%s}" % (slot(0), slot(1))
        if name == "root":
            return r"\sqrt[%s]{%s}" % (slot(1), slot(0)) if slot(1) else r"\sqrt{%s}" % slot(0)
        if name in ("sub", "sup", "subsup", "script"):
            return self._script(name, variation, slots, before)
        if name == "ubar":
            return r"\underline{%s}" % slot(0)
        if name == "obar":
            return r"\overline{%s}" % slot(0)
        if name in ("vec", "arrow"):
            return r"\overrightarrow{%s}" % slot(0)
        if name == "tilde":
            return r"\widetilde{%s}" % slot(0)
        if name in ("hat", "arc"):
            return r"\widehat{%s}" % slot(0)
        if name in _BIG_OP:
            out = _BIG_OP[name]
            if slot(1):
                out += "_{%s}" % slot(1)
            if slot(2):
                out += "^{%s}" % slot(2)
            return out + " " + slot(0)
        if name == "lim":
            return r"\lim_{%s} %s" % (slot(1), slot(0)) if slot(1) else r"\lim " + slot(0)
        if name == "ldiv":
            return "%s )\\overline{%s}" % (slot(1), slot(0))
        return "".join(filled)

    def _template_slots(self) -> list[str | None]:
        """Shablon ichidagi bo'limlar (LINE); bo'sh bo'lim None bilan belgilanadi."""
        slots: list[str | None] = []
        while True:
            kind, options = self._tag()
            if kind == _END:
                return slots
            if kind == _LINE:
                if options & 0x01:
                    self._record(kind, options, [])
                    slots.append(None)
                else:
                    slots.append(self._record(kind, options, []))
            elif kind in (_PILE, _MATRIX, _TMPL):
                slots.append(self._record(kind, options, []))
            else:
                self._record(kind, options, [])  # qavs belgilari va boshqalar

    @staticmethod
    def _script(name: str, variation: int, slots: list[str | None], before: list[str]) -> str:
        filled = [item for item in slots if item]
        sub = sup = ""
        if name == "subsup" or (name == "script" and len(filled) >= 2):
            sub = (slots[0] or "") if slots else ""
            sup = (slots[1] or "") if len(slots) > 1 else ""
        elif name == "sub":
            sub = filled[0] if filled else ""
        elif name == "sup":
            sup = filled[-1] if filled else ""
        else:  # v3 "script": variation 0 - yuqori, 1 - quyi
            if len(slots) >= 2 and slots[0] and not slots[1]:
                sub = slots[0]
            elif len(slots) >= 2 and slots[1] and not slots[0]:
                sup = slots[1]
            elif filled:
                if variation == 1:
                    sub = filled[0]
                else:
                    sup = filled[0]
        out = ""
        if not before:
            out = "{}"
        if sub:
            out += "_{%s}" % sub
        if sup:
            out += "^{%s}" % sup
        return out


def mtef_to_latex(data: bytes) -> str:
    """Xom MTEF baytlaridan LaTeX."""
    try:
        return MtefParser(data).convert()
    except MtefError:
        raise
    except (IndexError, struct.error, RecursionError, UnicodeError) as error:
        raise MtefError(str(error)) from error


def ole_equation_to_latex(blob: bytes) -> str:
    """OLE obyekt (oleObjectN.bin) ichidagi formulani LaTeX'ga o'giradi."""
    import olefile

    try:
        ole = olefile.OleFileIO(io.BytesIO(blob))
    except Exception as error:  # noqa: BLE001 - olefile turli xatolar otadi
        raise MtefError(f"OLE fayl o'qilmadi: {error}") from error

    try:
        if not ole.exists("Equation Native"):
            raise MtefError("obyektda formula oqimi yo'q")
        stream = ole.openstream("Equation Native").read()
    finally:
        ole.close()

    if len(stream) < 28:
        raise MtefError("formula oqimi juda qisqa")
    header_size = struct.unpack_from("<H", stream, 0)[0]
    if not 4 <= header_size < len(stream):
        header_size = 28
    latex = mtef_to_latex(stream[header_size:])
    if not latex:
        raise MtefError("formula bo'sh chiqdi")
    return latex
