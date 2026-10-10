"""Word(.docx)の段落を、自動番号(第○条・第○章・２．など)を文字にした状態で取り出す。

python-docx は自動番号を文字として返さないので、numbering.xml を読んで自前で番号を振る。
"""
import re
import docx
from docx.oxml.ns import qn

FW = str.maketrans("0123456789", "０１２３４５６７８９")
KANJI = "〇一二三四五六七八九"
AIUEO = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン"
IROHA = "イロハニホヘトチリヌルヲワカヨタレソツネナラムウヰノオクヤマケフコエテアサキユメミシヱヒモセスン"
CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"


def kanji_num(n):
    if n < 10:
        return KANJI[n]
    if n < 100:
        t, o = divmod(n, 10)
        return ("" if t == 1 else KANJI[t]) + "十" + (KANJI[o] if o else "")
    h, r = divmod(n, 100)
    return ("" if h == 1 else KANJI[h]) + "百" + (kanji_num(r) if r else "")


def fmt(n, f):
    if f in ("decimalFullWidth", "decimalFullWidth2"):
        return str(n).translate(FW)
    if f in ("japaneseCounting", "japaneseDigitalTenThousand", "ideographTraditional"):
        return kanji_num(n)
    if f in ("aiueoFullWidth", "aiueo"):
        return AIUEO[(n - 1) % len(AIUEO)]
    if f in ("irohaFullWidth", "iroha"):
        return IROHA[(n - 1) % len(IROHA)]
    if f == "decimalEnclosedCircle":
        return CIRCLED[n - 1] if n <= 20 else str(n)
    if f == "lowerLetter":
        return chr(ord("a") + (n - 1) % 26)
    if f == "upperLetter":
        return chr(ord("A") + (n - 1) % 26)
    if f == "lowerRoman":
        return roman(n).lower()
    if f == "upperRoman":
        return roman(n)
    if f in ("bullet", "none"):
        return ""
    return str(n)


def roman(n):
    vals = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
            (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    s = ""
    for v, r in vals:
        while n >= v:
            s += r
            n -= v
    return s


class Numbering:
    def __init__(self, document):
        self.abstract = {}   # abstractNumId -> {ilvl: (start, numFmt, lvlText, restart)}
        self.nums = {}       # numId -> (abstractNumId, {ilvl: startOverride})
        self.counters = {}   # abstractNumId -> [count per level]
        self.started = set()
        try:
            root = document.part.numbering_part.element
        except Exception:
            root = None
        if root is None:
            return
        for an in root.findall(qn("w:abstractNum")):
            aid = an.get(qn("w:abstractNumId"))
            lv = {}
            for l in an.findall(qn("w:lvl")):
                il = int(l.get(qn("w:ilvl")))
                g = lambda tag: (l.find(qn(tag)).get(qn("w:val")) if l.find(qn(tag)) is not None else None)
                lv[il] = (int(g("w:start") or 1), g("w:numFmt") or "decimal", g("w:lvlText") or "", g("w:lvlRestart"))
            self.abstract[aid] = lv
        for n in root.findall(qn("w:num")):
            nid = n.get(qn("w:numId"))
            a = n.find(qn("w:abstractNumId"))
            ov = {}
            for o in n.findall(qn("w:lvlOverride")):
                so = o.find(qn("w:startOverride"))
                if so is not None:
                    ov[int(o.get(qn("w:ilvl")))] = int(so.get(qn("w:val")))
            self.nums[nid] = (a.get(qn("w:val")) if a is not None else None, ov)

    def label(self, numId, ilvl):
        if numId not in self.nums or numId == "0":
            return ""
        aid, ov = self.nums[numId]
        lv = self.abstract.get(aid)
        if not lv or ilvl not in lv:
            return ""
        cnt = self.counters.setdefault(aid, [0] * 10)
        if numId not in self.started:
            self.started.add(numId)
            for il, s in ov.items():
                cnt[il] = s - 1
        if cnt[ilvl] == 0 and ilvl not in ov:
            cnt[ilvl] = lv[ilvl][0] - 1
        cnt[ilvl] += 1
        for deeper in range(ilvl + 1, 10):
            cnt[deeper] = 0
        start, f, text, _ = lv[ilvl]

        def rep(m):
            k = int(m.group(1)) - 1
            if k not in lv:
                return ""
            c = cnt[k] if cnt[k] else lv[k][0]
            return fmt(c, lv[k][1])
        return re.sub(r"%(\d)", rep, text)


def para_numpr(p, styles_numpr):
    pPr = p._p.pPr
    if pPr is not None and pPr.numPr is not None:
        nid = pPr.numPr.numId
        il = pPr.numPr.ilvl
        return (str(nid.val) if nid is not None else None, int(il.val) if il is not None else 0)
    sid = p.style.style_id if p.style is not None else None
    return styles_numpr.get(sid)


def style_numpr_map(document):
    out = {}
    for st in document.styles:
        try:
            el = st.element
        except Exception:
            continue
        pPr = el.find(qn("w:pPr"))
        if pPr is None:
            continue
        np = pPr.find(qn("w:numPr"))
        if np is None:
            continue
        nid = np.find(qn("w:numId"))
        il = np.find(qn("w:ilvl"))
        if nid is not None:
            out[st.style_id] = (nid.get(qn("w:val")), int(il.get(qn("w:val"))) if il is not None else 0)
    return out


def iter_blocks(document):
    """本文を順番に返す。('p', label, text, style) か ('table', rows)"""
    num = Numbering(document)
    snp = style_numpr_map(document)
    body = document.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            p = docx.text.paragraph.Paragraph(child, document)
            text = p.text.strip()
            np = para_numpr(p, snp)
            label = num.label(*np) if np and np[0] else ""
            if not text and not label:
                continue
            style = p.style.name if p.style is not None else ""
            yield ("p", label.strip(), text, style)
        elif child.tag == qn("w:tbl"):
            t = docx.table.Table(child, document)
            rows = []
            for r in t.rows:
                cells = []
                prev = None
                for c in r.cells:
                    if c._tc is prev:
                        continue
                    prev = c._tc
                    cells.append(" ".join(x.text.strip() for x in c.paragraphs if x.text.strip()))
                rows.append(cells)
            yield ("table", rows)
