"""원고 읽기 — 텍스트 추출(PDF·DOCX·HWP·TXT) → 절·캡션·참고문헌 구조 → 기계적 점검(LLM 없음).

  doc = load(path)            # {"name","lang","text","pages":[시작 offset],"sections","captions","refs",...}
  issues = mechanical(doc)    # [{"sev","cat","title","quote","pos","len","why","fix","src":"rule"}]

모든 위치는 doc["text"] 안의 글자 offset 이다. 쪽은 page_of(doc, pos), 절은 section_of(doc, pos).
"""
import bisect
import os
import re
import shutil
import subprocess
import zipfile
from xml.etree import ElementTree as ET

ROOT = os.path.dirname(os.path.abspath(__file__))


# ── 추출 ────────────────────────────────────────────────────────────────
def _kordoc_cli():
    """같은 저작자의 kordoc-local / notebook-local 에 설치된 kordoc 을 빌려 쓴다(HWP·HWPX 용)."""
    cands = [os.environ.get("KORDOC_CLI", "")] + [os.path.join(ROOT, "..", d, "node_modules", "kordoc", "dist", "cli.js")
                                                 for d in ("review-local", "kordoc-local", "notebook-local")]
    for c in cands:
        if c and os.path.exists(c) and shutil.which("node"):
            return os.path.abspath(c)
    return None


def tools_status():
    return {"pdftotext": bool(shutil.which("pdftotext")), "kordoc": bool(_kordoc_cli())}


def _pdf_pages(path):
    if shutil.which("pdftotext"):
        r = subprocess.run(["pdftotext", "-enc", "UTF-8", path, "-"], capture_output=True, timeout=300)
        if r.returncode != 0:
            raise RuntimeError("PDF 텍스트 추출 실패: " + r.stderr.decode(errors="replace")[-300:])
        pages = r.stdout.decode("utf-8", errors="replace").split("\f")
        if pages and not pages[-1].strip():
            pages.pop()
        return pages
    cli = _kordoc_cli()
    if cli:
        return [_kordoc(cli, path)]
    raise RuntimeError("PDF 를 읽을 도구가 없습니다 (poppler-utils 의 pdftotext 설치 또는 kordoc-local 필요)")


def _kordoc(cli, path):
    out = path + ".md"
    r = subprocess.run(["node", cli, "--silent", path, "-o", out], capture_output=True, text=True, timeout=600,
                       cwd=os.path.dirname(path))
    if r.returncode != 0 or not os.path.exists(out):
        raise RuntimeError("kordoc 파싱 실패: " + (r.stdout + r.stderr)[-300:])
    with open(out, encoding="utf-8") as f:
        md = f.read()
    os.remove(out)
    # 마크다운 표·강조 기호를 걷어 본문처럼 만든다(헤딩 # 은 절 인식에 쓰므로 둔다)
    md = re.sub(r"^\s*\|?\s*:?-{3,}.*$", "", md, flags=re.M)
    md = re.sub(r"\*\*(.+?)\*\*", r"\1", md)
    return md.replace(" | ", "  ").replace("|", " ")


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _docx_pages(path):
    """stdlib 로 DOCX 문단을 읽는다. 쪽은 Word 가 저장해 둔 쪽 나눔 표시(lastRenderedPageBreak)로 근사."""
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    pages, cur = [], []
    for p in root.iter(W + "p"):
        buf = []
        for el in p.iter():
            if el.tag == W + "lastRenderedPageBreak" or (el.tag == W + "br" and el.get(W + "type") == "page"):
                if cur or buf:
                    cur.append("".join(buf))
                    pages.append("\n".join(cur))
                    cur, buf = [], []
            elif el.tag == W + "t" and el.text:
                buf.append(el.text)
            elif el.tag == W + "tab":
                buf.append("\t")
        style = p.find(f"{W}pPr/{W}pStyle")
        line = "".join(buf)
        if style is not None and re.match(r"(?i)heading|제목", style.get(W + "val") or "") and line.strip():
            line = "\n" + line  # 헤딩 앞 빈 줄
        cur.append(line)
    pages.append("\n".join(cur))
    return [p for p in pages if p.strip()] or [""]


def extract_pages(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return _pdf_pages(path), "pdftotext" if shutil.which("pdftotext") else "kordoc"
    if ext == ".docx":
        return _docx_pages(path), "docx"
    if ext in (".hwp", ".hwpx"):
        cli = _kordoc_cli()
        if not cli:
            raise RuntimeError("HWP 를 읽으려면 kordoc-local(node_modules/kordoc)이 필요합니다. PDF 로 저장해 올려 주세요.")
        return [_kordoc(cli, path)], "kordoc"
    if ext in (".txt", ".md", ".tex"):
        with open(path, "rb") as f:
            raw = f.read()
        for enc in ("utf-8-sig", "cp949"):
            try:
                txt = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            txt = raw.decode("utf-8", errors="replace")
        return txt.split("\f"), "text"
    raise RuntimeError(f"지원하지 않는 형식: {ext} (PDF·DOCX·HWP·HWPX·TXT)")


# ── 정리 ────────────────────────────────────────────────────────────────
LIG = str.maketrans({"ﬁ": "fi", "ﬂ": "fl", "ﬀ": "ff", "ﬃ": "ffi", "ﬄ": "ffl", "­": "", " ": " ", " ": " ", " ": " "})


def clean_pages(pages):
    """쪽 번호 줄·반복 머리말/꼬리말 제거, 합자 풀기. 줄 구조는 유지한다."""
    pages = [p.translate(LIG).replace("\r\n", "\n").replace("\r", "\n") for p in pages]
    key = lambda s: re.sub(r"\d+", "#", s.strip().lower())
    edge = {}
    for p in pages:
        lines = [l for l in p.split("\n") if l.strip()]
        for l in set(lines[:2] + lines[-2:]):
            edge[key(l)] = edge.get(key(l), 0) + 1
    rep = {k for k, c in edge.items() if len(pages) >= 4 and c >= max(3, len(pages) * 0.5) and len(k) < 120 and re.search(r"[^\W\d_#]", k)}
    out = []
    for p in pages:
        lines = p.split("\n")
        nz = [i for i, l in enumerate(lines) if l.strip()]
        drop = {i for i in nz[:2] + nz[-2:] if key(lines[i]) in rep}
        for i in reversed(nz[-2:]):  # 맨 끝 쪽 번호 ("7", "- 7 -", "Page 7")
            if re.fullmatch(r"\s*(?:-\s*)?(?:page\s*)?\d{1,3}(?:\s*-)?\s*(?:/\s*\d+)?\s*", lines[i], re.I):
                drop.add(i)
                break
        out.append("\n".join(l for i, l in enumerate(lines) if i not in drop).strip("\n"))
    return out


def detect_lang(text):
    ko = len(re.findall(r"[가-힣]", text))
    en = len(re.findall(r"[A-Za-z]", text))
    return "ko" if ko > en * 0.3 else "en"


# ── 절 구조 ─────────────────────────────────────────────────────────────
KINDS = [  # (kind, 정규식) — 제목 글자만으로 판별
    ("abstract", r"abstract|summary|요\s*약|초\s*록|국문\s*초록|영문\s*초록"),
    ("intro", r"introduction|background|서\s*론|머리말|들어가며|연구\s*배경"),
    ("related", r"related\s+works?|literature\s+review|prior\s+work|관련\s*연구|선행\s*연구"),
    ("method", r"methods?|methodology|materials?\s+and\s+methods?|experimental(?:\s+(?:setup|methods?|procedure))?|approach|proposed|model|system|framework|algorithm|data(?:set|sets)?|연구\s*방법|실험\s*방법|방법|제안|모델|시스템|데이터|재료"),
    ("results", r"results?(?:\s+and\s+discussions?)?|experiments?|evaluation|performance|analysis|결과|실험|평가|분석"),
    ("discussion", r"discussions?|limitations?|고찰|논의|토의|한계"),
    ("conclusion", r"conclusions?|concluding\s+remarks|summary\s+and\s+conclusions?|결론|맺음말|마치며|요약\s*및\s*결론"),
    ("ack", r"acknowledge?ments?|funding|감사의\s*글|사사"),
    ("references", r"references|bibliography|works\s+cited|literature\s+cited|참\s*고\s*문\s*헌|인용\s*문헌"),
    ("appendix", r"appendix|appendices|supplementary|부\s*록"),
]
KIND_LABEL = {"front": "제목·저자", "abstract": "초록", "intro": "서론", "related": "관련 연구", "method": "방법", "results": "결과",
              "discussion": "논의", "conclusion": "결론", "ack": "감사의 글", "references": "참고문헌", "appendix": "부록", "body": "본문"}
ROMAN = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7, "VIII": 8, "IX": 9, "X": 10, "XI": 11, "XII": 12,
         "Ⅰ": 1, "Ⅱ": 2, "Ⅲ": 3, "Ⅳ": 4, "Ⅴ": 5, "Ⅵ": 6, "Ⅶ": 7, "Ⅷ": 8, "Ⅸ": 9, "Ⅹ": 10}
UNNUMBERED = re.compile(r"^\s*(?:#+\s*)?(abstract|summary|요\s*약|초\s*록|국문\s*초록|영문\s*초록|keywords?|key\s*words|주제어|"
                        r"acknowledge?ments?|funding|감사의\s*글|references|bibliography|참\s*고\s*문\s*헌|인용\s*문헌|"
                        r"appendix(?:\s+[A-Z])?|appendices|부\s*록|conclusions?|introduction|서\s*론|결\s*론|discussion)\s*[:.]?\s*$", re.I)
NUM_HEAD = re.compile(r"^\s*(?:#+\s*)?(?:(\d{1,2}(?:\.\d{1,2}){0,3})\.?|([IVX]{1,4}|[Ⅰ-Ⅹ])\.|제\s*(\d{1,2})\s*장)\s+(\S.{0,90}?)\s*$")
NUM_ONLY = re.compile(r"^\s*(\d{1,2}(?:\.\d{1,2}){0,3})\.?\s*$|^\s*([IVX]{1,4}|[Ⅰ-Ⅹ])\.\s*$")
MD_HEAD = re.compile(r"^\s*(#{1,4})\s+(\S.{0,100}?)\s*$")


def kind_of(title):
    t = re.sub(r"^(?:[\d.]+|[IVXⅠ-Ⅹ]+\.)\s*", "", title.strip()).strip().lower()
    t = re.sub(r"^제\s*\d+\s*장\s*", "", t)
    for k, rx in KINDS:
        if re.match(rf"(?:{rx})\b", t, re.I) or re.match(rf"(?:{rx})", t) and re.search(r"[가-힣]", t[:2]):
            return k
    return "body"


def _title_like(s):
    s = s.strip()
    if not s or len(s) > 90 or s.endswith((",", ";")) or re.search(r"[=<>∑∫·×^]|\(\w*\d", s) or not re.search(r"[A-Za-z]{3}|[가-힣]{2}", s):
        return False
    if re.search(r"[가-힣]", s):
        return len(s) <= 40 and not re.search(r"(다|요|음)\.$", s)
    words = s.split()
    return len(words) <= 12 and s[0].isupper() and not re.search(r"\.\s+[A-Z]", s) and not s.endswith(".") or bool(UNNUMBERED.match(s))


def _num_key(num, roman):
    if roman:
        return (ROMAN.get(roman, 0),)
    return tuple(int(x) for x in num.split("."))


def _next_ok(prev, cur):
    """번호가 붙은 헤딩은 앞 헤딩의 '다음 번호'일 때만 인정 (표 안 숫자·쪽 번호 오인 방지)"""
    if not prev:
        return cur[0] in (1, 2) or len(cur) == 1 and cur[0] <= 2
    for d in range(min(len(prev), len(cur))):
        if cur[d] != prev[d]:
            return cur[d] == prev[d] + 1 and all(x == 1 for x in cur[d + 1:])
    return len(cur) == len(prev) + 1 and cur[-1] == 1 or (len(cur) > len(prev) and all(x == 1 for x in cur[len(prev):]))


def find_sections(text):
    """[{title, kind, start, end, level}] — start 는 헤딩 줄 시작 offset. 첫 헤딩 앞은 front."""
    heads, prev_num, roman_mode = [], None, None
    lines = text.split("\n")
    offs, o = [], 0
    for l in lines:
        offs.append(o)
        o += len(l) + 1
    i = 0
    while i < len(lines):
        line = lines[i]
        s = line.strip()
        prev_blank = i == 0 or not lines[i - 1].strip() or lines[i - 1].rstrip().endswith((".", ":", "다.")) or len(lines[i - 1].strip()) < 50
        hit = None
        m = MD_HEAD.match(line)
        if m and not re.match(r"^\s*#+\s*\d+\s*$", line):
            title = m.group(2).strip()
            nm = NUM_HEAD.match(title)
            hit = (title, len(m.group(1)), _num_key(nm.group(1), None) if nm and nm.group(1) else None, i)
        elif UNNUMBERED.match(s) and prev_blank:
            hit = (s.rstrip(":."), 1, None, i)
        else:
            m = NUM_HEAD.match(line)
            mo = NUM_ONLY.match(line) if not m else None
            if mo and i + 1 < len(lines):  # "3.1" 다음 줄(빈 줄 하나 건너뛸 수 있음)이 제목인 pdftotext 배치
                j = i + 1 + (1 if i + 2 < len(lines) and not lines[i + 1].strip() else 0)
                if _title_like(lines[j]):
                    num, rom = mo.group(1), mo.group(2)
                    key = _num_key(num, rom)
                    if (rom and (roman_mode is not False) and (not prev_num or key[0] == prev_num[0] + 1)) or (num and _next_ok(prev_num, key)):
                        hit = (f"{num or rom} {lines[j].strip()}", len(key), key, i)
                        if rom:
                            roman_mode = True
                        i = j
            elif m and prev_blank:
                num, rom, chap, title = m.group(1), m.group(2), m.group(3), m.group(4)
                if _title_like(title) and not re.match(r"^[\d.,%)]", title):
                    key = _num_key(num or chap, rom)
                    if (rom and roman_mode is not False and (not prev_num or key[0] == prev_num[0] + 1)) or ((num or chap) and _next_ok(prev_num, key)):
                        hit = (s, len(key), key, i)
        if hit:
            title, level, key, li = hit
            if key:
                prev_num = key
                if roman_mode is None and not re.match(r"^[IVXⅠ-Ⅹ]+\.", title):
                    roman_mode = False
            heads.append({"title": re.sub(r"\s+", " ", title.lstrip("#").strip())[:100], "start": offs[li], "level": level})
        i += 1
    secs = [{"title": "", "kind": "front", "start": 0, "level": 0}] if not heads or heads[0]["start"] > 0 else []
    last_kind = None
    for h in heads:
        k = kind_of(h["title"])
        if last_kind and h["level"] > 1 and k not in ("references", "ack", "appendix"):
            k = last_kind  # 하위 절은 상위 절의 종류를 잇는다 (3.2 Multi-Head Attention → 방법, 6.2 Model Variations → 결과)
        if k == "body" and last_kind in ("references",):
            k = "appendix"
        h["kind"] = k
        if h["level"] <= 1 or last_kind is None:
            last_kind = k
        secs.append(h)
    # 초록 헤딩이 없으면: 첫 헤딩 앞 블록에서 가장 긴 문단을 초록으로 본다
    if not any(s["kind"] == "abstract" for s in secs) and secs and secs[0]["kind"] == "front":
        front_end = secs[1]["start"] if len(secs) > 1 else len(text)
        paras = [(m.start(), m.end()) for m in re.finditer(r"(?:[^\n]+\n?)+", text[:front_end])]
        if paras:
            a, b = max(paras, key=lambda p: p[1] - p[0])
            if b - a > 400:
                secs.insert(1, {"title": "(초록으로 추정)", "kind": "abstract", "start": a, "level": 1})
    for n, s in enumerate(secs):
        s["end"] = secs[n + 1]["start"] if n + 1 < len(secs) else len(text)
    return secs


# ── 캡션·언급 ───────────────────────────────────────────────────────────
_NUMTOK = r"(?:\d+|[IVX]{1,5})"
CAPTION = re.compile(r"^\s*(?:<|\[|\()?\s*(Fig(?:ure)?\.?|FIG(?:URE)?\.?|Table|TABLE|그림|표)\s*(" + _NUMTOK + r")(?:\s*>|\s*\]|\s*\))?"
                     r"\s*(?:[.:|]|\s[-–—]\s|$|\s+(?=[A-Z가-힣(]))(.*)$")
MENTION = re.compile(r"(?<![A-Za-z가-힣])(Fig(?:ure)?s?\.?|FIG(?:URE)?S?\.?|Tables?|TABLES?|그림|표)\s*(" + _NUMTOK + r"[a-z]?"
                     r"(?:\s*\([a-z]\))?(?:\s*(?:,|and|&|to|or|–|—|-|~|및|와|과)\s*" + _NUMTOK + r"[a-z]?)*)")


def _fignum(s):
    s = s.strip()
    return int(s) if s.isdigit() else ROMAN.get(s, 0)


def _kind(label):
    return "fig" if re.match(r"(?i)fig|그림", label) else "tab"


def _expand(spec):
    """'2–4' → [2,3,4], '1, 3 and 5' → [1,3,5], '2a' → [2]"""
    spec = re.sub(r"\([a-z]\)", "", spec)
    nums, out = re.findall(r"\d+|[IVX]{1,5}", spec), []
    vals = [_fignum(n) for n in nums]
    for k, v in enumerate(vals):
        out.append(v)
        if k and re.search(rf"{nums[k - 1]}[a-z]?\s*(?:–|—|-|~|to)\s*{nums[k]}", spec) and 0 < v - vals[k - 1] <= 30:
            out += list(range(vals[k - 1] + 1, v))
    return sorted(set(x for x in out if x > 0))


def find_captions(text, secs):
    caps, lines, o = [], text.split("\n"), 0
    for i, line in enumerate(lines):
        m = CAPTION.match(line)
        if m:
            prev = lines[i - 1].strip() if i else ""
            body = m.group(3).strip()
            short_label = re.match(r"^\s*(?:Table|TABLE|표)\s*\d+\s*$", line)  # 단독 "Table 3" 줄
            # 본문 문장이 줄바꿈으로 "Table 3. ..." 처럼 시작한 경우 제외: 앞 줄이 문장 중간(구두점 없이 끝남)이면 캡션 아님
            if (not prev or prev.endswith((".", ":", "!", "?", ")")) or len(prev) < 40 or short_label) and \
                    (body or i + 1 < len(lines)) and not re.match(r"^(shows?|presents?|lists?|summarizes?|illustrates?|depicts?|is|are|에서|은|는|과|와|에)\b", body):
                caps.append({"type": _kind(m.group(1)), "num": _fignum(m.group(2)), "label": f"{m.group(1)} {m.group(2)}",
                             "text": (body or lines[i + 1].strip())[:240], "pos": o, "len": len(line)})
        o += len(line) + 1
    # 같은 번호가 여럿이면 첫 것(콜론이 있는 것 우선)만
    seen, out = {}, []
    for c in caps:
        k = (c["type"], c["num"])
        if k not in seen:
            seen[k] = c
            out.append(c)
    return out


def find_mentions(text, caps, lo, hi):
    """본문(lo~hi) 의 그림·표 언급 [(type, num, pos, len)] — 캡션 줄 자체는 뺀다"""
    cap_spans = [(c["pos"], c["pos"] + 12) for c in caps]
    out = []
    for m in MENTION.finditer(text, lo, hi):
        if any(a <= m.start() < b for a, b in cap_spans):
            continue
        if re.match(r"\s*S\d", text[m.end(1):m.end(1) + 3]):  # 보충 자료 Fig. S1
            continue
        if m.group(1) == "표" and not re.match(r"\s*\d", m.group(2)):
            continue
        for n in _expand(m.group(2)):
            out.append((_kind(m.group(1)), n, m.start(), m.end() - m.start()))
    return out


# ── 참고문헌·인용 ───────────────────────────────────────────────────────
def find_refs(text, secs):
    ref = next((s for s in secs if s["kind"] == "references"), None)
    if not ref:
        return [], None
    body = text[ref["start"]:ref["end"]]
    base = ref["start"]
    first_nl = body.find("\n")
    entries = []
    for m in re.finditer(r"(?m)^\s*(?:\[(\d{1,3})\]|(\d{1,3})\.(?=\s+[A-Z가-힣\"“])|(\d{1,3})\)\s)", body):
        n = int(m.group(1) or m.group(2) or m.group(3))
        entries.append({"n": n, "pos": base + m.start()})
    style = None
    if len(entries) >= 3:
        style = "numeric"
        for k, e in enumerate(entries):
            end = entries[k + 1]["pos"] if k + 1 < len(entries) else ref["end"]
            e["text"] = re.sub(r"\s+", " ", text[e["pos"]:end]).strip()[:400]
            e["len"] = min(end - e["pos"], 400)
    else:
        # 저자-연도: 줄 첫머리가 '성, 이니셜' 이고 연도가 있는 문단
        entries = []
        chunk = body[first_nl + 1:] if first_nl >= 0 else ""
        for m in re.finditer(r"(?m)^(?:[A-Z][A-Za-z'\-]+(?: [A-Z][A-Za-z'\-]+)?,\s*(?:[A-Z]\.|[A-Z][a-z]+)|[가-힣]{2,4}(?:,|\s))[^\n]*", chunk):
            entries.append({"n": None, "pos": base + first_nl + 1 + m.start()})
        for k, e in enumerate(entries):
            end = entries[k + 1]["pos"] if k + 1 < len(entries) else ref["end"]
            e["text"] = re.sub(r"\s+", " ", text[e["pos"]:end]).strip()[:400]
            e["len"] = min(end - e["pos"], 400)
        entries = [e for e in entries if re.search(r"(19|20)\d\d", e["text"])]
        style = "author-year" if len(entries) >= 3 else None
    return entries, style


CITE_NUM = re.compile(r"\[(\d{1,3}(?:\s*[-–—,]\s*\d{1,3})*)\]")
CITE_AY = re.compile(r"([A-Z][A-Za-z'\-]+)(?:\s+(?:et\s+al\.?|and|&)\s*(?:[A-Z][A-Za-z'\-]+)?)?,?\s*\(?((?:19|20)\d\d)[a-z]?\)?")


def _cite_nums(spec):
    out = []
    for part in re.split(r"\s*,\s*", spec):
        a = re.split(r"\s*[-–—]\s*", part)
        if len(a) == 2 and a[0].isdigit() and a[1].isdigit() and 0 < int(a[1]) - int(a[0]) <= 40:
            out += list(range(int(a[0]), int(a[1]) + 1))
        elif part.strip().isdigit():
            out.append(int(part))
    return out


def guess_title(front):
    """제목·저자 블록에서 제목: 1~2줄짜리 문단, 3~30단어, 마침표로 끝나지 않음, 이메일·저작권 문구 아님"""
    cands = []
    for para in re.split(r"\n\s*\n", front):
        lines = [l.strip() for l in para.split("\n") if l.strip()]
        if 1 <= len(lines) <= 3:
            cands.append(" ".join(lines))
        if len(lines) > 1:
            cands.append(lines[0])
    for t in cands:
        if re.search(r"@|arxiv|doi|©|copyright|permission|license|journal|vol\.|https?:|university|institute|연구원|대학교", t, re.I):
            continue
        n = len(t.split())
        if (2 <= n <= 30 or re.search(r"[가-힣]", t) and 6 <= len(t) <= 120) and not t.endswith((".", ",")) and not re.match(r"^[\d\W]", t):
            return t[:200]
    return ""


# ── 읽기 진입점 ─────────────────────────────────────────────────────────
def build(pages, name="", via=""):
    pages = clean_pages(pages)
    starts, parts, o = [], [], 0
    for p in pages:
        starts.append(o)
        parts.append(p)
        o += len(p) + 2
    text = "\n\n".join(parts)
    secs = find_sections(text)
    caps = find_captions(text, secs)
    refs, ref_style = find_refs(text, secs)
    lang = detect_lang(text)
    title = guess_title(text[:secs[1]["start"] if len(secs) > 1 else 3000])
    return {"name": name, "via": via, "lang": lang, "title": title, "text": text, "pages": starts, "npages": len(pages),
            "sections": secs, "captions": caps, "refs": refs, "ref_style": ref_style}


def load(path, name=None):
    pages, via = extract_pages(path)
    doc = build(pages, name or os.path.basename(path), via)
    if len(re.sub(r"\s", "", doc["text"])) < 200:
        raise RuntimeError("추출된 글자가 거의 없습니다. 스캔(이미지) PDF 이면 OCR 한 PDF 나 원본 DOCX·HWP 를 올려 주세요.")
    return doc


def page_of(doc, pos):
    if doc["npages"] <= 1:
        return None
    return bisect.bisect_right(doc["pages"], pos)


def section_of(doc, pos):
    s = None
    for x in doc["sections"]:
        if x["start"] <= pos:
            s = x
    return s


def where(doc, pos):
    s, p = section_of(doc, pos), page_of(doc, pos)
    name = (s["title"] or KIND_LABEL.get(s["kind"], "")) if s else ""
    if s and s["kind"] == "references" and doc["refs"] and pos > doc["refs"][-1]["pos"] + doc["refs"][-1]["len"]:
        name = "참고문헌 뒤(부록)"
    return ((f"p.{p} · " if p else "") + name).strip(" ·")


def region(doc, *kinds):
    return [(s["start"], s["end"]) for s in doc["sections"] if s["kind"] in kinds]


def body_span(doc):
    """초록부터 참고문헌 전까지(제목·저자 블록, 참고문헌 이후 제외)"""
    secs = doc["sections"]
    lo = next((s["start"] for s in secs if s["kind"] != "front"), 0)
    hi = next((s["start"] for s in secs if s["kind"] in ("references",)), len(doc["text"]))
    return lo, hi


def sentence_at(text, pos, ln=0, maxlen=220):
    """pos 를 포함한 문장(줄바꿈은 공백으로)을 (quote, start, len) 로"""
    a = max(text.rfind(". ", 0, pos), text.rfind(".\n", 0, pos), text.rfind("\n\n", 0, pos), text.rfind("다. ", 0, pos))
    a = 0 if a < 0 else a + 2
    nl = text.rfind("\n", 0, pos)
    if nl >= a:  # 앞 줄이 짧거나(제목) 문장이 끝난 줄이면 줄 경계에서 시작
        pl = text[text.rfind("\n", 0, nl) + 1:nl].strip()
        if len(pl) < 50 or re.search(r"[.!?:다)]$", pl):
            a = nl + 1
    if pos - a > maxlen // 2:
        a = pos - maxlen // 2
        a = text.find(" ", a) + 1 or a
    m = re.compile(r"\.(?:\s|$)|\n\n").search(text, pos + max(ln, 1))
    b = m.start() + 1 if m else len(text)
    b = min(b, a + maxlen)
    return re.sub(r"\s+", " ", text[a:b]).strip(), a, b - a


# ── 인용(LLM 지적의 근거 문장) 찾기 ─────────────────────────────────────
QMAP = str.maketrans({"‘": "'", "’": "'", "“": '"', "”": '"', "–": "-", "—": "-", "−": "-", "…": "...", "·": ".", "×": "x"})


def compact(text):
    """공백 없애고 소문자·따옴표·대시 통일한 문자열과, 각 글자의 원래 offset"""
    t = text.translate(LIG).translate(QMAP).lower()
    chars, idx = [], []
    for i, ch in enumerate(t):
        if ch.isspace() or ch in "\"'`*":
            continue
        chars.append(ch)
        idx.append(i)
    return "".join(chars), idx


def locate(doc, quote, lo=0, hi=None, cache=None):
    """quote 를 원고에서 찾아 (start, len, 정확도) — 못 찾으면 None. lo~hi 구간 우선, 없으면 전체."""
    import difflib
    if cache is None or "c" not in cache:
        c, idx = compact(doc["text"])
        if cache is not None:
            cache.update(c=c, idx=idx)
    else:
        c, idx = cache["c"], cache["idx"]
    pieces = [p for p in re.split(r"\.\.\.|…|\[\.\.\.\]", quote) if len(compact(p)[0]) >= 8]
    if not pieces:
        return None
    best = None
    for piece in sorted(pieces, key=len, reverse=True)[:2]:
        q = compact(piece)[0]
        if len(q) < 8:
            continue
        hi_c = len(c) if hi is None else bisect.bisect_left(idx, hi)
        lo_c = bisect.bisect_left(idx, lo)
        k = c.find(q, lo_c, hi_c)
        if k < 0:
            k = c.find(q)
        if k >= 0:
            best = (k, len(q), 1.0)
            break
        # 부분 일치: 구간 안에서 가장 긴 공통 조각
        seg_lo, seg_hi = (lo_c, hi_c) if hi_c - lo_c > len(q) else (0, len(c))
        sm = difflib.SequenceMatcher(None, c[seg_lo:seg_hi], q, autojunk=False)
        m = sm.find_longest_match(0, seg_hi - seg_lo, 0, len(q))
        if m.size >= max(15, 0.25 * len(q)):
            # 가장 긴 공통 조각을 기준으로 정렬한 창 전체의 유사도 — 낱말 한두 개 빠뜨린 인용은 살리고, 지어낸 문장은 버린다
            start = max(0, seg_lo + m.a - m.b)
            ratio = difflib.SequenceMatcher(None, c[start:start + len(q) + 8], q, autojunk=False).ratio()
            if ratio >= 0.85 and (not best or ratio > best[2]):
                best = (start, len(q), ratio)
    if not best:
        return None
    k, n, r = best
    a = idx[min(k, len(idx) - 1)]
    b = idx[min(k + n - 1, len(idx) - 1)] + 1
    return a, b - a, r


# ── 기계적 점검 ─────────────────────────────────────────────────────────
WHITELIST = set("""US USA UK EU UN OK ID PDF URL URLS HTTP HTTPS HTML XML JSON CSV API GPU GPUS CPU CPUS RAM ROM USB PC PCS IT TV DNA RNA
PHD MSC BSC CEO CTO ISO IEC IEEE ACM ASME ANS IAEA KAERI NASA DOE NRC KINS KHNP SI MATLAB NIPS NEURIPS ICML ICLR CVPR ECCV ICCV AAAI IJCAI
ACL EMNLP NAACL ARXIV DOI ISBN ISSN II III IV VI VII VIII IX XI XII XIII XIV XV AM PM UTC GMT AD BC TM QED RGB CMYK FAQ PS NB CC BY ND NC
OR AND NOT XOR IF ETC LTD INC CO VS IE EG AI NVIDIA IBM AMD INTEL FLOPS GFLOPS TFLOPS PFLOPS FLOP MACS""".split())
SI_UNITS = r"(?:nm|μm|µm|mm|cm|km|kg|mg|μg|ms|μs|ns|Hz|kHz|MHz|GHz|kW|MW|GW|mW|kV|mV|MeV|keV|GeV|eV|kPa|MPa|GPa|Pa|mSv|μSv|Sv|Bq|GBq|MBq|Gy|mGy|mL|ml|mol|mmol|wt%|at%|dpa)"


def _issue(doc, sev, cat, title, pos, ln, why, fix, quote=None):
    q = quote if quote is not None else sentence_at(doc["text"], pos, ln)[0]
    return {"sev": sev, "cat": cat, "title": title, "quote": q, "pos": pos, "len": ln, "why": why, "fix": fix, "src": "rule",
            "where": where(doc, pos), "page": page_of(doc, pos)}


def check_figures(doc):
    text, caps, out = doc["text"], doc["captions"], []
    lo, hi = body_span(doc)
    ments = find_mentions(text, caps, lo, hi)
    ko = doc["lang"] == "ko"
    for typ, name in (("fig", "그림"), ("tab", "표")):
        c = {x["num"]: x for x in caps if x["type"] == typ}
        m = [x for x in ments if x[0] == typ]
        mentioned = {x[1] for x in m}
        if not c and m:
            out.append(_issue(doc, "info", f"{name} 번호", f"{name} 캡션을 찾지 못함", m[0][2], m[0][3],
                              f"본문에서 {name} {len(mentioned)}개 번호를 언급하지만 '{('그림 1.' if typ == 'fig' else '표 1.') if ko else ('Fig. 1.' if typ == 'fig' else 'Table 1.')}' 형식의 캡션 줄을 찾지 못했습니다. 캡션이 이미지 안에 있거나 형식이 달라 확인을 건너뜁니다.",
                              "캡션을 텍스트로 넣었는지 확인하세요."))
            continue
        for n, cap in sorted(c.items()):
            if n not in mentioned:
                out.append(_issue(doc, "minor", f"{name} 미언급", f"{cap['label']} 이(가) 본문에서 언급되지 않음", cap["pos"], cap["len"],
                                  f"{cap['label']} 캡션은 있지만 본문 어디에서도 '{cap['label']}' 을(를) 가리키지 않습니다. 심사위원은 본문이 설명하지 않는 그림·표를 군더더기나 누락으로 봅니다.",
                                  f"해당 내용을 설명하는 문장에 '({cap['label']})' 를 넣거나, 필요 없으면 빼세요.", quote=f"{cap['label']}: {cap['text'][:160]}"))
        seen_bad = set()
        for _, n, pos, ln in m:
            if c and n not in c and n not in seen_bad:
                seen_bad.add(n)
                out.append(_issue(doc, "minor", f"{name} 번호", f"없는 {name} 번호 언급: {name} {n}", pos, ln,
                                  f"본문이 {name} {n} 을(를) 가리키지만 그 번호의 캡션이 없습니다(있는 번호: {', '.join(map(str, sorted(c)))}).",
                                  "번호를 고치거나 빠진 그림·표를 넣으세요."))
        if c:
            nums = sorted(c)
            gaps = [n for n in range(1, nums[-1] + 1) if n not in c]
            if gaps:
                out.append(_issue(doc, "minor", f"{name} 번호", f"{name} 번호 건너뜀: {', '.join(map(str, gaps))}", c[nums[0]]["pos"], c[nums[0]]["len"],
                                  f"캡션 번호가 {', '.join(map(str, nums))} 로, {', '.join(map(str, gaps))} 이(가) 빠져 있습니다.",
                                  "번호를 연속되게 다시 매기세요.", quote=f"{c[nums[0]]['label']}: {c[nums[0]]['text'][:120]}"))
            first = []
            for _, n, pos, ln in m:
                if n in c and n not in [f[0] for f in first]:
                    first.append((n, pos, ln))
            for k in range(1, len(first)):
                if first[k][0] < max(f[0] for f in first[:k]) and first[k][0] != 1:
                    n, pos, ln = first[k]
                    out.append(_issue(doc, "info", f"{name} 번호", f"{name} {n} 이(가) 더 큰 번호보다 늦게 처음 언급됨", pos, ln,
                                      f"{name} 번호는 본문에서 처음 언급되는 순서대로 매기는 것이 관례입니다.", "번호를 처음 언급 순서로 다시 매기세요."))
                    break
    # 표기 혼용: Fig. 과 Figure
    if not ko:
        forms = {}
        for x in re.finditer(r"(?<![A-Za-z])(Fig\.|Figure)\s*\d", text[lo:hi]):
            forms.setdefault(x.group(1), lo + x.start())
        if len(forms) == 2:
            pos = max(forms.values())
            out.append(_issue(doc, "info", "표기 일관성", "그림 표기 혼용: 'Fig.' 과 'Figure'", pos, 6,
                              "본문에서 'Fig.' 과 'Figure' 를 섞어 씁니다(문장 첫머리의 'Figure' 는 관례상 허용).", "대상 저널 규정에 맞춰 하나로 통일하세요."))
    return out


def check_citations(doc):
    text, refs, out = doc["text"], doc["refs"], []
    lo, hi = body_span(doc)
    if not refs:
        if doc["sections"] and not any(s["kind"] == "references" for s in doc["sections"]):
            out.append(_issue(doc, "info", "참고문헌", "참고문헌 절을 찾지 못함", max(0, hi - 1), 1,
                              "'References'/'참고문헌' 제목을 찾지 못해 인용 점검을 건너뜁니다.", "참고문헌 절 제목을 넣었는지 확인하세요.", quote=""))
        return out
    if doc["ref_style"] == "numeric":
        nums = [r["n"] for r in refs]
        N = max(nums)
        listed = set(nums)
        missing_entries = [n for n in range(1, N + 1) if n not in listed]
        if missing_entries:
            out.append(_issue(doc, "minor", "참고문헌", f"참고문헌 목록 번호 누락: {', '.join(map(str, missing_entries[:10]))}", refs[0]["pos"], 10,
                              f"참고문헌 목록이 1~{N} 인데 {', '.join(map(str, missing_entries[:10]))} 번 항목이 없습니다(추출 오류일 수도 있음).",
                              "목록 번호를 확인하세요.", quote=refs[0]["text"][:120]))
        cited, first_pos = [], {}
        for m in CITE_NUM.finditer(text, lo, hi):
            ns = _cite_nums(m.group(1))
            if 0 in ns:
                continue  # [0, 1] 같은 구간 표기
            for n in ns:
                cited.append(n)
                first_pos.setdefault(n, (m.start(), m.end() - m.start()))
        if not cited:
            out.append(_issue(doc, "info", "인용", "본문에서 [n] 형식 인용을 찾지 못함", lo, 1,
                              "번호식 참고문헌인데 본문에 [1] 같은 인용 표기가 없습니다. 위첨자 인용이면 텍스트 추출로는 확인할 수 없습니다.",
                              "인용 표기를 확인하세요.", quote=""))
            return out
        bad = sorted({n for n in cited if n not in listed})
        for n in bad[:10]:
            pos, ln = first_pos[n]
            out.append(_issue(doc, "minor", "인용", f"목록에 없는 참고문헌 번호 인용: [{n}]", pos, ln,
                              f"본문이 [{n}] 을(를) 인용하지만 참고문헌 목록은 1~{N} 입니다.", "인용 번호를 고치거나 빠진 문헌을 목록에 넣으세요."))
        unused = [r for r in refs if r["n"] not in set(cited)]
        for r in unused[:15]:
            out.append(_issue(doc, "minor", "인용", f"본문에서 인용되지 않은 참고문헌: [{r['n']}]", r["pos"], r["len"],
                              "참고문헌 목록에 있지만 본문에서 한 번도 인용되지 않았습니다. 대부분의 저널은 미인용 문헌을 허용하지 않습니다.",
                              "본문의 알맞은 곳에서 인용하거나 목록에서 빼세요.", quote=r["text"][:200]))
        if len(unused) > 15:
            out.append(_issue(doc, "minor", "인용", f"미인용 참고문헌이 {len(unused) - 15}개 더 있음", unused[15]["pos"], 5,
                              ", ".join(f"[{r['n']}]" for r in unused[15:40]), "목록을 정리하세요.", quote=""))
        order = sorted(first_pos, key=lambda n: first_pos[n][0])
        inorder = sum(1 for k in range(1, len(order)) if order[k] > order[k - 1])
        if len(order) >= 6 and inorder / (len(order) - 1) >= 0.7 and inorder < len(order) - 1:
            # 등장 순서식으로 보이는데 어긋난 첫 번호
            mx = 0
            for n in order:
                if n > mx + 1 and n not in bad:
                    pos, ln = first_pos[n]
                    out.append(_issue(doc, "info", "인용", f"인용 번호 순서: [{n}] 이 [{mx + 1}] 보다 먼저 처음 인용됨", pos, ln,
                                      "등장 순서 번호식(Vancouver·IEEE)이면 처음 인용되는 순서대로 번호를 매겨야 합니다.",
                                      "저널이 등장 순서식이면 번호를 다시 매기세요(알파벳순 번호식이면 무시)."))
                    break
                mx = max(mx, n)
    elif doc["ref_style"] == "author-year":
        body = text[lo:hi]
        cites = {}
        for m in CITE_AY.finditer(body):
            name, year = m.group(1), m.group(2)
            if name.lower() in ("in", "the", "table", "figure", "fig", "since", "until", "from", "by", "of", "and", "january", "february",
                                "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"):
                continue
            if "et al" in m.group(0) or "(" in m.group(0) or re.search(r"\(\s*$", body[max(0, m.start() - 2):m.start()]):
                cites.setdefault((name, year), lo + m.start())
        reftexts = [r["text"] for r in refs]
        for (name, year), pos in list(cites.items())[:200]:
            if not any(name in t and year in t for t in reftexts):
                out.append(_issue(doc, "minor", "인용", f"참고문헌 목록에 없는 인용: {name} ({year})", pos, len(name) + 6,
                                  f"본문의 '{name} … {year}' 인용과 맞는 항목(저자 성 + 연도)을 참고문헌 목록에서 찾지 못했습니다.",
                                  "목록에 추가하거나 저자·연도를 고치세요."))
        for r in refs:
            m = re.match(r"([A-Z][A-Za-z'\-]+|[가-힣]{2,4})", r["text"])
            y = re.search(r"((?:19|20)\d\d)", r["text"])
            if m and y and not re.search(re.escape(m.group(1)) + r"[^.;]{0,60}?" + y.group(1), body):
                out.append(_issue(doc, "minor", "인용", f"본문에서 인용되지 않은 참고문헌: {m.group(1)} ({y.group(1)})", r["pos"], r["len"],
                                  "첫 저자 성과 연도로 본문 인용을 찾지 못했습니다(표기가 다르면 오탐일 수 있음).",
                                  "본문에서 인용하거나 목록에서 빼세요.", quote=r["text"][:200]))
    return out


ABBR = re.compile(r"(?<![A-Za-z0-9\-/])([A-Z][A-Z]{1,9})(s?)(?![A-Za-z0-9]|-[A-Z])")


def check_abbrev(doc):
    text, out = doc["text"], []
    lo, hi = body_span(doc)
    cap_lines = [(c["pos"], c["pos"] + c["len"]) for c in doc["captions"]]
    first, defs, counts = {}, {}, {}
    for m in ABBR.finditer(text, lo, hi):
        a = m.group(1)
        if a in WHITELIST or len(a) < 2 or len(a) > 8:
            continue
        ls = text.rfind("\n", 0, m.start()) + 1
        le = text.find("\n", m.end())
        line = text[ls:le if le >= 0 else len(text)]
        letters = re.findall(r"[A-Za-z]", line)
        if not re.search(r"[가-힣]", line) and (len(letters) >= 8 and sum(ch.isupper() for ch in letters) / len(letters) > 0.6
                                                or len(re.findall(r"[a-z]{3,}", line)) < 4):  # 대문자 제목 줄·표 칸·수식 줄
            continue  # 대문자 제목 줄
        if re.fullmatch(r"[A-Z]", a[:1]) and re.match(r"\s*\.", text[m.end():m.end() + 2]) and len(a) <= 2:
            continue  # 이니셜
        if re.match(r",?\s*(?:USA|U\.S\.A|Korea|Japan|China)", text[m.end():m.end() + 12]):
            continue  # 주(州) 약자: Long Beach, CA, USA
        counts[a] = counts.get(a, 0) + 1
        first.setdefault(a, (m.start(), len(a)))
    for a in list(first):
        pat = re.compile(rf"\(\s*{a}s?\s*[),;]|(?<![A-Za-z]){a}s?\s*\(\s*[A-Za-z가-힣][^()]{{3,80}}\)|(?<![A-Za-z]){a}s?\s*(?:,\s*i\.e\.|stands for|denotes|refers to|is short for|은|는)\s")
        for d in pat.finditer(text, lo, hi):
            inner = d.group(0)
            if "(" in inner and inner.strip().startswith(a) and not re.search(r"[a-z가-힣]{3}", inner.split("(", 1)[1]):
                continue  # LSTM (2) 처럼 숫자·기호만
            if re.match(rf"{a}s?\s*\(", inner):
                exp = inner.split("(", 1)[1]
                if not re.search(r"[가-힣]", exp) and len(re.findall(r"[A-Za-z]{2,}", exp)) < 2:
                    continue
            defs.setdefault(a, d.start())
            break
    undefined = [a for a in first if a not in defs]
    late = [a for a in first if a in defs and defs[a] > first[a][0] + 40]
    for a in sorted(undefined, key=lambda x: first[x][0])[:25]:
        pos, ln = first[a]
        out.append(_issue(doc, "minor", "약어", f"약어 정의 없음: {a}", pos, ln,
                          f"'{a}' 를 {counts[a]}회 쓰지만 풀어 쓴 정의(예: '전체 이름 ({a})')를 찾지 못했습니다.",
                          f"처음 나오는 곳에서 '풀어 쓴 이름 ({a})' 로 정의하세요. 분야에서 아주 흔한 약어라도 저널 규정을 확인하세요."))
    for a in sorted(late, key=lambda x: first[x][0])[:15]:
        pos, ln = first[a]
        out.append(_issue(doc, "minor", "약어", f"약어를 정의 전에 사용: {a}", pos, ln,
                          f"'{a}' 는 {where(doc, defs[a])} 에서 정의되지만 그보다 앞({where(doc, pos)})에서 먼저 쓰였습니다.",
                          "처음 쓰는 곳으로 정의를 옮기세요."))
    if len(undefined) > 25:
        out.append(_issue(doc, "info", "약어", f"정의 없는 약어 {len(undefined) - 25}개 더", lo, 1, ", ".join(sorted(undefined[25:])[:40]), "목록을 확인하세요.", quote=""))
    return out


NUMRX = re.compile(r"(?<![\w.\-/])(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+\.\d+|\d+)(\s?%|\s?퍼센트)?")


def _val(s):
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return None


def _words_before(text, pos, n=3):
    w = re.findall(r"[A-Za-z가-힣]+", text[max(0, pos - 60):pos].lower())
    return tuple(w[-n:])


def _ko_anchor(text, pos):
    """한글 문맥: 공백·줄바꿈을 지운 앞 글자들 (PDF 가 낱말 중간에서 줄을 바꿔도 같게)"""
    return re.sub(r"[^A-Za-z가-힣]", "", text[max(0, pos - 40):pos].lower())


def _same_ko(a, b):
    """'이상탐지정확도' ↔ '이상탐지정확도는' — 끝 조사 1~2자 차이를 허용해 끝 5글자 비교"""
    if not (a and b and re.match(r"[가-힣]", a[-1]) and re.match(r"[가-힣]", b[-1])):
        return False
    return any(len(x) >= 5 and len(y) >= 5 and x[-5:] == y[-5:] for x in (a, a[:-1], a[:-2]) for y in (b, b[:-1], b[:-2]))


def _same_words(a, b):
    """낱말 묶음 비교. 한글은 조사 차이를 허용(정확도 ↔ 정확도는)."""
    if len(a) != len(b):
        return False
    for x, y in zip(a, b):
        if x != y and not (re.match(r"[가-힣]", x) and min(len(x), len(y)) >= 2 and (x.startswith(y) or y.startswith(x)) and abs(len(x) - len(y)) <= 2):
            return False
    return True


def _word_after(text, pos):
    m = re.match(r"\s*%?\s*([A-Za-z가-힣]+)", text[pos:pos + 20])
    return m.group(1).lower() if m else ""


def check_numbers(doc):
    """초록·결론의 수치가 본문(그 밖의 곳)에서 확인되는지, 같은 문맥에서 다른 값이 쓰였는지"""
    text, out = doc["text"], []
    lo, hi = body_span(doc)
    for kind, name in (("abstract", "초록"), ("conclusion", "결론")):
        spans = region(doc, kind)
        if not spans:
            continue
        a0, a1 = spans[0]
        summ = region(doc, "abstract")[:1] + region(doc, "conclusion")[:1]  # 근거는 초록·결론 밖(본문·표)에서 찾는다
        others = [(m, _val(m.group(1))) for m in NUMRX.finditer(text, lo, hi) if not any(x <= m.start() < y for x, y in summ)]
        reported = 0
        own = {_val(m.group(1)) for m in NUMRX.finditer(text, a0, a1)}
        for m in NUMRX.finditer(text, a0, a1):
            s, pct = m.group(1), bool(m.group(2))
            v = _val(s)
            if v is None or (not pct and "." not in s and (v < 10 or 1900 <= v <= 2100)):
                continue
            if re.match(r"\s*\]", text[m.end():m.end() + 2]) or re.search(r"\[\s*$", text[max(0, m.start() - 2):m.start()]):
                continue  # 인용 번호
            same = [o for o, ov in others if ov is not None and (ov == v or (pct and abs(ov * 100 - v) < 1e-6) or
                                                                   ("." in s and abs(ov - v) < 0.5 * 10 ** -len(s.split(".")[1]) and o.group(1) != s))]
            before, after = _words_before(text, m.start()), _word_after(text, m.end())
            kob = _ko_anchor(text, m.start())
            conflict, best = None, 9.0
            for o, ov in others:
                if ov is None or ov == v or ov in own or not (0.5 <= ov / v <= 2 if v else False):
                    continue
                ob, oa = _words_before(text, o.start()), _word_after(text, o.end())
                if (_same_ko(kob, _ko_anchor(text, o.start())) and bool(o.group(2)) == pct) or (len(before) == 3 and _same_words(ob, before)) or (len(before) >= 2 and after and _same_words(ob[-2:], before[-2:]) and oa == after
                                                           and bool(o.group(2)) == pct):
                    if abs(ov - v) / abs(v) < best:  # 같은 문맥 후보가 여럿이면 값이 가장 가까운 것(오기일 가능성)
                        conflict, best = o, abs(ov - v) / abs(v)
            if conflict:
                bq = sentence_at(text, conflict.start(), len(conflict.group(0)))[0]
                out.append(_issue(doc, "minor" if same else "major", "수치 불일치", f"{name} {s}{'%' if pct else ''} ↔ 본문 {conflict.group(1)}{'%' if conflict.group(2) else ''}",
                                  m.start(), len(m.group(0)),
                                  f"{name}의 '{' '.join(before)} {s}' 와 같은 표현인 본문 문장({where(doc, conflict.start())})의 값이 {conflict.group(1)} 로 다릅니다: “{bq[:160]}”",
                                  ("같은 값은 다른 곳(표 등)에도 있으니 본문 문장 쪽 오기일 수 있습니다. " if same else "") + "어느 값이 맞는지 확인해 초록·본문·표를 일치시키세요.") | {"pos2": conflict.start(), "len2": len(conflict.group(0))})
                reported += 1
            elif not same:
                out.append(_issue(doc, "minor", "수치 확인", f"{name}의 수치 {s}{'%' if pct else ''} 를 본문에서 찾지 못함", m.start(), len(m.group(0)),
                                  f"{name}에 쓴 {s}{'%' if pct else ''} 이(가) 본문·표 어디에도 같은 값으로 나오지 않습니다(반올림·단위 변환 표기면 무시).",
                                  f"본문 결과나 표에 같은 값을 제시하거나 {name} 수치를 본문과 맞추세요."))
                reported += 1
            if reported >= 12:
                break
    return out


def check_style(doc):
    text, out = doc["text"], []
    lo, hi = body_span(doc)
    body = text[lo:hi]
    # 반복 단어 (the the)
    n = 0
    for m in re.finditer(r"(?<![A-Za-z])([A-Za-z]{2,})\s+\1(?![A-Za-z])", body, re.I):
        if m.group(1).lower() in ("that", "had", "is", "bora", "can", "do", "no", "very"):
            continue
        out.append(_issue(doc, "minor", "오탈자", f"단어 반복: '{m.group(0)}'", lo + m.start(), len(m.group(0)),
                          "같은 단어가 연달아 나옵니다.", "하나를 지우세요."))
        n += 1
        if n >= 8:
            break
    # 숫자와 단위 붙여쓰기 (SI: 10 mm)
    hits = [m for m in re.finditer(rf"(?<![\w.])\d+(?:\.\d+)?{SI_UNITS}(?![A-Za-z])", body)]
    if hits and doc["lang"] == "en":
        m = hits[0]
        out.append(_issue(doc, "info", "단위", f"숫자와 단위를 붙여 씀 ({len(hits)}곳, 예: '{m.group(0)}')", lo + m.start(), len(m.group(0)),
                          "SI 표기에서는 숫자와 단위 사이에 한 칸 띄웁니다(예: 10 mm). 같은 원고 안에서 혼용되면 더 눈에 띕니다.",
                          "숫자와 단위 사이를 띄우세요(% 와 ° 는 예외인 저널이 많음): " + ", ".join(dict.fromkeys(h.group(0) for h in hits[:8]))))
    # 용어 표기 혼용: data set / dataset / data-set
    if doc["lang"] == "en":
        seen = 0
        for m in re.finditer(r"(?<![\w-])([a-z]{3,})-([a-z]{3,})(?![\w-])", body):
            w1, w2 = m.group(1), m.group(2)
            joined, spaced = w1 + w2, w1 + " " + w2
            forms = {f: len(re.findall(rf"\b{f}\b", body, re.I)) for f in (f"{w1}-{w2}", joined, spaced)}
            if forms[joined] + forms[spaced] and forms[f"{w1}-{w2}"]:
                key = f"{w1}-{w2}"
                if key in [i.get("_k") for i in out]:
                    continue
                alt = joined if forms[joined] else spaced
                if w1 in ("non", "self", "pre", "re", "co", "multi", "sub") and alt == spaced:
                    continue
                mpos = re.search(rf"\b{alt}\b", body, re.I)
                out.append(_issue(doc, "info", "용어 일관성", f"용어 표기 혼용: '{key}' ({forms[key]}회) / '{alt}' ({forms[alt]}회)", lo + mpos.start(), len(alt),
                                  "같은 용어를 서로 다른 형태로 씁니다.", "한 가지 표기로 통일하세요.") | {"_k": key})
                seen += 1
                if seen >= 6:
                    break
    for i in out:
        i.pop("_k", None)
    return out


EQ_REF = re.compile(r"(?:Eqs?\.|Equations?|equations?|식|수식)\s*\(?\s*(\d{1,3})\s*\)?(?:\s*(?:[-–—~]|and|to|,|및)\s*\(?(\d{1,3})\)?)?")
EQ_LABEL = re.compile(r"(?m)^(?=.*[=+\-−×·∑∫√≤≥<>^_]|\s*\(\d{1,3}\)\s*$).{0,200}?\((\d{1,3})\)\s*$")


def check_equations(doc):
    text, out = doc["text"], []
    lo, hi = body_span(doc)
    labels = {}
    for m in EQ_LABEL.finditer(text, lo, hi):
        labels.setdefault(int(m.group(1)), m.start(1))
    if not labels:
        return out
    mx = max(labels)
    plausible = {n for n in labels if n <= mx and n <= len(labels) + 3}  # 표 안 숫자 오인을 줄이기
    if not plausible:
        return out
    seen = set()
    for m in EQ_REF.finditer(text, lo, hi):
        ns = [int(m.group(1))] + ([int(m.group(2))] if m.group(2) else [])
        for n in ns:
            if n not in labels and n > max(plausible) and n not in seen:
                seen.add(n)
                out.append(_issue(doc, "minor", "수식 번호", f"없는 수식 번호 언급: ({n})", m.start(), m.end() - m.start(),
                                  f"본문이 수식 ({n}) 을 가리키지만 번호가 붙은 수식은 ({max(plausible)}) 까지입니다(추출 한계로 오탐일 수 있음).",
                                  "수식 번호를 확인하세요."))
    return out


def check_structure(doc, mtype=""):
    out, kinds = [], {s["kind"] for s in doc["sections"]}
    need = [("abstract", "초록")]
    if not re.search(r"리뷰|review|survey|letter|레터|짧은|short", mtype, re.I):
        need += [("intro", "서론"), ("method", "방법"), ("results", "결과"), ("conclusion", "결론")]
    miss = [n for k, n in need if k not in kinds and not (k == "conclusion" and "discussion" in kinds) and not (k == "results" and "discussion" in kinds and "method" in kinds)]
    if miss:
        out.append({"sev": "info", "cat": "구조", "title": "찾지 못한 절: " + ", ".join(miss), "quote": "", "pos": 0, "len": 0,
                    "why": "절 제목으로 " + ", ".join(miss) + " 에 해당하는 절을 찾지 못했습니다(제목이 다르거나 추출 오류일 수 있음). 인식한 절: "
                           + " / ".join(s["title"] or KIND_LABEL[s["kind"]] for s in doc["sections"][:14]),
                    "fix": "IMRaD 구성(서론·방법·결과·논의·결론)이 저널 요구와 맞는지 확인하세요.", "src": "rule", "where": "", "page": None})
    ab = region(doc, "abstract")
    if ab and doc["lang"] == "en":
        a, b = ab[0]
        words = len(re.findall(r"[A-Za-z]+", doc["text"][a:b]))
        if words > 320:
            out.append(_issue(doc, "info", "구조", f"초록이 깁니다 ({words}단어)", a, 10,
                              "대부분의 저널은 초록을 150~300단어로 제한합니다.", "대상 저널의 초록 분량 규정을 확인하세요.", quote=""))
    return out


def mechanical(doc, mtype=""):
    out = []
    for f in (check_structure, check_figures, check_citations, check_abbrev, check_numbers, check_equations, check_style):
        try:
            out += f(doc, mtype) if f is check_structure else f(doc)
        except Exception as e:  # 한 점검이 깨져도 나머지는 낸다
            out.append({"sev": "info", "cat": "점검 오류", "title": f"{f.__name__} 실패: {type(e).__name__}: {e}", "quote": "", "pos": 0,
                        "len": 0, "why": "", "fix": "", "src": "rule", "where": "", "page": None})
    return out


def outline(doc):
    return [{"title": s["title"] or KIND_LABEL[s["kind"]], "kind": s["kind"], "label": KIND_LABEL.get(s["kind"], ""),
             "start": s["start"], "end": s["end"], "page": page_of(doc, s["start"]), "level": s["level"], "chars": s["end"] - s["start"]}
            for s in doc["sections"]]
