#!/usr/bin/env python3
"""review local — 투고 전 논문 사전 리뷰어. 로컬 LLM, 외부 의존성 없음(stdlib), 원고는 외부로 나가지 않음.

  python3 app.py                                  # http://localhost:8782
  LLM_API=openai LLM_BASE_URL=http://gpu:8000/v1 LLM_MODEL=Qwen3-32B python3 app.py
  python3 app.py --cli paper.pdf [quick|deep] [ko|en|auto] > review.md     # 터미널에서 바로

흐름: 업로드 → 텍스트 추출(pdftotext / DOCX stdlib / HWP kordoc) → 절·캡션·참고문헌 인식 → 기계적 점검(규칙, LLM 없음)
     → [LLM] 개요(기여 3줄) → 절 단위 발췌마다 심사 지적(원문 인용 필수) → 인용을 원고에서 찾아 확인(못 찾으면 버림)
     → [LLM] 종합(판정·이유·예상 질문 5개) → 화면·MD·DOCX·TXT
"""
import base64
import datetime
import json
import os
import re
import secrets
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import export
import manuscript as M

ROOT = os.path.dirname(os.path.abspath(__file__))
WS = os.environ.get("WORKSPACE") or os.path.join(ROOT, "_workspace")  # 포털이 AGENT_DATA/<도구> 로 모아 줌
LLM_API = os.environ.get("LLM_API", "ollama")            # ollama | openai (vLLM·LM Studio·llama.cpp 등)
LLM_BASE = os.environ.get("LLM_BASE_URL", "http://localhost:8000/v1" if LLM_API == "openai" else "http://localhost:11434").rstrip("/")
MODEL = os.environ.get("LLM_MODEL", "qwen3:8b")
LLM_KEY = os.environ.get("LLM_API_KEY", "")
NUM_CTX = int(os.environ.get("NUM_CTX", "32768"))
PORT = int(os.environ.get("PORT", "8782"))
MAX_UPLOAD = 60 * 1024 * 1024
LOCK = threading.Lock()
GPU = threading.Semaphore(1)  # 리뷰는 한 번에 하나 (LLM 서버 하나를 나눠 쓴다)

MTYPES = ["연구논문", "리뷰(총설)", "학회 논문", "단신·레터", "학위논문(장)", "기술보고서"]
DEPTHS = {"quick": ("빠름", 20000, 5, 4), "deep": ("정밀", 8000, 24, 6)}  # (이름, 발췌 글자 수, 최대 발췌 수, 발췌당 최대 지적)
LANGS = {"auto": "원고 언어", "ko": "국문", "en": "영문"}


def read(p):
    with open(p, encoding="utf-8") as f:
        return f.read()


# ── LLM (mail-local 과 동일) ─────────────────────────────────────────────
def _clean(out):
    out = re.sub(r"<think>.*?</think>", "", out or "", flags=re.S).strip()
    out = re.sub(r"^```\w*\s*\n", "", out)
    return re.sub(r"\n?```\s*$", "", out).strip()


def _openai(msgs, model, temperature, on_token):
    body = {"model": model, "stream": True, "temperature": temperature, "messages": msgs}
    hdr = {"Content-Type": "application/json", **({"Authorization": f"Bearer {LLM_KEY}"} if LLM_KEY else {})}
    req = urllib.request.Request(LLM_BASE + "/chat/completions", json.dumps(body).encode(), hdr)
    buf = []
    try:
        with urllib.request.urlopen(req, timeout=1800) as r:
            for line in r:
                line = line.decode().strip()
                if not line.startswith("data:") or line == "data: [DONE]":
                    continue
                tok = (json.loads(line[5:])["choices"][0].get("delta") or {}).get("content") or ""
                if tok:
                    buf.append(tok)
                    on_token(tok)
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"LLM HTTP {e.code}: {e.read().decode(errors='replace')[:300]}")
    return "".join(buf)


def llm(system, user, model=None, temperature=0.3, on_token=lambda t: None):
    """스트리밍 채팅. Ollama /api/chat (think:false 미지원 모델이면 자동 재시도) 또는 OpenAI 호환."""
    model = model or MODEL
    msgs = [{"role": "system", "content": system}, {"role": "user", "content": user}]
    try:
        if LLM_API == "openai":
            return _openai(msgs, model, temperature, on_token)
        body = {"model": model, "stream": True, "think": False, "messages": msgs,
                "options": {"temperature": temperature, "num_ctx": NUM_CTX}}
        for attempt in (0, 1):
            try:
                req = urllib.request.Request(LLM_BASE + "/api/chat", json.dumps(body).encode(), {"Content-Type": "application/json"})
                buf = []
                with urllib.request.urlopen(req, timeout=1800) as r:
                    for line in r:
                        if not line.strip():
                            continue
                        j = json.loads(line)
                        if "error" in j:
                            raise RuntimeError(j["error"])
                        tok = j.get("message", {}).get("content", "")
                        if tok:
                            buf.append(tok)
                            on_token(tok)
                        if j.get("done"):
                            break
                return "".join(buf)
            except urllib.error.HTTPError as e:
                msg = e.read().decode(errors="replace")
                if attempt == 0 and "think" in msg:
                    body.pop("think")
                    continue
                raise RuntimeError(f"LLM HTTP {e.code}: {msg[:300]}")
    except urllib.error.URLError as e:
        raise RuntimeError(f"LLM 서버 연결 실패 ({LLM_BASE}): {e.reason}")


def models():
    try:
        if LLM_API == "openai":
            req = urllib.request.Request(LLM_BASE + "/models", headers={"Authorization": f"Bearer {LLM_KEY}"} if LLM_KEY else {})
            with urllib.request.urlopen(req, timeout=3) as r:
                return [m["id"] for m in json.load(r)["data"]]
        with urllib.request.urlopen(LLM_BASE + "/api/tags", timeout=3) as r:
            return [m["name"] for m in json.load(r)["models"]]
    except Exception:
        return []


# ── 발췌 나누기 ─────────────────────────────────────────────────────────
SKIP_KINDS = ("front", "references", "ack", "appendix")


def flat(s):
    """LLM 에 줄 글: 줄바꿈 하나는 공백으로(빈 줄은 문단 경계로 유지), 하이픈 줄바꿈 이음"""
    s = re.sub(r"(?<=[a-z])-\n(?=[a-z])", "", s)
    s = re.sub(r"[ \t]*\n(?!\n)[ \t]*", " ", s)
    return re.sub(r"\n{2,}", "\n\n", s).strip()


def make_chunks(doc, size, max_n):
    """절 경계를 따라 ~size 글자씩. 너무 많으면 size 를 키워 max_n 개 안에 넣는다."""
    secs = [s for s in doc["sections"] if s["kind"] not in SKIP_KINDS]
    if not secs:  # 절을 못 찾았으면 본문 전체
        lo, hi = M.body_span(doc)
        secs = [{"title": "본문", "kind": "body", "start": lo, "end": hi}]
    total = sum(s["end"] - s["start"] for s in secs)
    size = max(size, min(24000, total // max_n + 1))
    chunks, cur = [], None

    def add_title(s):
        t = s["title"] or M.KIND_LABEL[s["kind"]]
        if t not in cur["titles"]:
            cur["titles"].append(t)
    for s in secs:
        a, b = s["start"], s["end"]
        if cur and cur["end"] != a:  # 사이에 건너뛴 절(감사의 글 등)이 있으면 끊는다
            chunks.append(cur)
            cur = None
        while a < b:
            if cur is None:
                cur = {"start": a, "end": a, "titles": []}
            room = size - (cur["end"] - cur["start"])
            if b - a <= room:
                cur["end"] = b
                add_title(s)
                a = b
            elif cur["end"] - cur["start"] > size * 0.35:
                chunks.append(cur)
                cur = None
            else:
                cut = _cut(doc["text"], a, a + room)
                cur["end"] = cut
                add_title(s)
                a = cut
                chunks.append(cur)
                cur = None
    if cur:
        chunks.append(cur)
    if len(chunks) > max_n:  # 너무 길면 앞부분만이 아니라 처음·끝을 포함해 고르게 고른다
        pick = sorted({round(i * (len(chunks) - 1) / (max_n - 1)) for i in range(max_n)}) if max_n > 1 else [0]
        chunks = [chunks[i] for i in pick]
    return chunks


def _cut(text, a, b):
    """b 근처의 문단·문장 경계"""
    for pat in ("\n\n", ".\n", ". ", "다. ", "\n"):
        k = text.rfind(pat, a + (b - a) // 2, b)
        if k > 0:
            return k + len(pat)
    return b


# ── 프롬프트 ────────────────────────────────────────────────────────────
def out_lang(opt, doc):
    lang = opt.get("lang") or "auto"
    return doc["lang"] if lang == "auto" else lang


def lang_rule(lang):
    return ("모든 설명(요지·문제·수정·판정 이유·질문·답변 팁)은 자연스러운 학술 영어로 쓴다. 인용만 원고 원문 그대로."
            if lang == "en" else "모든 설명(요지·문제·수정·판정 이유·질문·답변 팁)은 한국어로 쓴다. 인용만 원고 원문 그대로(영어 원고면 영어 그대로).")


def context_block(doc, opt):
    lines = [f"[원고 제목] {doc.get('title') or doc['name']}",
             f"[원고 유형] {opt.get('mtype') or '연구논문'}"]
    if (opt.get("journal") or "").strip():
        lines.append(f"[대상 저널·분야] {opt['journal'].strip()} — 이 저널·분야 심사위원의 눈높이로 본다. 저널 규정을 모르면 지어내지 않는다.")
    lines.append("[절 목록] " + " / ".join(f"{s['title'] or M.KIND_LABEL[s['kind']]}" for s in doc["sections"] if s["kind"] != "front")[:1500])
    caps = [f"{c['label']}: {c['text'][:90]}" for c in doc["captions"]][:30]
    if caps:
        lines.append("[그림·표 캡션]\n" + "\n".join(caps))
    return "\n".join(lines)


OVERVIEW_SYS = """너는 학술지 심사위원이다. 원고의 핵심부(초록·서론·결론)를 읽고, 원고가 스스로 주장하는 내용을 정리한다. 평가하지 말고 원고가 말하는 것만.
출력 형식 (정확히 이 표시만, 앞뒤 설명 없이):
[기여]
1. (원고가 주장하는 기여 1 — 한 문장, 원고에 있는 수치·데이터·방법 이름을 그대로)
2. (기여 2)
3. (기여 3)
[핵심 주장]
- (원고가 내세우는 검증 대상 주장: 성능 수치·비교 대상·적용 범위 등, 3~5개)
[연구 설계]
- (데이터·실험·비교군·평가 지표를 원고에서 읽은 대로 한두 줄. 원고에 없는 것은 쓰지 않는다)"""

CHUNK_SYS = """너는 {field}분야 국제 학술지의 엄격하고 공정한 심사위원이다. 투고 전 원고의 [이번 발췌]를 읽고, 게재 거절·수정 요구로 이어질 구체적 문제를 찾는다.
반드시 지킬 것:
1. 지적마다 [이번 발췌]의 문장을 '인용:' 에 그대로 복사한다(20~250자, 번역·요약·의역 금지, 중간 생략 금지). 인용할 문장이 없는 지적은 쓰지 않는다.
   빠진 것(대조군·통계 검정·데이터 정보 등)을 지적할 때도, 그 부재 때문에 근거가 약해지는 '주장 문장'을 인용한다.
2. [이번 발췌]에 실제로 있는 내용만 다룬다. 다른 절(절 목록 참고)에 있을 법한 정보가 이 발췌에 없다는 이유만으로 '누락'이라 단정하지 않는다 — 그런 지적은 쓰지 않는다.
3. 일반론 금지: '더 자세히 설명하라', '영어 교정 필요', '한계를 논의하라', '최신 문헌을 더 인용하라', '그림 품질 개선'처럼 어느 논문에나 할 수 있는 말은 쓰지 않는다.
   '문제' 에는 이 원고의 무엇이 왜 심사위원을 설득하지 못하는지를, '수정' 에는 무엇을 추가·변경하면 되는지(어떤 실험·지표·비교·문장)를 이 원고에 맞게 구체적으로 쓴다.
4. 칭찬·요약·총평·서론 문장은 쓰지 않는다. 문제가 없으면 '없음' 한 단어만 출력한다.
5. 맞춤법, 약어 정의 여부, 그림·표·참고문헌 번호, 숫자-단위 띄어쓰기는 프로그램이 따로 점검하므로 다루지 않는다.
6. 심각도: major = 결론의 타당성을 흔드는 문제(새로움 근거 부족, 기준선·대조군 부재, 통계 검정·오차·반복 수 미제시, 데이터 누수·평가 편향 가능성, 재현에 필요한 정보 누락, 결과보다 강한 결론·일반화). minor = 쉽게 고칠 수 있는 문제(용어 혼용, 모호한 정의, 기호·변수 설명 누락, 그림·표 해석 불충분, 논리 연결 부족).
7. 최대 {n}개, 중요한 것부터. 확신이 없는 지적은 빼라 — 적게 쓰더라도 틀린 지적이 없어야 한다. 같은 문제를 다른 문장으로 되풀이하지 않는다.
8. 수식·기호는 LaTeX($...$, \\text) 로 쓰지 말고 원고 표기대로 일반 글자로 쓴다(예: d_k, PE_pos).
9. 이 글은 PDF 등에서 자동 추출한 텍스트라 수식·첨자·기호(√, Σ, 분수, 위첨자)가 빠지거나 순서가 뒤섞이고, 표가 낱말 나열로 깨져 있을 수 있다.
   깨진 수식 모양, 빠진 기호, 이상한 줄바꿈·띄어쓰기는 원고의 잘못이 아니라 추출 문제이므로 절대 지적하지 않는다.
{lang}
출력 형식 (지적마다 아래 블록을 반복, 다른 글 없이):
### 지적
심각도: major 또는 minor
분류: 새로움 | 방법론 | 통계 | 데이터 | 재현성 | 결론 과대해석 | 논리 | 명확성 | 용어 | 그림·표 | 기타 중 하나
요지: (한 줄 제목)
인용: (발췌 원문 그대로)
문제: (1~3문장)
수정: (1~2문장)
검색어: (이 문제의 답이 원고 다른 절에 있다면 거기 나올 낱말 2~4개, 원고 언어로, 쉼표로 구분)"""

SYNTH_SYS = """너는 학술지 심사위원장이다. 아래 [원고 개요]와 심사위원들이 원고에서 찾은 [지적 목록](원문 인용으로 확인된 것만)을 보고 종합 판정을 내린다.
규칙:
- 판정은 Accept, Minor Revision, Major Revision, Reject 중 하나. 지적 목록과 원고 개요에 근거해서만 판단한다. 목록에 없는 새 문제를 지어내지 않는다.
  기준: 핵심 주장이 실험으로 뒷받침되지 않거나 새 실험·재분석이 필요하면 Major, 서술 보강·표현 수정·추가 설명으로 해결되면 Minor,
  연구 설계 자체가 결론을 낼 수 없으면 Reject. 지적 개수가 아니라 고치는 데 드는 일의 크기로 판단한다.
- 수식·기호는 LaTeX 로 쓰지 말고 일반 글자로(예: d_k).
- 이유는 2~4개, 각 이유 끝에 근거 지적 번호를 (#M1, #m3 처럼) 단다.
- 예상 질문은 정확히 5개. 먼저 [지적 목록]의 중요한 지적에서 만들고 그 번호를 반드시 단다. 지적이 5개보다 적으면 나머지는 [원고 개요]의 핵심 주장(수치·비교 대상)을
  겨냥한다. 이 원고의 특정 주장·수치·방법을 겨냥해야 하며, 어느 논문에나 할 질문이나 원고가 표·절(캡션 목록 참고)에서 이미 다뤘을 법한 질문(하이퍼파라미터 영향, 과적합 여부 같은)은 피한다.
  '준비:' 에는 저자가 답변서(rebuttal) 또는 원고 수정으로 무엇을 준비하면 되는지(추가 실험·분석·문장)를 구체적으로.
{lang}
출력 형식 (정확히 이 표시만):
[판정]
(Accept | Minor Revision | Major Revision | Reject)
[이유]
- (이유) (#M1, #M2)
[질문]
1. 질문: (질문)
   준비: (답변 준비 팁)
2. 질문: ...
   준비: ...
(5번까지)"""


# ── 결과 파싱 ───────────────────────────────────────────────────────────
KEYS = {"검색어": "kw", "keywords": "kw", "심각도": "sev", "severity": "sev", "분류": "cat", "category": "cat", "요지": "title", "summary": "title", "title": "title",
        "인용": "quote", "quote": "quote", "문제": "why", "problem": "why", "issue": "why", "수정": "fix", "fix": "fix", "suggestion": "fix"}


def parse_issues(raw):
    raw = _clean(raw)
    out = []
    for block in re.split(r"(?m)^\s*#{2,4}\s*(?:지적|issue|comment)\b.*$", raw, flags=re.I)[1:]:
        it, cur = {}, None
        for line in block.split("\n"):
            m = re.match(r"^\s*[-*]?\s*\**\s*(검색어|keywords|심각도|severity|분류|category|요지|summary|title|인용|quote|문제|problem|issue|수정|fix|suggestion)\s*\**\s*[:：]\s*(.*)$", line, re.I)
            if m:
                cur = KEYS[m.group(1).lower()]
                it[cur] = m.group(2).strip()
            elif cur and line.strip():
                it[cur] = (it[cur] + " " + line.strip()).strip()
        if it.get("why") or it.get("quote"):
            q = it.get("quote", "").strip().strip("\"“”'‘’「」>").strip()
            it["quote"] = q
            sev = it.get("sev", "").lower()
            it["sev"] = "major" if "major" in sev or "주요" in sev else "minor"
            it["cat"] = re.sub(r"[*|]", "", it.get("cat") or "기타").strip()[:20] or "기타"
            out.append(it)
    return out


def parse_overview(raw):
    raw = _clean(raw)
    sec, buf = None, {"기여": [], "핵심 주장": [], "연구 설계": []}
    for line in raw.split("\n"):
        m = re.match(r"^\s*\**\s*\[(기여|핵심 주장|연구 설계)\]", line)
        if m:
            sec = m.group(1)
            continue
        t = re.sub(r"^\s*(?:\d+[.)]|[-*•])\s*", "", line).strip()
        if sec and t:
            buf[sec].append(t)
    return {"contrib": buf["기여"][:3], "claims": buf["핵심 주장"][:6], "design": buf["연구 설계"][:4]}


VERDICTS = [("Major Revision", r"major\s*revision|대폭\s*수정|major"), ("Minor Revision", r"minor\s*revision|소폭\s*수정|minor"),
            ("Reject", r"reject|게재\s*불가|거절"), ("Accept", r"accept|게재\s*가")]


def parse_synth(raw):
    raw = _clean(raw)
    sec, buf = None, {"판정": [], "이유": [], "질문": []}
    for line in raw.split("\n"):
        m = re.match(r"^\s*\**\s*\[(판정|이유|질문)\]\s*\**\s*(.*)$", line)
        if m:
            sec = m.group(1)
            if m.group(2).strip():
                buf[sec].append(m.group(2))
            continue
        if sec and line.strip():
            buf[sec].append(line)
    vt = " ".join(buf["판정"]).strip()
    label = next((name for name, rx in VERDICTS if re.search(rx, vt, re.I)), "")
    reasons = []
    for line in buf["이유"]:
        t = re.sub(r"^\s*(?:\d+[.)]|[-*•])\s*", "", line).strip()
        if t:
            reasons.append({"text": t, "refs": re.findall(r"#([Mm]\d+)", t)})
    questions, cur = [], None
    for line in buf["질문"]:
        m = re.match(r"^\s*(?:\d+[.)]|[-*•])?\s*\**\s*(질문|Q|Question)\s*\**\s*[:：]\s*(.*)$", line, re.I)
        t = re.match(r"^\s*[-*•]?\s*\**\s*(준비|답변\s*준비|Tip|Prep(?:aration)?|A)\s*\**\s*[:：]\s*(.*)$", line, re.I)
        if m:
            cur = {"q": m.group(2).strip(), "tip": ""}
            questions.append(cur)
        elif t and cur:
            cur["tip"] = t.group(2).strip()
        elif re.match(r"^\s*\d+[.)]\s+\S", line):
            cur = {"q": re.sub(r"^\s*\d+[.)]\s+", "", line).strip(), "tip": ""}
            questions.append(cur)
        elif cur and line.strip():
            if cur["tip"]:
                cur["tip"] += " " + line.strip()
            else:
                cur["q"] += " " + line.strip()
    for q in questions:
        q["refs"] = re.findall(r"#([Mm]\d+)", q["q"] + " " + q["tip"])
    return {"label": label, "raw": vt, "reasons": reasons}, questions[:5]


GREEK = {"alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε", "eta": "η", "theta": "θ", "lambda": "λ", "mu": "μ", "nu": "ν",
         "pi": "π", "rho": "ρ", "sigma": "σ", "tau": "τ", "phi": "φ", "omega": "ω", "Sigma": "Σ", "Delta": "Δ", "times": "×", "cdot": "·",
         "in": "∈", "le": "≤", "leq": "≤", "ge": "≥", "geq": "≥", "approx": "≈", "neq": "≠", "pm": "±", "sum": "Σ", "infty": "∞", "rightarrow": "→", "to": "→"}


def delatex(s):
    """모델이 습관처럼 쓰는 LaTeX($\\hat{x}_t$, \\text{PE}) 를 화면용 일반 글자로"""
    if not s or "\\" not in s and "$" not in s:
        return s
    s = re.sub(r"\\(?:text|mathrm|mathbf|mathit|operatorname|textbf|textit)\{([^{}]*)\}", r"\1", s)
    s = re.sub(r"\\hat\{(\w)\}", "\\1\u0302", s)
    s = re.sub(r"\\bar\{(\w)\}", "\\1\u0304", s)
    s = re.sub(r"\\(?:frac)\{([^{}]*)\}\{([^{}]*)\}", r"(\1)/(\2)", s)
    s = re.sub(r"\\sqrt\{([^{}]*)\}", r"√(\1)", s)
    s = re.sub(r"\\([A-Za-z]+)", lambda m: GREEK.get(m.group(1), m.group(1)), s)
    s = re.sub(r"_\{([^{}]*)\}", r"_\1", s)
    s = re.sub(r"\^\{([^{}]*)\}", r"^\1", s)
    return re.sub(r"\\([{}])", r"\1", s).replace("$", "").replace("\\,", " ")


GENERIC = re.compile(r"proofread|native\s+(?:english\s+)?speaker|language\s+editing|영어\s*(?:교정|감수)|원어민|"
                     r"more\s+recent\s+(?:references|literature)|최신\s*문헌을\s*(?:더\s*)?(?:추가|인용)|figure\s+quality|그림\s*(?:의\s*)?(?:해상도|품질)", re.I)


def verify(doc, items, chunk, cache, start_id):
    """LLM 지적의 인용을 원고에서 찾는다. 못 찾거나 일반론이면 버린다. → (확인된 것, 버린 것)"""
    ok, dropped = [], []
    for it in items:
        q = it.get("quote") or ""
        why = it.get("why") or ""
        if GENERIC.search(why + " " + (it.get("fix") or "")) and len(why) < 140:
            dropped.append(dict(it, reason="일반론"))
            continue
        if len(re.sub(r"\s", "", q)) < 8:
            dropped.append(dict(it, reason="인용 없음"))
            continue
        hit = M.locate(doc, q, chunk["start"], chunk["end"], cache)
        if not hit:
            dropped.append(dict(it, reason="원고에서 인용을 찾지 못함"))
            continue
        a, ln, ratio = hit
        it.update(pos=a, len=ln, match=round(ratio, 2), quote=re.sub(r"\s+", " ", doc["text"][a:a + ln]).strip(),
                  where=M.where(doc, a), page=M.page_of(doc, a), src="llm", fix=delatex(it.get("fix") or ""),
                  title=delatex(it.get("title") or ""), why=delatex(why))
        ok.append(it)
    return ok, dropped


def _sim(a, b):
    import difflib
    return difflib.SequenceMatcher(None, a, b, autojunk=False).ratio()


def dedupe(items, doc=None):
    """같은 곳을 가리키거나(인용 겹침) 같은 문제를 되풀이한(같은 분류 + 요지·문제 문장이 비슷) 지적은 하나로. 뒤엣것의 위치는 also 로 남긴다."""
    out = []
    for it in items:
        dup = next((o for o in out if o["pos"] < it["pos"] + it["len"] and it["pos"] < o["pos"] + o["len"]), None)
        if not dup:
            dup = next((o for o in out if o["cat"] == it["cat"] and (_sim(o.get("title", ""), it.get("title", "")) > 0.72
                                                                    or _sim(o["why"][:200], it["why"][:200]) > 0.6)), None)
        if dup:
            if not (dup["pos"] < it["pos"] + it["len"] and it["pos"] < dup["pos"] + dup["len"]):
                dup.setdefault("also", []).append({"pos": it["pos"], "len": it["len"], "where": it["where"], "quote": it["quote"]})
            if it["sev"] == "major":
                dup["sev"] = "major"
            continue
        out.append(it)
    return out


CROSS_SYS = """너는 심사 의견을 검수하는 편집위원이다. 각 [지적]은 원고의 한 발췌만 보고 쓴 것이라, 원고의 다른 곳에 이미 답이 있을 수 있다.
각 지적 아래의 [다른 곳 문단]은 원고 전체에서 그 지적과 관련 있어 보이는 문단을 프로그램이 찾아 붙인 것이다.
지적마다 판단한다:
- 해소: 지적이 '없다·부족하다·밝히지 않았다'고 한 정보·근거·실험이 [다른 곳 문단]에 실제로 있다. 또는 지적이 원고를 잘못 읽었다.
  또는 지적이 텍스트 추출 문제(깨진 수식·빠진 기호·뒤섞인 표)를 원고의 잘못으로 착각했다.
- 유지: 그 밖의 경우. 문단이 관련은 있어도 지적을 완전히 해소하지 못하면 유지. 애매하면 유지.
출력 형식 (지적마다 한 줄, 다른 글 없이):
#번호: 유지
#번호: 해소 — (근거가 된 문단의 핵심 구절을 원문 그대로 짧게)"""

STOP = set("""this that with from have been were which their there these those into also such than then them they when where while about
after before other more most some only over under very will would could should being does more less each both what using used based
paper authors author manuscript section table figure results result method methods model models data study proposed provide provided
explicitly specific clearly clarify unclear lack lacks missing without whether however although therefore described describe""".split())


def windows(doc):
    """교차 확인용 문단 창: 본문(참고문헌 전)을 ~700자 창(겹침 250)으로"""
    lo, hi = M.body_span(doc)
    text, out, a = doc["text"], [], lo
    while a < hi:
        b = min(hi, a + 700)
        out.append((a, b, text[a:b].lower()))
        a = b - 250 if b < hi else hi
    return out


def keywords(it, lang):
    src = (it.get("kw") or "") + " " + (it.get("title") or "") + " " + (it.get("why") or "")
    ws = [w.lower() for w in re.findall(r"[A-Za-z][A-Za-z\-]{3,}", src)]
    if lang == "ko":
        ws += [w[:3] if len(w) > 3 else w for w in re.findall(r"[가-힣]{2,}", it.get("kw") or "")]
    kws = [w for w in dict.fromkeys(ws) if w not in STOP]
    return kws[:12]


def related(doc, wins, it, k=2):
    kws = keywords(it, doc["lang"])
    if not kws:
        return []
    df = {w: sum(1 for _, _, t in wins if w in t) for w in kws}
    scored = []
    for a, b, t in wins:
        if a <= it["pos"] < b or a < it["pos"] + it["len"] <= b:
            continue  # 지적이 인용한 바로 그 자리는 빼고
        sc = sum(1.0 / (1 + df[w]) for w in kws if w in t)
        if sc > 0:
            scored.append((sc, a, b))
    scored.sort(reverse=True)
    out = []
    for sc, a, b in scored:
        if len(out) >= k:
            break
        if all(abs(a - x) > 500 for x, _ in out):
            out.append((a, b))
    return out


def parse_cross(raw):
    res = {}
    for m in re.finditer(r"#\s*([Mm]\d+|[A-Za-z]?\d+)\s*[:：]\s*(유지|해소|keep|resolved?)\s*(?:[—–-]+\s*(.*))?", _clean(raw)):
        res[m.group(1)] = ("drop" if m.group(2).startswith(("해소", "resolv")) else "keep", (m.group(3) or "").strip())
    return res


def cross_check(doc, found, model, emit, lang):
    """각 지적과 관련된 다른 곳 문단을 찾아 붙여 '이미 답이 있는가'를 한 번 더 묻는다 → (남길 것, 버린 것)"""
    wins = windows(doc)
    cand = [it for it in found if related(doc, wins, it)][:16]
    if not cand:
        return found, []
    blocks = []
    for n, it in enumerate(cand, 1):
        it["_cx"] = f"X{n}"
        paras = "\n".join(f"  ({M.where(doc, a)}) {flat(doc['text'][a:b])}" for a, b in related(doc, wins, it))
        blocks.append(f"#X{n} [{it['sev']}/{it['cat']}] {it.get('title') or ''}\n  인용: {it['quote'][:200]}\n  문제: {it['why'][:300]}\n[다른 곳 문단]\n{paras}")
    emit({"stage": "cross", "msg": f"교차 확인: 지적 {len(cand)}개가 원고 다른 곳에서 이미 해소되는지", "reset": True})
    raw = llm(CROSS_SYS + "\n" + lang_rule(lang), "\n\n".join(blocks), model, 0.1, on_token=lambda t: emit({"token": t}))
    verdict = parse_cross(raw)
    keep, drop = [], []
    for it in found:
        v = verdict.get(it.pop("_cx", ""), ("keep", ""))
        if v[0] == "drop":
            drop.append(dict(it, reason="교차 확인: 원고 다른 곳에서 해소" + (f" — {v[1][:160]}" if v[1] else "")))
        else:
            keep.append(it)
    return keep, drop


# ── 리뷰 실행 ───────────────────────────────────────────────────────────
def review(doc, opt, emit=lambda ev: None, model=None):
    model = model or opt.get("model") or MODEL
    depth = opt.get("depth") if opt.get("depth") in DEPTHS else "quick"
    dname, size, max_n, per = DEPTHS[depth]
    lang = out_lang(opt, doc)
    field = (opt.get("journal") or "").strip()
    field = f"'{field}' " if field else ""
    ctx = context_block(doc, opt)
    t0 = time.time()

    # 1) 개요
    text = doc["text"]

    def part(kind, n):
        sp = M.region(doc, kind)
        return flat(text[sp[0][0]:sp[0][1]])[:n] if sp else ""
    core = "\n\n".join(x for x in (
        "[초록]\n" + part("abstract", 3500),
        "[서론 앞부분]\n" + (part("intro", 4000) or flat(text[slice(*M.body_span(doc))])[:4000]),
        "[결론]\n" + (part("conclusion", 3000) or part("discussion", 3000))) if len(x) > 10)
    emit({"stage": "overview", "msg": "원고 개요 파악 (기여·주장)", "step": 1})
    raw = llm(OVERVIEW_SYS + "\n" + lang_rule(lang), ctx + "\n\n" + core, model, 0.2, on_token=lambda t: emit({"token": t}))
    summary = parse_overview(raw)

    # 2) 발췌별 지적
    chunks = make_chunks(doc, size, max_n)
    lo, hi = M.body_span(doc)
    covered = sum(c["end"] - c["start"] for c in chunks)
    coverage = round(100 * covered / max(1, sum(s["end"] - s["start"] for s in doc["sections"] if s["kind"] not in SKIP_KINDS) or hi - lo))
    if coverage < 95:
        emit({"stage": "note", "msg": f"원고가 길어 본문의 약 {coverage}%(발췌 {len(chunks)}개)를 고르게 골라 검토합니다 — 전체를 보려면 '정밀'"})
    sys_c = CHUNK_SYS.format(field=field, n=per, lang=lang_rule(lang))
    over = "[원고 개요 — 원고가 주장하는 것]\n" + "\n".join(f"- {c}" for c in summary["contrib"] + summary["claims"][:4] + summary["design"][:2])
    cache, found, dropped, raws = {}, [], [], [raw]
    for k, ch in enumerate(chunks):
        emit({"stage": "chunk", "msg": f"발췌 {k + 1}/{len(chunks)} 심사: {', '.join(ch['titles'])[:80]}", "step": 2 + k, "steps": len(chunks) + 2, "reset": True})
        body = flat(text[ch["start"]:ch["end"]])
        user = f"{ctx}\n\n{over}\n\n[이번 발췌 — {', '.join(ch['titles'])}]\n<<<\n{body}\n>>>"
        raw = llm(sys_c, user, model, 0.3, on_token=lambda t: emit({"token": t}))
        raws.append(raw)
        items = parse_issues(raw)[:per + 2]
        ok, bad = verify(doc, items, ch, cache, len(found))
        found += ok
        dropped += bad
        emit({"stage": "chunk_done", "msg": f"발췌 {k + 1}: 지적 {len(ok)}개 확인" + (f", {len(bad)}개 제외(인용 불일치·일반론)" if bad else "")})
    found = dedupe(found)  # 합칠 때는 모델이 먼저(중요하다고) 낸 것을 남긴다
    if found and opt.get("cross", True):
        found, bad = cross_check(doc, found, model, emit, lang)
        dropped += bad
        emit({"stage": "chunk_done", "msg": f"교차 확인: {len(bad)}개 제외" if bad else "교차 확인: 모두 유지"})
    found.sort(key=lambda x: (x["sev"] != "major", x["pos"]))
    nM = nm = 0
    for it in found:
        if it["sev"] == "major":
            nM += 1
            it["id"] = f"M{nM}"
        else:
            nm += 1
            it["id"] = f"m{nm}"
    found.sort(key=lambda x: (x["sev"] != "major", int(x["id"][1:])))

    # 3) 종합
    emit({"stage": "synth", "msg": "종합 판정·예상 질문", "step": len(chunks) + 2, "reset": True})
    listing = "\n".join(f"#{i['id']} [{i['sev']}/{i['cat']}] ({i['where']}) {i.get('title') or ''} — {i['why'][:220]}" for i in found[:40]) or "(확인된 지적 없음)"
    rule_issues = opt.get("_rules") or []
    rsum = {}
    for r in rule_issues:
        if r["sev"] != "info":
            rsum[r["cat"]] = rsum.get(r["cat"], 0) + 1
    user = (f"{ctx}\n\n[원고 개요]\n" + "\n".join(f"- {c}" for c in summary["contrib"] + summary["claims"]) +
            f"\n\n[지적 목록]\n{listing}\n\n[기계적 점검 결과(형식 문제)] " + (", ".join(f"{k} {v}건" for k, v in rsum.items()) or "없음"))
    raw = llm(SYNTH_SYS.format(lang=lang_rule(lang)), user, model, 0.3, on_token=lambda t: emit({"token": t}))
    raws.append(raw)
    verdict, questions = parse_synth(raw)
    for r in verdict["reasons"]:
        r["text"] = delatex(r["text"])
    for q in questions:
        q["q"], q["tip"] = delatex(q["q"]), delatex(q["tip"])
    ids = {i["id"] for i in found}
    for r in verdict["reasons"]:
        r["refs"] = [x for x in r["refs"] if x in ids]
    for q in questions:
        q["refs"] = [x for x in q["refs"] if x in ids]
    return {"summary": summary, "verdict": verdict, "questions": questions, "issues": found,
            "dropped": [{"quote": d.get("quote", "")[:200], "why": d.get("why", "")[:200], "reason": d["reason"], "title": d.get("title", ""),
                         "where": d.get("where", "")} for d in dropped],
            "chunks": [{"titles": c["titles"], "start": c["start"], "end": c["end"]} for c in chunks], "coverage": min(100, coverage),
            "lang": lang, "model": model, "elapsed": round(time.time() - t0), "options": {k: v for k, v in opt.items() if not k.startswith("_")},
            "depth": dname, "raw": raws}


# ── 이력 (WORKSPACE/history/<id>.json — 원고 하나 = 파일 하나) ────────────
def hist_path(hid):
    if not re.fullmatch(r"\d{8}-\d{6}-[0-9a-f]{4}", hid or ""):
        raise ValueError("잘못된 이력 id")
    return os.path.join(WS, "history", hid + ".json")


def hist_load(hid):
    return json.loads(read(hist_path(hid)))


def hist_save(rec):
    os.makedirs(os.path.join(WS, "history"), exist_ok=True)
    p = hist_path(rec["id"])
    with LOCK:
        with open(p + ".tmp", "w", encoding="utf-8") as f:
            json.dump(rec, f, ensure_ascii=False)
        os.replace(p + ".tmp", p)


def hist_list(limit=200):
    d = os.path.join(WS, "history")
    if not os.path.isdir(d):
        return []
    out = []
    for fn in sorted(os.listdir(d), reverse=True)[:limit]:
        if not fn.endswith(".json"):
            continue
        try:
            r = json.loads(read(os.path.join(d, fn)))
            rv = r.get("review") or {}
            out.append({"id": r["id"], "name": r["doc"]["name"], "title": r["doc"].get("title", ""), "ts": r["ts"],
                        "verdict": (rv.get("verdict") or {}).get("label", ""), "reviewed": bool(rv),
                        "n": {s: sum(1 for i in r["issues"] if i["sev"] == s) for s in ("major", "minor", "info")}})
        except Exception:
            pass
    return out


def hist_delete(hid):
    p = hist_path(hid)
    if os.path.exists(p):
        os.remove(p)


def client_doc(doc):
    """화면용 원고: 전문·쪽 시작 offset·절·캡션 (참고문헌 본문은 text 안에 있음)"""
    return {k: doc[k] for k in ("name", "via", "lang", "title", "text", "pages", "npages", "captions", "ref_style")} | \
        {"outline": M.outline(doc), "nrefs": len(doc["refs"])}


def ingest(name, data, opt=None):
    """업로드 → 추출·구조·기계적 점검 → 이력 저장. 원본 파일은 임시 폴더에서 바로 지운다."""
    name = os.path.basename(name or "manuscript.pdf")
    ext = os.path.splitext(name)[1].lower()
    if ext not in (".pdf", ".docx", ".hwp", ".hwpx", ".txt", ".md", ".tex"):
        raise ValueError(f"지원하지 않는 형식: {ext or '(확장자 없음)'} — PDF·DOCX·HWP·HWPX·TXT")
    with tempfile.TemporaryDirectory(prefix="review-") as d:
        p = os.path.join(d, "upload" + ext)
        with open(p, "wb") as f:
            f.write(data)
        doc = M.load(p, name)
    opt = opt or {}
    issues = M.mechanical(doc, opt.get("mtype", ""))
    for n, i in enumerate(issues, 1):
        i["id"] = f"R{n}"
    now = datetime.datetime.now()
    rec = {"id": f"{now:%Y%m%d-%H%M%S}-{secrets.token_hex(2)}", "ts": now.isoformat(timespec="seconds"), "doc": doc,
           "issues": issues, "review": None, "options": opt}
    hist_save(rec)
    return rec


def run_review(hid, opt, emit):
    rec = hist_load(hid)
    opt = {k: opt.get(k) for k in ("journal", "mtype", "depth", "lang", "model") if opt.get(k) is not None}
    if opt.get("mtype") != rec.get("options", {}).get("mtype"):  # 원고 유형이 바뀌면 구조 점검을 다시
        rules = M.mechanical(rec["doc"], opt.get("mtype", ""))
        for n, i in enumerate(rules, 1):
            i["id"] = f"R{n}"
    else:
        rules = [i for i in rec["issues"] if i["src"] == "rule"]
    if not GPU.acquire(blocking=False):
        emit({"stage": "wait", "msg": "다른 리뷰가 LLM 을 쓰는 중 — 순서를 기다립니다"})
        GPU.acquire()
    try:
        rv = review(rec["doc"], dict(opt, _rules=rules), emit, opt.get("model"))
    finally:
        GPU.release()
    rv["ts"] = datetime.datetime.now().isoformat(timespec="seconds")
    rec["issues"] = rules + rv.pop("issues")
    rec["review"] = rv
    rec["options"] = opt
    hist_save(rec)
    return rec


def client_rec(rec):
    rv = dict(rec["review"]) if rec.get("review") else None
    if rv:
        rv.pop("raw", None)
    return {"id": rec["id"], "ts": rec["ts"], "doc": client_doc(rec["doc"]), "issues": rec["issues"], "review": rv, "options": rec.get("options") or {}}


def meta():
    return {"mtypes": MTYPES, "depths": [{"id": k, "label": v[0]} for k, v in DEPTHS.items()], "langs": [{"id": k, "label": v} for k, v in LANGS.items()],
            "model": MODEL, "llm": f"{LLM_API} {LLM_BASE}", "tools": M.tools_status()}


# ── HTTP ───────────────────────────────────────────────────────────────
HTML = read(os.path.join(ROOT, "ui.html")) if os.path.exists(os.path.join(ROOT, "ui.html")) else "ui.html 없음"

# ── 저작권 표기 (LICENSE·NOTICE 참고) ─────────────────────────────────────
_SIG = __import__("base64").b64decode("wqkgMjAyNiBnZ2dnODY1NyDCtyBkb25nanVraW0uZGV2QGdtYWlsLmNvbQ==").decode()
_SIG_A = __import__("base64").b64decode("Z2dnZzg2NTcgPGRvbmdqdWtpbS5kZXZAZ21haWwuY29tPg==").decode()


def signed(html):
    """화면에 저작권 표기를 붙인다. ui.html 에서 지워져도 서버가 내보낼 때 다시 붙는다."""
    name, mail = _SIG.split(" · ")
    if 'name="author"' not in html:
        meta = f'<meta name="author" content="{name[7:]} <{mail}>">'
        html = html.replace("<head>", "<head>" + meta, 1) if "<head>" in html else meta + html
    if "data-sig" not in html:
        tag = (f'<!-- {_SIG} --><div data-sig title="{mail}" style="text-align:center;font-size:11px;color:#9aa0a6;'
               f'opacity:.55;margin:28px 0 8px">{name}</div>')
        html = html.replace("</body>", tag + "</body>", 1) if "</body>" in html else html + tag
    return html


class H(BaseHTTPRequestHandler):
    def log_message(self, fmt, *a):
        if "/api/run" in (a[0] if a else "") or "/api/upload" in (a[0] if a else ""):
            super().log_message(fmt, *a)

    def _send(self, body, ctype="application/json", code=200, extra=None):
        b = body if isinstance(body, bytes) else json.dumps(body, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("X-Author", _SIG_A)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(b)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(b)

    def do_GET(self):
        path = self.path.split("?")[0]
        try:
            if path == "/api/health":
                return self._send({"ok": True})
            if path == "/api/meta":
                return self._send(meta())
            if path == "/api/models":
                return self._send(models())
            if path == "/api/history":
                return self._send(hist_list())
            m = re.fullmatch(r"/api/history/([\w-]+)", path)
            if m:
                return self._send(client_rec(hist_load(m.group(1))))
            m = re.fullmatch(r"/api/export/([\w-]+)\.(md|txt|docx)", path)
            if m:
                rec = hist_load(m.group(1))
                base = re.sub(r"[^\w가-힣.-]+", "_", os.path.splitext(rec["doc"]["name"])[0])[:60] + "_review"
                fn = f"{base}.{m.group(2)}"
                disp = {"Content-Disposition": f"attachment; filename=\"review.{m.group(2)}\"; filename*=UTF-8''{urllib.request.quote(fn)}"}
                if m.group(2) == "docx":
                    return self._send(export.to_docx(rec), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", extra=disp)
                body = export.to_md(rec) if m.group(2) == "md" else export.to_txt(rec)
                return self._send(body.encode(), ("text/markdown" if m.group(2) == "md" else "text/plain") + "; charset=utf-8", extra=disp)
            self._send(signed(HTML).encode(), "text/html; charset=utf-8")
        except (FileNotFoundError, ValueError):
            self._send({"error": "없음"}, code=404)
        except Exception as e:
            self._send({"error": f"{type(e).__name__}: {e}"}, code=500)

    def do_POST(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n > MAX_UPLOAD * 1.4:
            return self._send({"error": f"파일이 너무 큽니다 (최대 {MAX_UPLOAD // 1024 // 1024}MB)"}, code=413)
        try:
            req = json.loads(self.rfile.read(n) or b"{}")
        except ValueError:
            return self._send({"error": "잘못된 요청"}, code=400)
        try:
            if self.path == "/api/history/delete":
                hist_delete(req.get("id"))
                return self._send({"ok": True})
            if self.path == "/api/upload":
                data = base64.b64decode(req.get("data") or "")
                if not data:
                    raise ValueError("파일이 비어 있습니다")
                return self._send(client_rec(ingest(req.get("name"), data, {"mtype": req.get("mtype") or ""})))
        except (ValueError, RuntimeError) as e:
            return self._send({"error": str(e)}, code=400)
        except Exception as e:
            return self._send({"error": f"{type(e).__name__}: {e}"}, code=500)
        if self.path != "/api/run":
            return self._send({"error": "없는 경로"}, code=404)
        self.send_response(200)
        self.send_header("X-Author", _SIG_A)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()

        def emit(ev):
            self.wfile.write(f"data: {json.dumps(ev, ensure_ascii=False)}\n\n".encode())
            self.wfile.flush()

        try:
            emit({"done": client_rec(run_review(req.get("id"), req, emit))})
        except (BrokenPipeError, ConnectionResetError):
            pass
        except Exception as e:
            try:
                emit({"error": str(e) if isinstance(e, (ValueError, RuntimeError, FileNotFoundError)) else f"{type(e).__name__}: {e}"})
            except OSError:
                pass


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--cli":
        path = sys.argv[2]
        depth = sys.argv[3] if len(sys.argv) > 3 else "quick"
        lang = sys.argv[4] if len(sys.argv) > 4 else "auto"
        with open(path, "rb") as f:
            rec = ingest(path, f.read(), {"mtype": os.environ.get("MTYPE", "연구논문")})

        def tok(ev):
            if "token" in ev:
                print(ev["token"], end="", file=sys.stderr, flush=True)
            elif "msg" in ev:
                print(f"\n\n== {ev['msg']}", file=sys.stderr, flush=True)
        rec = run_review(rec["id"], {"depth": depth, "lang": lang, "mtype": os.environ.get("MTYPE", "연구논문"), "journal": os.environ.get("JOURNAL", "")}, tok)
        print(export.to_md(rec))
        print(f"\n[이력 {rec['id']} · 제외된 지적 {len(rec['review']['dropped'])}개 · {rec['review']['elapsed']}초]", file=sys.stderr)
        sys.exit(0)
    print(f"review local → http://localhost:{PORT}  (llm={LLM_API} {LLM_BASE} {MODEL}, workspace={WS})  {_SIG}")
    ThreadingHTTPServer(("", PORT), H).serve_forever()
