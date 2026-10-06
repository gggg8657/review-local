#!/usr/bin/env python3
"""LLM 없이 검증 (가짜 LLM): 합성 원고(영문 TXT·국문 DOCX)에 일부러 심은 문제를 기계적 점검이 모두 잡는지
→ 절·캡션·참고문헌 인식 → 인용 찾기(공백·하이픈·부분 일치, 지어낸 인용은 못 찾음) → 리뷰 파이프라인(지어낸 인용·일반론 버림,
같은 지적 합치기, LaTeX 정리, 판정·질문 파싱) → MD·TXT·DOCX → HTTP(업로드·SSE·내보내기·이력) → ui.html 외부 자산 없음.
WORKSPACE 는 임시 폴더로 바꿔 실데이터 폴더에 흔적을 남기지 않는다.   python3 selftest.py"""
import base64, io, json, os, re, shutil, sys, tempfile, threading, urllib.request, zipfile

TMP = tempfile.mkdtemp(prefix="review-selftest-")
os.environ["WORKSPACE"] = TMP
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import app, export, manuscript as M  # noqa: E402

# ── 합성 원고 1: 영문, 쪽 나눔(\f), 심어 둔 문제 ─────────────────────────
EN = """Deep Learning for Early Fault Detection in Reactor Coolant Pumps

Jane Doe, John Roe
Korea Atomic Energy Research Institute, Daejeon, Korea

Abstract
We propose a long short-term memory (LSTM) network for early fault detection in reactor coolant pumps.
Trained on 12,400 hours of plant data, the model reaches a detection accuracy of 94.2% and an F1 score of 0.88,
outperforming a support vector machine baseline. The method raises alarms 35 minutes earlier than threshold alarms.

1. Introduction
Pump failures are a leading cause of unplanned outages [1]. Threshold alarms miss slow degradation [2, 3].
Data-driven methods based on SVM classifiers have been explored [4]. Later, convolutional networks
were adopted; a CNN extracts local patterns, and convolutional neural network (CNN) variants dominate recent work [5].
The overall framework is shown in Fig. 1.

2. Methods
2.1 Data
We used 24 sensor channels sampled at 1 Hz from two units [6]. The dataset is summarised in Table 1.
\f
Table 1. Sensor channels and units.
Channel Unit
Fig. 1. Overall framework of the proposed detector.
2.2 Model
The prediction error is defined as
e_t = || x_t - y_t ||   (1)
and the anomaly score as
s_t = mean(e_t)   (2)
The loss in Eq. (7) is minimised with Adam [7]. The data-set was split by time; the dataset contains 3 faults.

3. Results
The detection accuracy of 91.7% was obtained on the held-out year. Detection delay is shown in Fig. 5.
Our method raised alarms 35 minutes earlier on average [9].
Fig. 3. Anomaly score over time with maintenance records.
Table 2. Comparison with baselines.

4. Conclusion
The LSTM detector achieved high accuracy and can be applied to every plant without retraining.

References
[1] S. Lee, Pump failure analysis, Nucl. Eng. Technol. 50 (2018) 1-10.
[2] J. Park, Vibration monitoring, Ann. Nucl. Energy 120 (2019) 33-41.
[3] K. Kim, Threshold alarms, Prog. Nucl. Energy 101 (2020) 5-12.
[4] H. Choi, SVM fault diagnosis, IEEE Access 8 (2020) 1000-1010.
[5] A. Han, CNN diagnosis, Nucl. Eng. Des. 370 (2020) 1-9.
[6] B. Seo, Plant data, KNS (2021).
[7] D. Kingma, J. Ba, Adam, ICLR (2015).
[8] T. Uncited, Never cited survey, ACM Comput. Surv. 54 (2021) 1-38.
"""

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def make_docx(paras):
    """stdlib 로 만든 최소 DOCX — (스타일, 글) 목록. '#PAGE' 는 Word 의 쪽 나눔 표시"""
    body = []
    for style, text in paras:
        if style == "#PAGE":
            body.append(f'<w:p><w:r><w:lastRenderedPageBreak/><w:t>{text}</w:t></w:r></w:p>')
            continue
        ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
        body.append(f'<w:p>{ppr}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>')
    xml = f'<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="{W}"><w:body>{"".join(body)}</w:body></w:document>'
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("[Content_Types].xml", export.CT)
        z.writestr("_rels/.rels", export.RELS)
        z.writestr("word/document.xml", xml)
    return buf.getvalue()


KO = make_docx([
    ("Title", "LSTM 기반 원자로 냉각재 펌프 이상 징후 조기 탐지"), ("", "홍길동 / 한국원자력연구원"),
    ("Heading1", "초록"), ("", "원자로 냉각재 펌프(RCP)의 이상을 탐지하는 장단기 메모리(LSTM) 모델을 제안한다. 이상 탐지 정확도 97.3%를 달성하였다."),
    ("Heading1", "1. 서론"), ("", "RCP 고장은 원자로 정지로 이어진다[1]. 최근 CNN 기반 진단이 보고되었다[2]."),
    ("Heading1", "2. 연구 방법"), ("", "데이터는 표 1과 같다. 모델 구조는 그림 1에 나타내었다."),
    ("", "표 1. 데이터 구성"), ("", "그림 1. 모델 구조"),
    ("#PAGE", "3. 결과"), ("", "시험 데이터에서 이상 탐지 정확도는 95.1%였다. 시간 변화는 그림 4에 보였다."),
    ("", "그림 2. 이상 점수 시계열"),
    ("Heading1", "4. 결론"), ("", "제안 모델은 모든 원전에 바로 적용할 수 있다."),
    ("Heading1", "참고문헌"), ("", "[1] S. Lee, Pump failure, NET 50 (2018)."), ("", "[2] H. Choi, CNN diagnosis, IEEE Access 8 (2020)."),
    ("", "[3] T. Seo, Uncited survey, ACM CSUR 54 (2021)."),
])

SEEN = []


def fake(system, user, model=None, temperature=0.3, on_token=lambda t: None):
    """프롬프트 종류를 보고 정해진 답을 흘려보낸다(토큰 스트리밍 흉내)."""
    SEEN.append((system, user))
    if "편집위원" in system:  # 교차 확인: 첫 지적은 유지, 둘째는 다른 곳에서 해소
        assert "[다른 곳 문단]" in user and "#X2" in user, user
        out = "#X1: 유지\n#X2: 해소 — split by time"
    elif "[기여]" in system:
        out = "[기여]\n1. LSTM 으로 펌프 이상을 조기 탐지\n2. 94.2% 정확도\n3. 35분 조기 경보\n[핵심 주장]\n- SVM 보다 우수\n[연구 설계]\n- 2개 호기 24채널"
    elif "[판정]" in system:
        out = ("[판정]\nMajor Revision\n[이유]\n- 시간 분할 근거와 기준선 비교가 약함 (#M1, #M9)\n- 일반화 주장이 과도 (#m1)\n[질문]\n"
               + "\n".join(f"{n}. 질문: 질문 {n} — \\(d_k\\) 와 $\\mu$ 는? (#M1)\n   준비: 준비 {n}" for n in range(1, 7)))
    else:
        out = ("### 지적\n심각도: major\n분류: 방법론\n요지: 기준선 비교의 통계적 근거 부족\n인용: outperforming a support vector machine baseline.\n"
               "문제: SVM 대비 우수하다고 하지만 반복 실험·검정이 없다 — $\\hat{x}_t$ 기준.\n수정: 5회 반복 평균±표준편차와 검정 결과를 표로 제시하라.\n검색어: baseline, repeated, significance\n\n"
               "### 지적\n심각도: major\n분류: 방법론\n요지: 기준선 비교의 통계적 근거가 부족함\n인용: the model reaches a detection accuracy of 94.2% and an F1 score\n"
               "문제: SVM 대비 우수하다고 하지만 반복 실험과 통계 검정이 없다.\n수정: 반복 실험 결과를 제시하라.\n\n"
               "### 지적\n심각도: minor\n분류: 결론 과대해석\n요지: 일반화 주장\n인용: can be applied to every plant without retraining.\n"
               "문제: 두 호기 데이터만으로 모든 원전 적용을 주장한다.\n수정: 적용 범위를 학습 호기로 한정하거나 타 호기 검증을 추가하라.\n검색어: units, plant, split\n\n"
               "### 지적\n심각도: major\n분류: 데이터\n요지: 지어낸 인용\n인용: We randomly shuffled all windows before splitting into train and test sets.\n"
               "문제: 누수.\n수정: 시간 분할.\n\n"
               "### 지적\n심각도: minor\n분류: 명확성\n요지: 일반론\n인용: The overall framework is shown in Fig. 1.\n문제: 영어 교정이 필요하다.\n수정: 원어민 교정을 받으라.\n")
        if "<<<" not in user:
            out = "없음"
    for i in range(0, len(out), 9):
        on_token(out[i:i + 9])
    return out


app.llm = fake
try:
    assert app.WS == TMP, "WORKSPACE 가 임시 폴더가 아님"

    # 1) 영문 TXT: 구조 인식
    p = os.path.join(TMP, "paper.txt")
    with open(p, "w", encoding="utf-8") as f:
        f.write(EN)
    d = M.load(p)
    assert d["lang"] == "en" and d["npages"] == 2 and d["title"].startswith("Deep Learning for Early Fault"), (d["title"], d["npages"])
    kinds = [s["kind"] for s in d["sections"]]
    for k in ("front", "abstract", "intro", "method", "results", "conclusion", "references"):
        assert k in kinds, (k, kinds)
    assert [s["title"] for s in d["sections"] if s["level"] == 2] == ["2.1 Data", "2.2 Model"], d["sections"]
    caps = {(c["type"], c["num"]) for c in d["captions"]}
    assert caps == {("tab", 1), ("tab", 2), ("fig", 1), ("fig", 3)}, caps   # 'Fig. 5' 언급·'Fig. 1.' 언급은 캡션 아님
    assert d["ref_style"] == "numeric" and [r["n"] for r in d["refs"]] == list(range(1, 9))
    assert M.page_of(d, d["text"].index("Table 1. Sensor")) == 2 and M.page_of(d, 10) == 1

    # 2) 기계적 점검: 심은 문제를 모두 잡고, 정상인 것은 건드리지 않는다
    iss = M.mechanical(d, "연구논문")
    T = [i["title"] for i in iss]
    have = lambda s: any(s in t for t in T)
    assert have("Fig. 3 이(가) 본문에서 언급되지 않음") and have("Table 2 이(가) 본문에서 언급되지 않음"), T
    assert have("없는 그림 번호 언급: 그림 5"), T
    assert not have("Fig. 1 이(가)") and not have("Table 1 이(가)"), T
    assert have("그림 번호 건너뜀: 2"), T
    assert have("목록에 없는 참고문헌 번호 인용: [9]") and have("본문에서 인용되지 않은 참고문헌: [8]"), T
    assert not any(re.search(r"인용되지 않은 참고문헌: \[[1-7]\]", t) for t in T), T
    assert have("약어 정의 없음: SVM") and have("약어를 정의 전에 사용: CNN"), T
    assert not have(": LSTM") and not have(": RCP"), T                    # 처음에 정의한 약어
    assert have("없는 수식 번호 언급: (7)"), T
    mm = next(i for i in iss if i["cat"] == "수치 불일치")
    assert "94.2" in mm["title"] and "91.7" in mm["title"] and mm["sev"] == "major" and mm["pos2"] > mm["pos"], mm
    assert not any("35" in t and "수치" in t for t in T), T              # 35분은 본문에도 있음
    assert have("용어 표기 혼용: 'data-set'"), T
    for i in iss:
        assert i["src"] == "rule" and i["sev"] in ("major", "minor", "info") and isinstance(i["pos"], int)
        if i["quote"] and i["len"]:
            assert i["pos"] <= len(d["text"]), i

    # 3) 국문 DOCX (stdlib 추출, 쪽 표시, 조사 붙은 수치 문맥)
    p2 = os.path.join(TMP, "ko.docx")
    with open(p2, "wb") as f:
        f.write(KO)
    k = M.load(p2)
    assert k["lang"] == "ko" and k["title"].startswith("LSTM 기반") and k["npages"] == 2, (k["title"], k["npages"])
    kt = [i["title"] for i in M.mechanical(k)]
    for want in ("그림 2 이(가) 본문에서 언급되지 않음", "없는 그림 번호 언급: 그림 4", "약어 정의 없음: CNN",
                 "본문에서 인용되지 않은 참고문헌: [3]", "초록 97.3% ↔ 본문 95.1%"):
        assert any(want in t for t in kt), (want, kt)
    assert not any("표 1 이(가)" in t or ": LSTM" in t or ": RCP" in t for t in kt), kt

    # 4) 인용 찾기: 공백·줄바꿈·하이픈·따옴표 차이는 찾고, 지어낸 문장은 못 찾는다
    t = d["text"]
    a, ln, r = M.locate(d, "outperforming   a support\nvector machine baseline.")
    assert t[a:a + ln].startswith("outperforming") and r == 1.0
    a, ln, r = M.locate(d, "the model reaches a detection accuracy of 94.2 % and an F1-score of 0.88")
    assert t[a:a + ln].startswith("the model reaches") and r >= 0.75, r
    assert M.locate(d, "We randomly shuffled all windows before splitting into train and test sets.") is None
    assert M.locate(d, "짧음") is None

    # 5) 발췌 나누기: 참고문헌·제목 블록 제외, 순서·겹침 없음
    ch = app.make_chunks(d, 300, 50)
    assert len(ch) >= 3 and all(c["end"] > c["start"] for c in ch)
    assert all(ch[n]["start"] >= ch[n - 1]["end"] for n in range(1, len(ch)))
    ref0 = next(s["start"] for s in d["sections"] if s["kind"] == "references")
    assert ch[0]["start"] > 0 and ch[-1]["end"] <= ref0
    assert len(app.make_chunks(d, 300, 2)) <= 2

    # 6) 리뷰 파이프라인 (가짜 LLM)
    events = []
    rec = app.ingest("paper.txt", EN.encode(), {"mtype": "연구논문"})
    rec = app.run_review(rec["id"], {"depth": "quick", "lang": "auto", "journal": "Nucl. Eng. Technol.", "mtype": "연구논문"}, events.append)
    rv = rec["review"]
    llm_iss = [i for i in rec["issues"] if i["src"] == "llm"]
    assert rv["summary"]["contrib"][0].startswith("LSTM") and len(rv["summary"]["contrib"]) == 3
    assert [i["id"] for i in llm_iss] == ["M1"], [(i["id"], i["title"]) for i in llm_iss]   # 중복 합침, 지어낸 인용·일반론·교차 확인 해소 버림
    m1 = llm_iss[0]
    assert m1["also"] and "94.2%" in m1["also"][0]["quote"], m1            # 합친 지적의 다른 위치
    assert "$" not in m1["why"] and "x̂_t" in m1["why"], m1["why"]       # LaTeX 정리
    assert t[m1["pos"]:m1["pos"] + m1["len"]].startswith("outperforming") and m1["where"].endswith("Abstract")
    reasons = {x["reason"] for x in rv["dropped"]}
    assert reasons >= {"원고에서 인용을 찾지 못함", "일반론", "교차 확인: 원고 다른 곳에서 해소 — split by time"}, rv["dropped"]
    assert "추출 문제" in next(s for s, u in SEEN if "### 지적" in s)       # 깨진 수식·기호는 지적 금지
    assert rv["verdict"]["label"] == "Major Revision" and rv["verdict"]["reasons"][0]["refs"] == ["M1"]   # 없는 #M9 는 뺌
    assert rv["verdict"]["reasons"][1]["refs"] == []                       # 버린 지적(#m1) 참조도 뺌
    assert len(rv["questions"]) == 5 and rv["questions"][0]["tip"] == "준비 1" and "$" not in rv["questions"][0]["q"]
    assert any(e.get("stage") == "chunk" for e in events) and any("token" in e for e in events)
    sys_chunk = next(s for s, u in SEEN if "### 지적" in s)
    assert "Nucl. Eng. Technol." in sys_chunk and "영어" in sys_chunk  # 영문 원고 → 영어로 설명
    assert any("[대상 저널·분야] Nucl. Eng. Technol." in u for s, u in SEEN)
    rules_after = [i for i in rec["issues"] if i["src"] == "rule"]
    assert rules_after and all(i["id"].startswith("R") for i in rules_after)
    # 리뷰 언어 지정: 국문
    SEEN.clear()
    app.run_review(rec["id"], {"depth": "deep", "lang": "ko", "mtype": "리뷰(총설)"}, lambda e: None)
    assert all("한국어로 쓴다" in s for s, u in SEEN if "### 지적" in s)
    assert not any("찾지 못한 절" in i["title"] and "방법" in i["title"] for i in app.hist_load(rec["id"])["issues"])

    # 7) 파싱 견고성
    its = app.parse_issues("머리말\n### 지적\n**심각도**: Major\n**분류**: 통계\n**인용**: \"some quote here\"\n**문제**: 줄이\n이어짐\n수정: 고쳐라\n")
    assert its[0]["sev"] == "major" and its[0]["quote"] == "some quote here" and its[0]["why"] == "줄이 이어짐", its
    assert app.parse_issues("없음") == []
    v, q = app.parse_synth("[판정]\n소폭 수정 (Minor Revision)\n[이유]\n- a (#m2)\n[질문]\n1. 첫 질문?\n준비: 팁\n2. Question: second?\n   Tip: t2")
    assert v["label"] == "Minor Revision" and [x["q"] for x in q] == ["첫 질문?", "second?"] and q[1]["tip"] == "t2", (v, q)

    # 8) 내보내기
    full = app.hist_load(rec["id"])
    md = export.to_md(full)
    assert "## 예상 판정: Major Revision" in md
    assert "### M1." in md and "기계적 점검" in md and "예상 심사위원 질문" in md and "outperforming a support vector machine" in md
    txt = export.to_txt(full)
    assert "**" not in txt and "M1." in txt
    z = zipfile.ZipFile(io.BytesIO(export.to_docx(full)))
    docxml = z.read("word/document.xml").decode()
    assert "M1." in docxml and "Heading1" in docxml and z.testzip() is None
    import xml.dom.minidom
    xml.dom.minidom.parseString(docxml)                                   # 올바른 XML

    # 9) HTTP: 업로드 → SSE 리뷰 → 내보내기 → 이력 → 삭제
    srv = app.ThreadingHTTPServer(("127.0.0.1", 0), app.H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    base = f"http://127.0.0.1:{srv.server_address[1]}"
    r = urllib.request.urlopen(base + "/")
    html = r.read().decode()
    assert "data-sig" in html and 'name="author"' in html and r.headers["X-Author"]

    def post(path, body):
        return urllib.request.urlopen(urllib.request.Request(base + path, json.dumps(body).encode(), {"Content-Type": "application/json"}))
    up = json.load(post("/api/upload", {"name": "ko.docx", "data": base64.b64encode(KO).decode(), "mtype": "연구논문"}))
    assert up["doc"]["lang"] == "ko" and up["review"] is None and any(i["cat"] == "수치 불일치" for i in up["issues"])
    assert up["doc"]["outline"] and "text" in up["doc"]
    try:
        post("/api/upload", {"name": "x.exe", "data": base64.b64encode(b"MZ").decode()})
        raise AssertionError("exe 업로드가 통과함")
    except urllib.error.HTTPError as e:
        assert e.code == 400 and "지원하지 않는 형식" in json.load(e)["error"]
    evs = [json.loads(l[6:]) for l in post("/api/run", {"id": up["id"], "depth": "quick", "lang": "auto"}).read().decode().split("\n\n") if l.startswith("data: ")]
    assert "done" in evs[-1] and evs[-1]["done"]["review"]["verdict"]["label"] == "Major Revision", evs[-1]
    assert "raw" not in evs[-1]["done"]["review"]
    r = urllib.request.urlopen(f"{base}/api/export/{up['id']}.docx")
    assert r.headers["Content-Type"].startswith("application/vnd.openxmlformats") and "attachment" in r.headers["Content-Disposition"]
    assert urllib.request.urlopen(f"{base}/api/export/{up['id']}.md").read().decode().startswith("# 투고 전 사전 심사 리뷰")
    hist = json.load(urllib.request.urlopen(base + "/api/history"))
    assert {h["id"] for h in hist} >= {up["id"], rec["id"]} and next(h for h in hist if h["id"] == up["id"])["verdict"] == "Major Revision"
    post("/api/history/delete", {"id": up["id"]})
    assert up["id"] not in {h["id"] for h in json.load(urllib.request.urlopen(base + "/api/history"))}
    try:
        urllib.request.urlopen(base + "/api/history/not-an-id")
        raise AssertionError("잘못된 이력 id 가 통과함")
    except urllib.error.HTTPError as e:
        assert e.code == 404
    srv.shutdown()
    assert all(fn.endswith(".json") for fn in os.listdir(os.path.join(TMP, "history")))   # 원본 파일은 남기지 않음

    # 10) ui.html: 외부 CDN·API 없음
    ui = open(os.path.join(app.ROOT, "ui.html"), encoding="utf-8").read()
    assert not re.search(r"""(?:src|href)\s*=\s*["']https?://""", ui) and "fetch('http" not in ui
    print("selftest OK")
finally:
    shutil.rmtree(TMP, ignore_errors=True)
