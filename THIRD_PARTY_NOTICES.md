# 제3자 구성요소 고지 (THIRD_PARTY_NOTICES)

review-local 의 원저작물은 DongJu Kim(gggg8657)이 작성했습니다(`NOTICE`, `LICENSE`). 아래 제3자 구성요소는 각자의 라이선스를 따르며, 이 저장소의 MIT 라이선스로 바뀌지 않습니다.
조사일: 2026-10-12 (실제 파일·import·setup.sh 기준). "상업 이용" 칸은 확인한 라이선스 문구를 요약한 것이며 법률 판단이 아닙니다.

이 저장소는 파이썬 표준 라이브러리만 쓰며 제3자 코드·글꼴·데이터·가중치를 **동봉하지 않습니다**. DOCX 읽기·쓰기는 `zipfile` + XML 로 직접 처리합니다.

| 이름 | 버전 | 쓰임 | 라이선스 | 링크 | 상업 이용 | 비고 |
|---|---|---|---|---|---|---|
| Python 표준 라이브러리 | 3.x | 실행 환경 | PSF-2.0 | [python.org](https://docs.python.org/3/license.html) | 가능 | |
| poppler-utils `pdftotext` | 시스템 설치본 | 외부 프로그램 (PDF 텍스트 추출, `manuscript.py`) | GPL-2.0 / GPL-3.0 | [poppler](https://poppler.freedesktop.org/) | 가능 | 별도 프로세스로 호출만 함(링크·동봉 안 함). 없으면 kordoc 로 대체 |
| kordoc | kordoc-local 설치본 | 외부 프로그램 (HWP·HWPX·PDF 읽기, `node` 로 호출) | MIT (Copyright (c) 2026 chrisryugj) | [chrisryugj/kordoc](https://github.com/chrisryugj/kordoc) | 가능 | 같은 저작자의 kordoc-local 에 설치된 것을 씀, 동봉 안 함 |
| Node.js | 시스템 설치본 | 외부 프로그램 (kordoc 실행) | MIT | [nodejs.org](https://github.com/nodejs/node/blob/main/LICENSE) | 가능 | |
| Ollama | — | 외부 서비스 (LLM 서버) | MIT | [ollama/ollama](https://github.com/ollama/ollama) | 가능 | 동봉 안 함. 없으면 기계적 점검만 함 |
| Google Gemma 4 `gemma4:31b` | — | 외부 서비스가 돌리는 모델 (기본 배포) | Apache-2.0 | [google/gemma-4-31B-it](https://huggingface.co/google/gemma-4-31B-it) | 가능 | 가중치 동봉 안 함. `LLM_MODEL` 로 다른 모델 지정 가능 |

## 확인 필요

없음 (2026-10-12 조사 기준). 새 구성요소를 넣으면 이 표와 `NOTICE` 를 함께 갱신하세요.
