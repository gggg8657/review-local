"""portal_jobs — 도구가 포털 '작업 알림 센터'(🔔)에 작업 상태를 알리는 작은 도우미. 표준 라이브러리만, 실패해도 도구 동작에 영향 없음.

포털이 도구를 시작할 때 넘기는 환경변수(없으면 아무 일도 안 함 — 포털 밖에서 띄운 도구·시험):
  PORTAL_JOBS_URL    http://127.0.0.1:8700/api/internal/jobs  (다른 컨테이너 도구는 그쪽에서 이 값을 직접 지정)
  PORTAL_JOBS_TOKEN  이 도구 전용 토큰(포털 비밀값으로 만든 HMAC — 로그·화면에 찍지 말 것)
  PORTAL_JOBS_TOOL   도구 폴더 이름(예: protein-local). 없으면 notify(tool=...) 로 넘김
쓰는 법:
  import portal_jobs
  portal_jobs.notify(user, "running", id=jid, title="유비퀴틴 구조 예측", link="/?job=" + jid)   # user = 요청의 X-Agent-User
  portal_jobs.notify(user, "done", id=jid, title=..., link=..., message="pLDDT 87")              # done | failed | canceled | queued | running
보내기는 백그라운드 스레드 하나가 순서대로(요청 경로를 막지 않음), 한 번에 2초 타임아웃, 실패는 조용히 버림(대기열 200개 넘으면 버림).
Copyright (c) 2026 gggg8657 (DongJu Kim). 이 파일은 각 도구 저장소에 그대로 복사해 씀(원본: agent-page-portal/sdk/portal_jobs.py)."""
import json, os, queue, threading, urllib.request

STATUSES = ("queued", "running", "done", "failed", "canceled")
TIMEOUT = 2.0
_Q = queue.Queue(maxsize=200)
_T = {"t": None}
_LOCK = threading.Lock()


def enabled():
    return bool(os.environ.get("PORTAL_JOBS_URL") and os.environ.get("PORTAL_JOBS_TOKEN"))


def _post(url, token, payload):
    req = urllib.request.Request(url, data=json.dumps(payload, ensure_ascii=False).encode(), method="POST",
                                 headers={"Content-Type": "application/json", "X-Portal-Jobs-Token": token})
    with urllib.request.urlopen(req, timeout=TIMEOUT) as r:   # 프록시 환경변수가 있어도 127.0.0.1 은 보통 직접 — 실패하면 버림
        r.read()


def _worker():
    while True:
        url, token, payload = _Q.get()
        try:
            _post(url, token, payload)
        except Exception:
            pass
        finally:
            _Q.task_done()


def notify(user, status, id=None, title="", link="", message="", progress=None, tool=None):
    """작업 상태 알림 — 켜져 있으면 대기열에 넣고 True, 꺼져 있거나 값이 틀리면 False. 예외를 내지 않음"""
    try:
        url, token = os.environ.get("PORTAL_JOBS_URL"), os.environ.get("PORTAL_JOBS_TOKEN")
        tool = tool or os.environ.get("PORTAL_JOBS_TOOL")
        if not (url and token and tool and user and status in STATUSES):
            return False
        p = {"tool": tool, "user": str(user), "status": status, "title": str(title or "")[:200], "link": str(link or "")[:500], "message": str(message or "")[:500]}
        if id is not None:
            p["id"] = str(id)[:80]
        if progress is not None:
            p["progress"] = max(0, min(100, int(progress)))
        with _LOCK:
            if _T["t"] is None or not _T["t"].is_alive():
                _T["t"] = threading.Thread(target=_worker, name="portal_jobs", daemon=True)
                _T["t"].start()
        _Q.put_nowait((url, token, p))
        return True
    except Exception:
        return False


def flush(timeout=5.0):
    """대기열이 빌 때까지 기다림(시험용). 다 보냈으면 True"""
    import time
    end = time.time() + timeout
    while _Q.unfinished_tasks and time.time() < end:
        time.sleep(0.05)
    return not _Q.unfinished_tasks
