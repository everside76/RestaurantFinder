# -*- coding: utf-8 -*-
"""
식품안전나라 오픈 API 음식점 정보 조회 대시보드 (Windows GUI)

- ERP 대시보드 스타일 UI (다크 사이드바 + 카드형 레이아웃 + KPI 카드)
- 인증키는 Windows DPAPI로 암호화되어 exe 옆 config.json에 저장
- 검색 조건 최대 3개 AND 조합 + 서버 응답을 클라이언트에서 재검증하여
  조건에 맞는 결과만 표시
"""

import base64
import ctypes
import ctypes.wintypes
import json
import re
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# 정식 인증키는 HTTP 호출 시 오류가 발생하므로 반드시 HTTPS 사용
BASE_URL = "https://openapi.foodsafetykorea.go.kr/api"

# PyInstaller 실행 파일에서는 __file__이 임시 폴더를 가리키므로
# 설정 파일은 exe가 있는 폴더에 저장한다.
if getattr(sys, "frozen", False):
    CONFIG_PATH = Path(sys.executable).with_name("config.json")
else:
    CONFIG_PATH = Path(__file__).with_name("config.json")


def resource_path(name: str) -> Path:
    """번들 리소스(아이콘 등) 경로. PyInstaller onefile은 sys._MEIPASS에 풀린다."""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return Path(base) / name
    return Path(__file__).with_name(name)

MAX_CONDITIONS = 3   # 조합 검색 조건 최대 개수
PAGE_SIZE = 1000     # 한 번의 API 호출로 수신하는 건수 (API 최대치)
MAX_TREE_ROWS = 5000 # 표에 표시할 최대 행 수 (성능 보호, 저장은 전체)
CONFIRM_PAGES = 20   # 이 페이지 수(2만 건)를 넘으면 전체 수신 여부 확인

FONT = "맑은 고딕"

# NexaERP 스타일 색상 팔레트
C = {
    "sidebar":      "#1E1B4B",  # 다크 네이비
    "sidebar_hi":   "#312E81",  # 사이드바 호버/선택
    "sidebar_tx":   "#C7D2FE",  # 사이드바 텍스트
    "sidebar_dim":  "#818CF8",  # 사이드바 보조 텍스트
    "accent":       "#4F46E5",  # 인디고 액센트
    "accent_dark":  "#4338CA",
    "bg":           "#F4F5FB",  # 본문 배경
    "card":         "#FFFFFF",  # 카드
    "border":       "#E5E7EB",  # 카드 테두리
    "text":         "#111827",  # 본문 텍스트
    "muted":        "#6B7280",  # 보조 텍스트
    "green":        "#10B981",
    "red":          "#EF4444",
    "amber":        "#F59E0B",
}

SERVICES = {
    "I1200": {
        "name": "식품접객업정보",
        "desc": "전국 음식점 인허가 (약 313만 건)",
        "columns": {
            "BSSH_NM": "업소명",
            "PRSDNT_NM": "대표자명",
            "LOCP_ADDR": "소재지 주소",
            "INDUTY_NM": "업종",
            "TELNO": "전화번호",
            "PRMS_DT": "인허가일자",
            "CLSBIZ_DT": "폐업일자",
            "INSTT_NM": "관할기관",
            "LCNS_NO": "인허가번호",
        },
        "search_fields": {
            "업소명": "BSSH_NM",
            "소재지 주소": "LOCP_ADDR",
            "대표자명(김** 형태)": "PRSDNT_NM",
            "업종": "INDUTY_NM",
            "관할기관": "INSTT_NM",
            "인허가번호": "LCNS_NO",
        },
        "default_fields": ["소재지 주소", "업소명"],  # 기본으로 선택해 둘 조건 필드
        "has_clsbiz": True,  # 폐업일자 필드 보유 → '영업 중만' 필터 지원
    },
    "I2861": {
        "name": "인허가 변경 정보",
        "desc": "음식점업소 변경 이력",
        "columns": {
            "BSSH_NM": "업소명",
            "SITE_ADDR": "주소",
            "INDUTY_CD_NM": "업종",
            "CHNG_PRVNS": "변경사유",
            "CHNG_BF_CN": "변경 전",
            "CHNG_AF_CN": "변경 후",
            "CHNG_DT": "변경일",
            "LCNS_NO": "인허가번호",
        },
        "search_fields": {
            "업소명": "BSSH_NM",
            "주소": "SITE_ADDR",
            "변경 전 내용(대표자 등)": "CHNG_BF_CN",
            "변경 후 내용(대표자 등)": "CHNG_AF_CN",
            "인허가번호": "LCNS_NO",
        },
        "default_fields": ["주소", "업소명"],
        "has_clsbiz": False,
    },
}

HELP_TEXT = """\
■ 기본 사용법
  1. 왼쪽 사이드바에서 조회할 데이터셋을 선택합니다.
  2. 검색 조건을 입력하고 [조회] 버튼을 누릅니다. (조건 없이 조회하면 앞에서부터 순서대로 나옵니다)
  3. 결과 행을 더블클릭하면 모든 필드의 상세 정보를 볼 수 있습니다.

■ 검색 조건 조합 (최대 3개, AND 조건)
  여러 조건을 함께 쓰면 모두 만족하는 업소만 검색됩니다.
    예) 조건1: 소재지 주소 = 서울특별시 강남구
        조건2: 업소명 = 김밥
    → 강남구에 있으면서 이름에 '김밥'이 들어간 업소

■ 필드별 검색 팁
  · 업소명      : 이름의 일부만 입력해도 됩니다. 예) "김밥", "곱창", "스타벅"
  · 소재지 주소  : 넓은 지역부터 좁혀가세요.
                  예) "서울특별시" → "서울특별시 강남구" → "서울특별시 강남구 역삼동"
                  지번 주소와 도로명 주소가 섞여 있으므로, 동 이름까지만 쓰는 것이 잘 찾아집니다.
  · 대표자명    : 개인정보 보호로 가운데 글자가 * 처리되어 있습니다.
                  반드시 마스킹된 형태로 검색하세요. 예) "김*수", "이**" (원래 이름으로는 검색 안 됨)
  · 업종        : 예) "일반음식점", "휴게음식점", "단란주점", "유흥주점", "위탁급식영업", "제과점영업"
  · 인허가번호  : 번호 전체 또는 앞부분. 특정 업소를 정확히 짚을 때 가장 확실합니다.

■ 결과 표시 방식
  · 서버 검색 결과를 프로그램이 한 번 더 검증하여, 입력한 모든 조건이
    실제로 포함된 행만 표에 표시합니다. (상단 카드에서 수신/일치 건수 확인)
  · '영업 중인 업소만 표시'를 체크하면 폐업일자가 있는 업소를 제외합니다.

■ 검색 속도 관련
  · 여러 조건을 입력하면 서버에는 가장 구체적인 조건 1개(인허가번호 > 업소명 > 대표자명 > 업종 > 주소 순)만
    보내고, 나머지 조건은 프로그램이 수신 데이터에서 자동으로 걸러냅니다.
    (서버가 여러 조건을 정확히 처리하지 못하고, 넓은 조건은 시간 초과가 잦기 때문)
  · '서울'처럼 대상이 매우 많은 주소·지역 '단독' 검색은 서버가 느려 시간 초과가 날 수 있습니다.
    이때는 업소명 등 구체적인 조건을 함께 입력하세요.

■ 데이터 수신 방식
  · [조회]를 누르면 조건에 맞는 서버 데이터 '전체'를 1,000건씩 자동으로 이어서 수신합니다.
  · 수신 중 [중지]를 누르면 그때까지 수신한 결과만 표시합니다.
  · 일치 데이터가 매우 많으면(2만 건 초과) 전체를 계속 수신할지 물어봅니다.
  · 표에는 최대 5,000건까지 표시되며, 전체 결과는 CSV/JSON 저장으로 확인할 수 있습니다.
  · 조건 없이 조회하면 처음 1,000건만 가져옵니다.

■ 인증키
  · 식품안전나라(foodsafetykorea.go.kr) 회원가입 → 마이페이지 → 오픈API 인증키 발급
  · 'sample' 키는 테스트용: 5건만 반환되고 검색 조건이 적용되지 않습니다.
  · [키 저장] 체크 시 Windows 계정에 묶인 암호화(DPAPI)로 저장됩니다.
"""


# ── Windows DPAPI 암호화 (같은 PC·같은 사용자 계정에서만 복호화 가능) ──
class _DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ("cbData", ctypes.wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_char)),
    ]


def _blob_to_bytes(blob: _DATA_BLOB) -> bytes:
    data = ctypes.string_at(blob.pbData, blob.cbData)
    ctypes.windll.kernel32.LocalFree(blob.pbData)
    return data


def _bytes_to_blob(data: bytes) -> _DATA_BLOB:
    buf = ctypes.create_string_buffer(data, len(data))
    return _DATA_BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))


def dpapi_encrypt(text: str) -> str:
    blob_in = _bytes_to_blob(text.encode("utf-8"))
    blob_out = _DATA_BLOB()
    if not ctypes.windll.crypt32.CryptProtectData(
        ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)
    ):
        raise OSError("DPAPI 암호화에 실패했습니다.")
    return base64.b64encode(_blob_to_bytes(blob_out)).decode("ascii")


def dpapi_decrypt(b64: str) -> str:
    blob_in = _bytes_to_blob(base64.b64decode(b64))
    blob_out = _DATA_BLOB()
    if not ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)
    ):
        raise OSError("DPAPI 복호화에 실패했습니다. (다른 PC/계정에서 저장된 키일 수 있음)")
    return _blob_to_bytes(blob_out).decode("utf-8")


# ── 설정 파일 ─────────────────────────────────────
def load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            pass
    return {}


def save_config(cfg: dict):
    try:
        CONFIG_PATH.write_text(
            json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except OSError:
        pass


def load_saved_key() -> str:
    cfg = load_config()
    enc = cfg.get("api_key_enc")
    if enc:
        try:
            return dpapi_decrypt(enc)
        except OSError:
            return ""
    return cfg.get("api_key", "")  # 구버전 평문 저장분 호환


def save_key(key: str):
    cfg = load_config()
    cfg.pop("api_key", None)
    if key:
        cfg["api_key_enc"] = dpapi_encrypt(key)
    else:
        cfg.pop("api_key_enc", None)
    save_config(cfg)


# ── API 호출 ─────────────────────────────────────
def build_url(key: str, service: str, start: int, end: int, filters: dict) -> str:
    """조건이 여러 개면 &로 연결: /FIELD1=값1&FIELD2=값2"""
    url = f"{BASE_URL}/{key}/{service}/json/{start}/{end}"
    if filters:
        cond = "&".join(
            f"{f}={urllib.parse.quote(str(v))}" for f, v in filters.items()
        )
        url += f"/{cond}"
    return url


# 서버에 보낼 조건 선택 우선순위 (선별력 높은 필드 순).
# 서버는 다중 조건을 정확히 AND 처리하지 못하고, '서울'처럼 대상이 많은
# 광범위한 조건은 처리 시간이 매우 길어 시간 초과가 잦다.
FIELD_SELECTIVITY = [
    "LCNS_NO", "BSSH_NM", "PRSDNT_NM", "CHNG_BF_CN", "CHNG_AF_CN", "TELNO",
    "INDUTY_NM", "INDUTY_CD_NM", "INSTT_NM", "CHNG_PRVNS",
    "LOCP_ADDR", "SITE_ADDR", "SIGNGU_NM",
]


def pick_server_filters(filters: dict) -> dict:
    """서버에는 가장 선별력 높은 조건 1개만 보내고 나머지는 클라이언트에서 필터링"""
    if not filters:
        return {}
    for field in FIELD_SELECTIVITY:
        if field in filters:
            return {field: filters[field]}
    return dict(list(filters.items())[:1])


def api_fetch(url: str, service: str, timeout: int = 180, retries: int = 2):
    """API 호출. (total, rows) 반환, 오류 시 RuntimeError.
    시간 초과와 인증키 동시접속 거부는 자동 재시도."""
    data = None
    for attempt in range(retries + 1):
        try:
            with urllib.request.urlopen(url, timeout=timeout) as resp:
                text = resp.read().decode("utf-8", errors="replace")
        except urllib.error.HTTPError as e:
            raise RuntimeError(f"HTTP 오류: {e.code} {e.reason}")
        except (TimeoutError, urllib.error.URLError, OSError) as e:
            is_timeout = isinstance(e, TimeoutError) or "timed out" in str(e).lower()
            if is_timeout and attempt < retries:
                continue  # 재시도
            if is_timeout:
                raise RuntimeError(
                    "서버 응답 시간 초과입니다.\n\n"
                    "· '서울'처럼 대상이 매우 많은 주소·지역 단독 검색은 서버 처리가 느립니다.\n"
                    "· 업소명, 인허가번호 등 구체적인 조건을 함께 입력하면 빨라집니다.\n"
                    "· 잠시 후 다시 시도해 보세요."
                )
            raise RuntimeError(f"서버 연결 실패: {e}")

        try:
            data = json.loads(text)
            break
        except json.JSONDecodeError:
            # 서버가 JSON 대신 HTML 안내문(alert)을 보내는 경우
            m = re.search(r"alert\('([^']+)'\)", text)
            msg = m.group(1) if m else ""
            if "접속 중" in msg:
                # 같은 인증키의 이전 요청이 아직 처리 중 → 잠시 대기 후 재시도
                if attempt < retries:
                    time.sleep(3)
                    continue
                raise RuntimeError(
                    "같은 인증키로 이미 요청이 진행 중입니다.\n"
                    "이전 조회가 끝나기를 기다렸다가 다시 시도하세요.\n"
                    "(다른 창이나 프로그램에서 같은 키를 쓰고 있는지도 확인)"
                )
            if msg:
                raise RuntimeError(
                    f"서버 안내: {msg}\n"
                    f"(인증키가 '{service}' 데이터셋에 등록되지 않았을 수 있습니다. "
                    f"식품안전나라 마이페이지에서 활용 신청 데이터셋을 확인하세요.)"
                )
            raise RuntimeError("응답을 JSON으로 해석할 수 없습니다. 인증키를 확인하세요.")

    body = data.get(service)
    if body is None:
        result = data.get("RESULT", {})
        raise RuntimeError(
            f"{result.get('CODE', '오류')}: "
            f"{result.get('MSG', '응답 형식이 올바르지 않습니다.')}"
        )

    result = body.get("RESULT", {})
    code = result.get("CODE", "")
    if code == "INFO-200":  # 해당하는 데이터가 없음 → 빈 결과로 처리
        return 0, []
    if code != "INFO-000":
        raise RuntimeError(f"{code}: {result.get('MSG', '')}")

    return int(body.get("total_count", 0)), body.get("row", [])


def row_matches(row: dict, filters: dict) -> bool:
    """모든 검색 조건 값이 해당 필드에 실제로 포함되는지 검증 (대소문자 무시)"""
    for field, value in filters.items():
        if value.lower() not in str(row.get(field, "") or "").lower():
            return False
    return True


def format_date(value: str) -> str:
    if value and len(value) == 8 and value.isdigit():
        return f"{value[:4]}-{value[4:6]}-{value[6:]}"
    return value or ""


# ── GUI ─────────────────────────────────────────
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.rows = []          # 표에 표시 중인(조건 일치) 결과
        self.rows_received = 0  # 서버에서 수신한 건수
        self._timer_id = None
        self._elapsed = 0
        self._cond_desc = ""
        self._page_info = ""
        self._cancel = False
        self._open_cnt = 0
        self._open_only_pref = True  # '영업 중만 표시' 기본값
        self._current_service = "I1200"

        root.title("음식점 정보 조회 대시보드 — 식품안전나라 오픈 API")
        root.geometry("1180x740")
        root.minsize(980, 600)
        root.configure(bg=C["bg"])
        try:
            ico = resource_path("app.ico")
            if ico.exists():
                root.iconbitmap(default=str(ico))
        except Exception:
            pass  # 아이콘 적용 실패해도 실행에는 지장 없음

        self._init_styles()
        self._build_sidebar()
        self._build_main()
        self.select_service("I1200", initial=True)

    # ── 스타일 ──────────────────────────────
    def _init_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(".", font=(FONT, 10), background=C["card"])
        style.configure("TCombobox", fieldbackground="white")
        style.configure("Card.TFrame", background=C["card"])
        style.configure(
            "Treeview",
            font=(FONT, 10),
            rowheight=28,
            background="white",
            fieldbackground="white",
            foreground=C["text"],
            bordercolor=C["border"],
        )
        style.configure(
            "Treeview.Heading",
            font=(FONT, 10, "bold"),
            background="#EEF0FA",
            foreground=C["text"],
            relief="flat",
            padding=6,
        )
        style.map("Treeview.Heading", background=[("active", "#E3E6F5")])
        style.map(
            "Treeview",
            background=[("selected", C["accent"])],
            foreground=[("selected", "white")],
        )
        style.configure(
            "Accent.Horizontal.TProgressbar",
            troughcolor=C["border"],
            background=C["accent"],
        )

    # ── 사이드바 ─────────────────────────────
    def _build_sidebar(self):
        sb = tk.Frame(self.root, bg=C["sidebar"], width=230)
        sb.pack(side="left", fill="y")
        sb.pack_propagate(False)

        tk.Label(
            sb, text="🍽  FoodSafe", bg=C["sidebar"], fg="white",
            font=(FONT, 16, "bold"), anchor="w",
        ).pack(fill="x", padx=20, pady=(24, 0))
        tk.Label(
            sb, text="Restaurant Finder", bg=C["sidebar"], fg=C["sidebar_dim"],
            font=(FONT, 10), anchor="w",
        ).pack(fill="x", padx=22, pady=(0, 24))

        tk.Label(
            sb, text="데이터셋", bg=C["sidebar"], fg=C["sidebar_dim"],
            font=(FONT, 9, "bold"), anchor="w",
        ).pack(fill="x", padx=20, pady=(0, 4))

        self.side_btns = {}
        for sid, info in SERVICES.items():
            btn = tk.Frame(sb, bg=C["sidebar"], cursor="hand2")
            btn.pack(fill="x", padx=10, pady=2)
            lbl1 = tk.Label(
                btn, text=info["name"], bg=C["sidebar"], fg=C["sidebar_tx"],
                font=(FONT, 11, "bold"), anchor="w", padx=12, pady=2,
            )
            lbl1.pack(fill="x", pady=(6, 0))
            lbl2 = tk.Label(
                btn, text=info["desc"], bg=C["sidebar"], fg=C["sidebar_dim"],
                font=(FONT, 8), anchor="w", padx=12,
            )
            lbl2.pack(fill="x", pady=(0, 6))
            for w in (btn, lbl1, lbl2):
                w.bind("<Button-1>", lambda e, s=sid: self.select_service(s))
            self.side_btns[sid] = (btn, lbl1, lbl2)

        # 하단 도움말/정보
        bottom = tk.Frame(sb, bg=C["sidebar"])
        bottom.pack(side="bottom", fill="x", pady=16)
        help_btn = tk.Label(
            bottom, text="❓  검색 방법 안내", bg=C["sidebar_hi"], fg="white",
            font=(FONT, 10, "bold"), pady=10, cursor="hand2",
        )
        help_btn.pack(fill="x", padx=14, pady=(0, 10))
        help_btn.bind("<Button-1>", lambda e: self.show_help())
        tk.Label(
            bottom, text="식품안전나라 오픈 API\nfoodsafetykorea.go.kr",
            bg=C["sidebar"], fg=C["sidebar_dim"], font=(FONT, 8), justify="left",
        ).pack(fill="x", padx=20)

    def _highlight_sidebar(self, active_sid: str):
        for sid, (btn, lbl1, lbl2) in self.side_btns.items():
            bg = C["sidebar_hi"] if sid == active_sid else C["sidebar"]
            fg1 = "white" if sid == active_sid else C["sidebar_tx"]
            btn.configure(bg=bg)
            lbl1.configure(bg=bg, fg=fg1)
            lbl2.configure(bg=bg)

    # ── 본문 ────────────────────────────────
    def _card(self, parent, **pack_opts):
        """흰색 카드 프레임 생성"""
        outer = tk.Frame(parent, bg=C["border"])
        outer.pack(**pack_opts)
        inner = tk.Frame(outer, bg=C["card"])
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        return inner

    def _build_main(self):
        main = tk.Frame(self.root, bg=C["bg"])
        main.pack(side="left", fill="both", expand=True)

        # 헤더
        header = tk.Frame(main, bg=C["bg"])
        header.pack(fill="x", padx=18, pady=(16, 8))
        self.var_title = tk.StringVar()
        tk.Label(
            header, textvariable=self.var_title, bg=C["bg"], fg=C["text"],
            font=(FONT, 15, "bold"), anchor="w",
        ).pack(side="left")

        # 인증키 (헤더 오른쪽)
        keybox = tk.Frame(header, bg=C["bg"])
        keybox.pack(side="right")
        tk.Label(keybox, text="인증키", bg=C["bg"], fg=C["muted"], font=(FONT, 9)).pack(side="left", padx=(0, 6))
        saved_key = load_saved_key()
        self.var_key = tk.StringVar(value=saved_key or "sample")
        self.ent_key = ttk.Entry(keybox, textvariable=self.var_key, width=30, show="*")
        self.ent_key.pack(side="left")
        self.var_show = tk.BooleanVar(value=False)
        tk.Checkbutton(
            keybox, text="표시", variable=self.var_show, bg=C["bg"], font=(FONT, 9),
            command=lambda: self.ent_key.config(show="" if self.var_show.get() else "*"),
        ).pack(side="left", padx=2)
        self.var_save_key = tk.BooleanVar(value=bool(saved_key))
        tk.Checkbutton(
            keybox, text="저장(암호화)", variable=self.var_save_key, bg=C["bg"], font=(FONT, 9),
        ).pack(side="left")

        # 검색 조건 카드
        cond_card = self._card(main, fill="x", padx=18, pady=6)
        cond = tk.Frame(cond_card, bg=C["card"])
        cond.pack(fill="x", padx=14, pady=12)

        tk.Label(
            cond, text="검색 조건", bg=C["card"], fg=C["text"], font=(FONT, 11, "bold"),
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))
        tk.Label(
            cond,
            text="조건을 여러 개 입력하면 모두 만족(AND)하는 업소만 검색됩니다.   예) 소재지 주소=서울특별시 강남구 + 업소명=김밥",
            bg=C["card"], fg=C["muted"], font=(FONT, 9),
        ).grid(row=0, column=1, columnspan=5, sticky="w", padx=8, pady=(0, 6))

        self.cond_rows = []
        for i in range(MAX_CONDITIONS):
            tk.Label(cond, text=f"조건 {i + 1}", bg=C["card"], fg=C["muted"], font=(FONT, 9)).grid(
                row=1 + i, column=0, sticky="e", padx=(0, 8), pady=3
            )
            var_field = tk.StringVar()
            cbo = ttk.Combobox(cond, textvariable=var_field, state="readonly", width=22, font=(FONT, 10))
            cbo.grid(row=1 + i, column=1, sticky="w", pady=3)
            var_value = tk.StringVar()
            ent = ttk.Entry(cond, textvariable=var_value, width=36, font=(FONT, 10))
            ent.grid(row=1 + i, column=2, sticky="w", padx=8, pady=3)
            ent.bind("<Return>", lambda e: self.search())
            self.cond_rows.append((var_field, cbo, var_value))

        # 조건 3개 행의 높이를 동일하게 고정 (오른쪽 패널 높이에 따라 간격이 달라지지 않도록)
        for r in (1, 2, 3):
            cond.rowconfigure(r, uniform="condrow")

        # 오른쪽 패널: 안내·옵션(위) + 버튼(아래), 조건 3행 전체에 걸침
        side = tk.Frame(cond, bg=C["card"])
        side.grid(row=1, column=3, rowspan=3, sticky="nsew", padx=(20, 0))
        tk.Label(
            side,
            text="조건에 맞는 서버 데이터 전체를\n1,000건씩 자동으로 이어서 수신합니다.",
            bg=C["card"], fg=C["muted"], font=(FONT, 8), justify="left", anchor="w",
        ).pack(anchor="w", pady=(0, 4))

        self.var_open_only = tk.BooleanVar(value=True)  # 기본: 영업 중인 업소만
        self.chk_open = tk.Checkbutton(
            side, text="영업 중인 업소만 표시 (폐업 제외)", variable=self.var_open_only,
            bg=C["card"], font=(FONT, 9), anchor="w",
        )
        self.chk_open.pack(anchor="w")

        # 조회/중지/지우기 버튼
        btns = tk.Frame(side, bg=C["card"])
        btns.pack(side="bottom", anchor="e", pady=(6, 0))
        self.btn_search = tk.Button(
            btns, text="조회", command=self.search, bg=C["accent"], fg="white",
            font=(FONT, 11, "bold"), relief="flat", padx=28, pady=6, cursor="hand2",
            activebackground=C["accent_dark"], activeforeground="white",
        )
        self.btn_search.pack(side="right", padx=(8, 0))
        self.btn_stop = tk.Button(
            btns, text="중지", command=self.cancel_search, bg=C["red"], fg="white",
            font=(FONT, 10, "bold"), relief="flat", padx=14, pady=6, cursor="hand2",
            state="disabled", disabledforeground="#FCA5A5",
        )
        self.btn_stop.pack(side="right", padx=(8, 0))
        tk.Button(
            btns, text="조건 지우기", command=self.clear_conditions, bg="#E5E7EB",
            fg=C["text"], font=(FONT, 10), relief="flat", padx=12, pady=6, cursor="hand2",
        ).pack(side="right")

        # KPI 카드 행
        kpi_row = tk.Frame(main, bg=C["bg"])
        kpi_row.pack(fill="x", padx=18, pady=6)
        self.kpi_vars = {}
        kpi_defs = [
            ("total", "서버 일치(전체)", C["accent"]),
            ("received", "수신", C["amber"]),
            ("matched", "조건 일치(표시)", C["green"]),
            ("open", "영업 중", C["text"]),
        ]
        for key, label, color in kpi_defs:
            card = self._card(kpi_row, side="left", fill="x", expand=True, padx=(0, 10) if key != "open" else 0)
            var = tk.StringVar(value="-")
            tk.Label(card, text=label, bg=C["card"], fg=C["muted"], font=(FONT, 9), anchor="w").pack(
                fill="x", padx=14, pady=(10, 0)
            )
            tk.Label(card, textvariable=var, bg=C["card"], fg=color, font=(FONT, 17, "bold"), anchor="w").pack(
                fill="x", padx=14, pady=(0, 10)
            )
            self.kpi_vars[key] = var

        # 결과 테이블 카드
        table_card = self._card(main, fill="both", expand=True, padx=18, pady=6)
        tf = tk.Frame(table_card, bg=C["card"])
        tf.pack(fill="both", expand=True, padx=10, pady=10)

        self.tree = ttk.Treeview(tf, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tf, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tf.rowconfigure(0, weight=1)
        tf.columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", self.show_detail)
        self.tree.tag_configure("odd", background="#F8F9FE")
        self.tree.tag_configure("closed", foreground="#9CA3AF")

        # 하단 상태줄
        bottom = tk.Frame(main, bg=C["bg"])
        bottom.pack(fill="x", padx=18, pady=(4, 14))
        self.var_status = tk.StringVar(value="준비됨. 인증키를 확인하고 [조회]를 누르세요.")
        tk.Label(bottom, textvariable=self.var_status, bg=C["bg"], fg=C["muted"], font=(FONT, 9)).pack(side="left")
        self.progress = ttk.Progressbar(
            bottom, mode="indeterminate", length=180, style="Accent.Horizontal.TProgressbar"
        )

        tk.Button(
            bottom, text="CSV로 저장", command=self.save_csv, bg="#E5E7EB", fg=C["text"],
            font=(FONT, 9), relief="flat", padx=12, pady=4, cursor="hand2",
        ).pack(side="right", padx=(6, 0))
        tk.Button(
            bottom, text="JSON으로 저장", command=self.save_json, bg="#E5E7EB", fg=C["text"],
            font=(FONT, 9), relief="flat", padx=12, pady=4, cursor="hand2",
        ).pack(side="right")

    # ── 동작 ────────────────────────────────
    @property
    def service_id(self) -> str:
        return self._current_service

    def select_service(self, sid: str, initial: bool = False):
        self._current_service = sid
        info = SERVICES[sid]
        self._highlight_sidebar(sid)
        self.var_title.set(f"{info['name']}  —  {info['desc']}")

        field_names = ["(사용 안 함)"] + list(info["search_fields"].keys())
        defaults = info.get("default_fields", [])
        for i, (var_field, cbo, var_value) in enumerate(self.cond_rows):
            cbo["values"] = field_names
            # 기본 검색 필드를 미리 선택해 바로 입력할 수 있게 함
            if i < len(defaults):
                var_field.set(defaults[i])
            else:
                cbo.current(0)
            var_value.set("")

        if info["has_clsbiz"]:
            self.chk_open.config(state="normal")
            self.var_open_only.set(self._open_only_pref)
        else:
            # 폐업 필드가 없는 데이터셋으로 전환: 직전 선택을 기억해 두고 비활성화
            if str(self.chk_open.cget("state")) == "normal":
                self._open_only_pref = self.var_open_only.get()
            self.chk_open.config(state="disabled")
            self.var_open_only.set(False)

        self.setup_columns(info["columns"])
        self.rows = []
        self.rows_received = 0
        for var in self.kpi_vars.values():
            var.set("-")
        if not initial:
            self.var_status.set(f"'{info['name']}' 데이터셋 선택됨. 조건을 입력하고 [조회]를 누르세요.")

    def clear_conditions(self):
        info = SERVICES[self.service_id]
        defaults = info.get("default_fields", [])
        for i, (var_field, cbo, var_value) in enumerate(self.cond_rows):
            if i < len(defaults):
                var_field.set(defaults[i])
            else:
                cbo.current(0)
            var_value.set("")

    def collect_filters(self) -> dict:
        info = SERVICES[self.service_id]
        filters = {}
        for var_field, _cbo, var_value in self.cond_rows:
            label = var_field.get()
            value = var_value.get().strip()
            if label and label != "(사용 안 함)" and value:
                filters[info["search_fields"][label]] = value
        return filters

    def setup_columns(self, columns: dict):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(columns.keys())
        for field, label in columns.items():
            width = 95 if field.endswith(("_DT", "YMD")) or field == "YEAR" else 130
            if field in ("SITE_ADDR", "LOCP_ADDR", "BSSH_NM"):
                width = 240
            if field == "PRSDNT_NM":
                width = 85
            self.tree.heading(field, text=label)
            self.tree.column(field, width=width, anchor="w")

    def search(self):
        key = self.var_key.get().strip()
        if not key:
            messagebox.showwarning("확인", "인증키를 입력하세요.")
            return

        service = self.service_id
        filters = self.collect_filters()
        self._last_filters = filters

        if not filters:
            if not messagebox.askyesno(
                "확인",
                "검색 조건이 없습니다.\n"
                "조건 없이 조회하면 처음 1,000건만 표시합니다.\n계속할까요?",
            ):
                return

        # 키 저장 처리 (DPAPI 암호화)
        try:
            if self.var_save_key.get() and key != "sample":
                save_key(key)
            else:
                save_key("")
        except OSError as e:
            messagebox.showwarning("키 저장", str(e))

        info = SERVICES[service]
        # 결과 초기화
        self.setup_columns(info["columns"])
        self.rows = []
        self.rows_received = 0
        self._open_cnt = 0
        self._cancel = False
        for var in self.kpi_vars.values():
            var.set("-")

        label_of = {v: k for k, v in info["search_fields"].items()}
        self._cond_desc = ", ".join(f"{label_of.get(f, f)}={v}" for f, v in filters.items()) or "조건 없음"
        server_filters = pick_server_filters(filters)
        if len(filters) > 1:
            sf = next(iter(server_filters))
            self._cond_desc += f" · 서버검색: {label_of.get(sf, sf)}, 나머지는 자동 필터링"
        self._page_info = "1페이지 요청 중"
        self.btn_search.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.start_progress()

        open_only = self.var_open_only.get() and info["has_clsbiz"]
        threading.Thread(
            target=self._worker, args=(key, service, filters, open_only), daemon=True
        ).start()

    def _worker(self, key: str, service: str, filters: dict, open_only: bool):
        """조건에 맞는 서버 데이터 전체를 1,000건씩 이어서 수신.
        서버에는 가장 선별력 높은 조건 1개만 보내고, 전체 조건 검증은 _on_page에서 수행."""
        server_filters = pick_server_filters(filters)
        start = 1
        total = None
        pages = 1
        page_no = 0
        try:
            while not self._cancel:
                url = build_url(key, service, start, start + PAGE_SIZE - 1, server_filters)
                t, rows = api_fetch(url, service)
                if total is None:
                    total = t
                    pages = max(1, (total + PAGE_SIZE - 1) // PAGE_SIZE)
                page_no += 1
                self.root.after(0, self._on_page, total, rows, page_no, pages, open_only)

                if not filters:  # 조건 없으면 첫 페이지만
                    break
                start += PAGE_SIZE
                if start > total or not rows:
                    break
                # 대량 데이터는 전체 수신 여부를 한 번 확인
                if page_no == 1 and pages > CONFIRM_PAGES:
                    if not self._ask_main(
                        f"조건과 일치하는 서버 데이터가 {total:,}건({pages}페이지)입니다.\n"
                        f"모두 수신하면 시간이 오래 걸릴 수 있습니다.\n\n"
                        f"전체를 수신할까요?\n(아니오: 지금까지 수신분만 표시)"
                    ):
                        break
            self.root.after(0, self._on_done, total or 0, pages)
        except RuntimeError as e:
            self.root.after(0, self.on_error, str(e))

    def _ask_main(self, msg: str) -> bool:
        """작업 스레드에서 메인 스레드의 확인 대화상자를 띄우고 답을 기다림"""
        ev = threading.Event()
        res = {"ok": False}

        def ask():
            res["ok"] = messagebox.askyesno("전체 수신 확인", msg)
            ev.set()

        self.root.after(0, ask)
        ev.wait()
        return res["ok"]

    def cancel_search(self):
        self._cancel = True
        self.btn_stop.config(state="disabled")
        self.var_status.set("중지 요청됨. 수신 중인 페이지까지 반영하고 멈춥니다...")

    def _on_page(self, total: int, rows: list, page_no: int, pages: int, open_only: bool):
        """수신한 한 페이지를 조건으로 재검증하여 표에 추가"""
        info = SERVICES[self.service_id]
        filters = getattr(self, "_last_filters", {})

        matched = [r for r in rows if row_matches(r, filters)]
        if open_only:
            matched = [r for r in matched if not (r.get("CLSBIZ_DT") or "").strip()]

        self.rows_received += len(rows)
        base = len(self.rows)
        for i, row in enumerate(matched):
            pos = base + i
            if pos >= MAX_TREE_ROWS:
                break  # 표는 상한까지만 (저장은 전체)
            values = []
            for field in info["columns"]:
                v = row.get(field, "") or ""
                if field.endswith("_DT") or field.endswith("YMD"):
                    v = format_date(v)
                values.append(v)
            tags = ["odd"] if pos % 2 == 1 else []
            if info["has_clsbiz"] and (row.get("CLSBIZ_DT") or "").strip():
                tags.append("closed")  # 폐업 업소는 회색으로 표시
            self.tree.insert("", "end", values=values, tags=tags)

        if info["has_clsbiz"]:
            self._open_cnt += sum(
                1 for r in matched if not (r.get("CLSBIZ_DT") or "").strip()
            )
        self.rows.extend(matched)

        self.kpi_vars["total"].set(f"{total:,}")
        self.kpi_vars["received"].set(f"{self.rows_received:,}")
        self.kpi_vars["matched"].set(f"{len(self.rows):,}")
        self.kpi_vars["open"].set(
            f"{self._open_cnt:,}" if info["has_clsbiz"] else "해당 없음"
        )
        self._page_info = f"페이지 {page_no}/{pages} 수신, 일치 {len(self.rows):,}건"

    def _on_done(self, total: int, pages: int):
        self.stop_progress()
        self.btn_search.config(state="normal")
        self.btn_stop.config(state="disabled")

        n = len(self.rows)
        prefix = "중지됨" if self._cancel else "완료"
        msg = f"{prefix}: 서버 일치 {total:,}건 중 {self.rows_received:,}건 수신 → 조건 일치 {n:,}건"
        if n > MAX_TREE_ROWS:
            msg += f" (표에는 처음 {MAX_TREE_ROWS:,}건만 표시, 전체는 CSV/JSON 저장으로 확인)"
        self.var_status.set(msg)

        if n == 0 and not self._cancel:
            messagebox.showinfo(
                "결과 없음",
                "조건에 맞는 데이터가 없습니다.\n\n"
                "· 검색어를 더 짧게(일부만) 입력해 보세요.\n"
                "· 주소는 '서울특별시 강남구'처럼 행정구역 단위로 입력해 보세요.\n"
                "· 대표자명은 김*수 처럼 마스킹된 형태여야 합니다.\n"
                "· [검색 방법 안내]에서 자세한 팁을 확인하세요.",
            )

    def start_progress(self):
        self._elapsed = 0
        self.progress.pack(side="left", padx=(12, 4))
        self.progress.start(12)
        self.root.config(cursor="watch")
        self._update_wait_status()

    def _update_wait_status(self):
        self.var_status.set(
            f"조회 중... {self._elapsed}초 경과 · {self._page_info} ({self._cond_desc})"
        )
        self._elapsed += 1
        self._timer_id = self.root.after(1000, self._update_wait_status)

    def stop_progress(self):
        if self._timer_id is not None:
            self.root.after_cancel(self._timer_id)
            self._timer_id = None
        self.progress.stop()
        self.progress.pack_forget()
        self.root.config(cursor="")

    def on_error(self, msg: str):
        self.stop_progress()
        self.btn_search.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.var_status.set("오류 발생")
        messagebox.showerror("오류", msg)

    def show_detail(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        if idx >= len(self.rows):
            return
        row = self.rows[idx]
        info = SERVICES[self.service_id]

        win = tk.Toplevel(self.root)
        win.title(row.get("BSSH_NM", "상세 정보"))
        win.geometry("500x440")
        win.configure(bg=C["card"])
        txt = tk.Text(win, wrap="word", font=(FONT, 10), relief="flat", bg=C["card"])
        txt.pack(fill="both", expand=True, padx=12, pady=12)
        for field, value in row.items():
            label = info["columns"].get(field, field)
            txt.insert("end", f"{label} ({field})\n    {value or '-'}\n\n")
        txt.config(state="disabled")

    def show_help(self):
        win = tk.Toplevel(self.root)
        win.title("검색 방법 안내")
        win.geometry("680x620")
        win.configure(bg=C["card"])
        frame = tk.Frame(win, bg=C["card"])
        frame.pack(fill="both", expand=True, padx=8, pady=8)
        txt = tk.Text(frame, wrap="word", font=(FONT, 10), relief="flat", bg=C["card"], spacing1=2, spacing3=2)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=vsb.set)
        txt.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        txt.insert("1.0", HELP_TEXT)
        txt.config(state="disabled")

    def save_json(self):
        if not self._check_rows():
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json", filetypes=[("JSON 파일", "*.json")],
            initialfile="음식점조회결과.json",
        )
        if not path:
            return
        Path(path).write_text(
            json.dumps(self.rows, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        self.var_status.set(f"저장 완료: {path}")

    def save_csv(self):
        if not self._check_rows():
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV 파일", "*.csv")],
            initialfile="음식점조회결과.csv",
        )
        if not path:
            return
        import csv

        info = SERVICES[self.service_id]
        # 엑셀에서 한글이 깨지지 않도록 BOM 포함 UTF-8 사용
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(info["columns"].values())
            for row in self.rows:
                writer.writerow([row.get(field, "") for field in info["columns"]])
        self.var_status.set(f"저장 완료: {path}")

    def _check_rows(self) -> bool:
        if not self.rows:
            messagebox.showinfo("안내", "저장할 조회 결과가 없습니다. 먼저 조회하세요.")
            return False
        return True


def main():
    root = tk.Tk()
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)  # 고해상도 화면에서 흐릿함 방지
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
