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
import hashlib
import json
import os
import re
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ── 버전 / 자동 업데이트 ──────────────────────────
# 프로그램 버전의 단일 출처. release.ps1이 릴리스 시 이 줄을 정규식으로 교체한다.
APP_VERSION = "1.0.0"
UPDATE_REPO = "everside76/RestaurantFinder"   # GitHub 저장소 (owner/repo)
EXE_ASSET_NAME = "RestaurantFinder.exe"       # 릴리스 에셋 중 내려받을 exe 이름
UPDATE_API_URL = f"https://api.github.com/repos/{UPDATE_REPO}/releases/latest"
UPDATE_USER_AGENT = f"RestaurantFinder-Updater/{APP_VERSION}"

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
    "MOIS_I1": {
        "name": "일반음식점 (행정안전부)",
        "desc": "공공데이터포털 · 별도 인증키",
        "provider": "data_go_kr",  # 행정안전부 일반음식점 조회서비스 (data.go.kr)
        "columns": {
            "BPLC_NM": "사업장명",
            "ROAD_NM_ADDR": "도로명주소",
            "LOTNO_ADDR": "지번주소",
            "SALS_STTS_NM": "영업상태",
            "DTL_SALS_STTS_NM": "상세영업상태",
            "BZSTAT_SE_NM": "업태",
            "TELNO": "전화번호",
            "LCPMT_YMD": "인허가일자",
            "CLSBIZ_YMD": "폐업일자",
            "MNG_NO": "관리번호",
        },
        # 검색 필드의 값은 응답 필드명(cond 파라미터에 그대로 쓰는 이름)으로 정의
        "search_fields": {
            "사업장명": "BPLC_NM",
            "도로명주소": "ROAD_NM_ADDR",
        },
        "default_fields": ["사업장명", "도로명주소"],
        "has_clsbiz": False,  # 폐업 판정은 영업상태코드로 처리(is_row_closed 참조)
        "default_sort": "LCPMT_YMD",  # 조회 시 인허가일자 최근순으로 기본 정렬
    },
    "I1200": {
        "name": "식품접객업정보",
        "desc": "전국 음식점 인허가 (약 313만 건)",
        "provider": "mfds",  # 식품안전나라(foodsafetykorea.go.kr) 오픈 API
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
        "default_fields": ["업소명", "소재지 주소"],  # 기본으로 선택해 둘 조건 필드
        "has_clsbiz": True,  # 폐업일자 필드 보유 → '영업 중만' 필터 지원
        "default_sort": "PRMS_DT",  # 조회 시 인허가일자 최근순으로 기본 정렬
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

■ 결과 표 다루기
  · 열 제목을 클릭하면 그 필드 기준으로 정렬됩니다. 다시 클릭하면 오름차순/내림차순이
    바뀝니다. (제목 옆 ▲/▼가 현재 정렬 방향)
  · 행을 더블클릭하면 모든 필드의 상세 정보를 볼 수 있습니다.
  · 셀에서 마우스 오른쪽 버튼을 누르면 '셀 값 / 행 전체 / 열 전체'를 복사할 수 있습니다.
    (Ctrl+C 는 선택한 행 복사)

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

■ 'ERROR-500 서버오류' 안내
  · 식품안전나라(공공데이터) 서버 쪽 오류이며 이 프로그램의 문제가 아닙니다.
  · 프로그램이 자동으로 몇 차례 재시도하며, 그래도 안 되면 잠시 후 다시 시도하세요.
  · 하루 호출량이 많으면 일시 제한될 수 있으니, 계속되면 시간을 두고(다음 날) 시도하세요.

■ '일반음식점 (행정안전부)' 데이터셋 — 공공데이터포털(data.go.kr) · 기본 데이터셋
  · 사이드바 맨 위, 프로그램 시작 시 기본으로 선택되는 데이터셋입니다.
    식품안전나라와는 별개의 데이터·별개의 인증키를 씁니다.
  · 검색 조건: 사업장명 / 도로명주소 (모두 서버에서 AND로 정확히 처리됩니다.
    식품안전나라와 달리 여러 조건을 서버가 직접 걸러 주므로 자동 필터링 안내가 없습니다.)
  · 영업상태(영업/폐업)를 함께 제공하며, '영업 중인 업소만 표시'를 체크하면
    영업/정상(코드 01)만 서버에서 받아 오고 폐업 행은 회색으로 구분됩니다.
  · 페이지당 100건씩 나눠 받으므로 결과가 많으면 수신이 느릴 수 있습니다.
    (2만 건을 넘으면 전체 수신 여부를 한 번 물어봅니다.)

■ 인증키 (두 개는 서로 완전히 별개입니다)
  · 데이터셋을 바꾸면 헤더의 인증키 입력란 라벨과 값이 자동으로 그 제공처의 키로 전환됩니다.
  · 식품안전나라(foodsafetykorea.go.kr): 회원가입 → 마이페이지 → 오픈API 인증키 발급
    - 'sample' 키는 테스트용으로 5건만 반환되고 검색 조건이 적용되지 않습니다.
  · 공공데이터포털(data.go.kr): 회원가입 → 해당 API('행정안전부_식품_일반음식점 조회서비스')
    활용신청 → 승인 후 마이페이지 → '일반 인증키(Decoding)'를 복사해 입력하세요.
    - 반드시 Decoding 키를 입력해야 합니다(Encoding 키를 넣으면 인증 오류가 납니다).
    - 활용신청 승인 직후에는 반영에 시간이 걸릴 수 있습니다(수십 분~1시간).
  · [저장(암호화)] 체크 시 각 키가 Windows 계정에 묶인 암호화(DPAPI)로 따로 저장됩니다.

■ 자동 업데이트
  · 프로그램을 시작하고 몇 초 뒤 GitHub 릴리스에서 새 버전이 있는지 자동으로 확인합니다.
    (인터넷 연결이 없거나 확인에 실패해도 조용히 넘어가며, 조회 기능에는 영향이 없습니다.)
  · 새 버전이 있으면 백그라운드에서 자동으로 내려받고, 아래 상태줄에 진행률(%)이 표시됩니다.
    내려받은 파일은 크기와 SHA-256 해시로 검증한 뒤에만 사용합니다.
  · 다운로드가 끝나면 적용 방법을 묻습니다.
    - [예]   : 즉시 재시작하여 새 버전을 적용합니다.
    - [아니오]: 지금은 계속 사용하고, 프로그램을 닫을 때 자동으로 적용합니다.
    - 조회가 진행 중일 때는 결과를 잃지 않도록 재시작하지 않고 '닫을 때 적용'으로 전환합니다.
  · 사이드바 아래의 [🔄 업데이트 확인] 버튼으로 언제든 직접 확인할 수 있습니다.
    (현재 버전은 그 아래 v 숫자로 표시됩니다.)
  · 업데이트는 실행 파일(RestaurantFinder.exe)만 교체합니다. exe 옆 config.json에
    저장된 인증키는 그대로 유지되므로 다시 입력할 필요가 없습니다.
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


# 데이터 제공처(provider)별 인증키 저장 필드명.
# mfds=식품안전나라, data_go_kr=공공데이터포털. 두 키는 서로 완전히 별개다.
PROVIDER_KEY_FIELD = {
    "mfds": "api_key_enc",
    "data_go_kr": "datagokr_key_enc",
}


def load_saved_key(provider: str = "mfds") -> str:
    cfg = load_config()
    field = PROVIDER_KEY_FIELD.get(provider, "api_key_enc")
    enc = cfg.get(field)
    if enc:
        try:
            return dpapi_decrypt(enc)
        except OSError:
            return ""
    if provider == "mfds":
        return cfg.get("api_key", "")  # 구버전 평문 저장분 호환
    return ""


def save_key(key: str, provider: str = "mfds"):
    cfg = load_config()
    field = PROVIDER_KEY_FIELD.get(provider, "api_key_enc")
    if provider == "mfds":
        cfg.pop("api_key", None)  # 구버전 평문 필드 제거
    if key:
        cfg[field] = dpapi_encrypt(key)
    else:
        cfg.pop(field, None)
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
    "LCNS_NO", "BSSH_NM", "PRSDNT_NM", "TELNO",
    "INDUTY_NM", "INSTT_NM",
    "LOCP_ADDR", "SIGNGU_NM",
]


def pick_server_filters(filters: dict) -> dict:
    """서버에는 가장 선별력 높은 조건 1개만 보내고 나머지는 클라이언트에서 필터링"""
    if not filters:
        return {}
    for field in FIELD_SELECTIVITY:
        if field in filters:
            return {field: filters[field]}
    return dict(list(filters.items())[:1])


# 서버 측 일시 오류 코드(재시도 대상). 500=서버오류, 600=DB연결실패, 601=SQL오류
SERVER_ERROR_CODES = {"ERROR-500", "ERROR-600", "ERROR-601"}


def _server_error_message(code: str, msg: str) -> str:
    return (
        f"식품안전나라 서버가 응답하지 않습니다 ({code}: {msg or '서버 오류'}).\n\n"
        "· 이 프로그램이 아니라 공공데이터 API 서버 쪽 문제입니다.\n"
        "· 잠시 후(수 분~수십 분) 다시 시도해 주세요.\n"
        "· 계속 반복되면 오늘 호출량이 많아 일시적으로 제한됐을 수 있으니,\n"
        "  시간을 두고(가능하면 다음 날) 다시 시도하세요."
    )


def api_fetch(url: str, service: str, timeout: int = 180, retries: int = 3):
    """API 호출. (total, rows) 반환, 오류 시 RuntimeError.
    시간 초과·동시접속 거부·서버 오류(ERROR-500 등)는 자동 재시도."""
    for attempt in range(retries + 1):
        last = attempt == retries
        try:
            with urllib.request.urlopen(url, timeout=timeout) as resp:
                text = resp.read().decode("utf-8", errors="replace")
        except urllib.error.HTTPError as e:
            raise RuntimeError(f"HTTP 오류: {e.code} {e.reason}")
        except (TimeoutError, urllib.error.URLError, OSError) as e:
            is_timeout = isinstance(e, TimeoutError) or "timed out" in str(e).lower()
            if is_timeout and not last:
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
        except json.JSONDecodeError:
            # 서버가 JSON 대신 HTML 안내문(alert)을 보내는 경우
            m = re.search(r"alert\('([^']+)'\)", text)
            msg = m.group(1) if m else ""
            if "접속 중" in msg:
                # 같은 인증키의 이전 요청이 아직 처리 중 → 잠시 대기 후 재시도
                if not last:
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

        # RESULT 코드는 body 안(정상) 또는 최상위(오류)에 올 수 있음
        body = data.get(service)
        result = (body or {}).get("RESULT") or data.get("RESULT", {})
        code = result.get("CODE", "")

        # 서버 측 일시 오류 → 점증 대기 후 재시도
        if code in SERVER_ERROR_CODES:
            if not last:
                time.sleep(2 * (attempt + 1))
                continue
            raise RuntimeError(_server_error_message(code, result.get("MSG", "")))

        if code == "INFO-200":  # 해당하는 데이터가 없음 → 빈 결과로 처리
            return 0, []
        if body is None:
            raise RuntimeError(
                f"{code or '오류'}: {result.get('MSG', '응답 형식이 올바르지 않습니다.')}"
            )
        if code and code != "INFO-000":
            raise RuntimeError(f"{code}: {result.get('MSG', '')}")

        return int(body.get("total_count", 0)), body.get("row", [])

    # 이론상 도달하지 않음 (마지막 시도에서 반환/예외)
    raise RuntimeError("알 수 없는 오류로 조회에 실패했습니다.")


# ── 공공데이터포털(data.go.kr) 일반음식점 API ──────────────
DATAGOKR_URL = "https://apis.data.go.kr/1741000/general_restaurants/info"
DATAGOKR_PAGE_SIZE = 100        # 공공데이터포털 numOfRows 최대치
DATAGOKR_CONFIRM_COUNT = 20000  # 이 건수(2만)를 넘으면 전체 수신 여부 확인

# 검색 필드(응답 필드명) → cond 파라미터 연산자.
# LIKE=포함 검색, EQ=일치. 정의에 없는 필드는 기본 LIKE.
DATAGOKR_COND_OP = {
    "BPLC_NM": "LIKE",
    "ROAD_NM_ADDR": "LIKE",
    "SALS_STTS_CD": "EQ",
    "OPN_ATMY_GRP_CD": "EQ",
}

# 헤더 resultCode 중 일시 오류로 보고 재시도할 코드
DATAGOKR_RETRY_CODES = {"22"}  # LIMITED_NUMBER_OF_SERVICE_REQUESTS_EXCEEDS
# 정상 응답 코드 (게이트웨이마다 '00'/'0'/'000'/'INFO-000' 등으로 제각각)
DATAGOKR_SUCCESS_CODES = {"", "0", "00", "000", "0000", "INFO-000"}


def _datagokr_normalize_items(items) -> list:
    """items.item을 항상 리스트로 정규화.
    - 결과 1건이면 dict 하나로 오므로 리스트로 감싼다.
    - 결과 0건이면 items가 빈 문자열/None이거나 item이 없을 수 있다 → 빈 리스트."""
    if not items:
        return []
    item = items.get("item") if isinstance(items, dict) else items
    if not item:
        return []
    if isinstance(item, dict):
        return [item]
    if isinstance(item, list):
        return item
    return []


def _datagokr_auth_message(auth_msg: str, reason_code: str) -> str:
    return (
        f"공공데이터포털 인증키가 유효하지 않습니다"
        f"({auth_msg or '인증 오류'}"
        f"{', 코드 ' + reason_code if reason_code else ''}).\n\n"
        "· data.go.kr 마이페이지에서 일반 인증키(Decoding)를 확인하고 그대로 입력하세요.\n"
        "· 해당 API 활용신청이 '승인'됐는지 확인하세요.\n"
        "· 승인 직후에는 반영에 시간이 걸릴 수 있습니다(수십 분~1시간)."
    )


def datagokr_fetch(key: str, filters: dict, page_no: int,
                   rows_per_page: int = DATAGOKR_PAGE_SIZE,
                   timeout: int = 180, retries: int = 3):
    """공공데이터포털 일반음식점 API 호출. (total_count, rows) 반환.
    filters는 응답 필드명(BPLC_NM 등) → 검색어이며, cond[FIELD::OP]로 서버에 전달한다.
    시간 초과·일시 오류는 자동 재시도. resultCode가 '00'이 아니면 RuntimeError."""
    # serviceKey는 사용자가 Decoding 키를 입력한다는 전제로 직접 인코딩
    parts = [
        "serviceKey=" + urllib.parse.quote(key, safe=""),
        "pageNo=" + str(page_no),
        "numOfRows=" + str(rows_per_page),
        "returnType=json",
    ]
    for field, value in filters.items():
        op = DATAGOKR_COND_OP.get(field, "LIKE")
        cond_key = "cond[{}::{}]".format(field, op)
        parts.append(
            urllib.parse.quote(cond_key, safe="") + "="
            + urllib.parse.quote(str(value), safe="")
        )
    url = DATAGOKR_URL + "?" + "&".join(parts)

    for attempt in range(retries + 1):
        last = attempt == retries
        try:
            with urllib.request.urlopen(url, timeout=timeout) as resp:
                text = resp.read().decode("utf-8", errors="replace")
        except urllib.error.HTTPError as e:
            # 공공데이터포털 게이트웨이는 인증키 형식이 아예 잘못된 경우
            # JSON/XML 본문 없이 HTTP 401/403으로 바로 거부하기도 한다.
            if e.code in (401, 403):
                raise RuntimeError(_datagokr_auth_message(e.reason, str(e.code)))
            raise RuntimeError(f"HTTP 오류: {e.code} {e.reason}")
        except (TimeoutError, urllib.error.URLError, OSError) as e:
            is_timeout = isinstance(e, TimeoutError) or "timed out" in str(e).lower()
            if is_timeout and not last:
                continue  # 재시도
            if is_timeout:
                raise RuntimeError(
                    "서버 응답 시간 초과입니다.\n\n"
                    "· 대상이 매우 많은 조건(넓은 지역 등)은 서버 처리가 느립니다.\n"
                    "· 사업장명 등 구체적인 조건을 함께 입력하면 빨라집니다.\n"
                    "· 잠시 후 다시 시도해 보세요."
                )
            raise RuntimeError(f"서버 연결 실패: {e}")

        try:
            data = json.loads(text)
        except json.JSONDecodeError:
            # 인증키 오류/미등록 시 게이트웨이가 returnType과 무관하게 XML을 반환한다.
            auth = re.search(r"<returnAuthMsg>([^<]*)</returnAuthMsg>", text)
            code = re.search(r"<returnReasonCode>([^<]*)</returnReasonCode>", text)
            if auth:
                raise RuntimeError(
                    _datagokr_auth_message(auth.group(1), code.group(1) if code else "")
                )
            raise RuntimeError(
                "응답을 JSON으로 해석할 수 없습니다.\n"
                "공공데이터포털 인증키(Decoding)와 활용신청 상태를 확인하세요."
            )

        resp_obj = data.get("response") or {}
        header = resp_obj.get("header") or {}
        result_code = str(header.get("resultCode", "")).strip()
        result_msg = str(header.get("resultMsg", "")).strip()
        msg_upper = result_msg.upper()
        # 코드가 '정상' 계열이거나 메시지가 정상/NORMAL이면 성공으로 간주
        is_success = (
            result_code in DATAGOKR_SUCCESS_CODES
            or "정상" in result_msg
            or msg_upper.startswith("NORMAL")
        )

        # 데이터 없음 → 빈 결과로 처리
        if "NODATA" in msg_upper or (result_code == "03" and not is_success):
            return 0, []
        # 일시 오류(트래픽 초과 등)는 점증 대기 후 재시도
        if result_code in DATAGOKR_RETRY_CODES and not last:
            time.sleep(2 * (attempt + 1))
            continue
        # 성공이 아니면서 오류 코드가 있으면 오류로 처리
        if not is_success and result_code:
            raise RuntimeError(
                f"공공데이터포털 오류({result_code}): {result_msg or '알 수 없는 오류'}"
            )

        body = resp_obj.get("body") or {}
        total = int(body.get("totalCount") or 0)
        rows = _datagokr_normalize_items(body.get("items"))
        return total, rows

    # 이론상 도달하지 않음 (마지막 시도에서 반환/예외)
    raise RuntimeError("알 수 없는 오류로 조회에 실패했습니다.")


def is_row_closed(info: dict, row: dict) -> bool:
    """폐업/말소 등 '영업 중이 아닌' 행인지 데이터셋별로 판정."""
    if info.get("provider") == "data_go_kr":
        # 영업상태코드 01=영업/정상. 코드가 있으면 01이 아닌 것을 폐업으로 간주.
        cd = (row.get("SALS_STTS_CD") or "").strip()
        if cd:
            return cd != "01"
        nm = (row.get("SALS_STTS_NM") or "") + (row.get("DTL_SALS_STTS_NM") or "")
        return ("폐업" in nm) or ("말소" in nm)
    # 식품안전나라: 폐업일자가 채워져 있으면 폐업
    if info.get("has_clsbiz"):
        return bool((row.get("CLSBIZ_DT") or "").strip())
    return False


def supports_open_filter(info: dict) -> bool:
    """'영업 중만 표시' 기능(폐업 제외)을 지원하는 데이터셋인지."""
    return bool(info.get("has_clsbiz")) or info.get("provider") == "data_go_kr"


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


# ── 자동 업데이트 (GitHub Releases) ─────────────────
# 주의: 업데이트는 실행 파일(exe)만 교체하며 config.json(암호화된 인증키)은
#       읽지도, 쓰지도, 옮기지도 않는다. 단위 테스트에서 목으로 바꿀 수 있도록
#       urllib.request.urlopen / subprocess.Popen 을 모듈 경로 그대로 참조한다.

class UpdateCheckError(Exception):
    """업데이트 확인 실패의 기본 예외 (수동 확인 시 메시지로 표시)."""


class UpdateNoReleaseError(UpdateCheckError):
    """저장소에 아직 릴리스가 없음 (HTTP 404)."""


class UpdateRateLimitError(UpdateCheckError):
    """GitHub API 호출 한도 초과 (HTTP 403/429)."""


class UpdateNetworkError(UpdateCheckError):
    """네트워크 연결 실패·시간 초과."""


def parse_version(text) -> tuple:
    """'v1.10.0' → (1, 10). 앞의 v 제거, 각 마디의 앞자리 숫자만 사용.
    끝의 0 마디는 제거하여 '1.0' 과 '1.0.0' 을 같은 버전으로 취급한다."""
    s = str(text or "").strip().lstrip("vV")
    s = re.split(r"[-+ ]", s, maxsplit=1)[0]  # 1.2.3-beta, 1.2.3+build 꼬리 제거
    parts = []
    for piece in s.split("."):
        m = re.match(r"\d+", piece)
        parts.append(int(m.group(0)) if m else 0)
    while parts and parts[-1] == 0:
        parts.pop()
    return tuple(parts)


def is_newer_version(remote: str, local: str = APP_VERSION) -> bool:
    """remote가 local보다 높은 버전이면 True."""
    return parse_version(remote) > parse_version(local)


def check_latest_release(timeout=10, current_version: str = APP_VERSION):
    """GitHub 최신 릴리스를 확인한다.
    새 버전이고 EXE_ASSET_NAME 에셋이 있으면 dict(version, tag, notes, asset_url,
    size, digest, html_url)를, 아니면 None을 반환한다.
    실패 시 UpdateCheckError 계열 예외를 던진다(호출부가 종류별로 구분)."""
    req = urllib.request.Request(
        UPDATE_API_URL,
        headers={
            "User-Agent": UPDATE_USER_AGENT,
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            text = resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        if e.code == 404:
            raise UpdateNoReleaseError("아직 게시된 릴리스가 없습니다.") from e
        remaining = ""
        try:
            remaining = (e.headers or {}).get("X-RateLimit-Remaining", "") or ""
        except Exception:
            remaining = ""
        if e.code == 429 or (e.code == 403 and str(remaining).strip() == "0"):
            raise UpdateRateLimitError(
                "GitHub 업데이트 서버 호출 한도를 초과했습니다. 잠시 후(최대 1시간) 다시 시도하세요."
            ) from e
        raise UpdateCheckError(f"업데이트 서버 오류: HTTP {e.code} {e.reason}") from e
    except (TimeoutError, urllib.error.URLError, OSError) as e:
        raise UpdateNetworkError(f"업데이트 서버에 연결할 수 없습니다: {e}") from e

    try:
        data = json.loads(text)
    except json.JSONDecodeError as e:
        raise UpdateCheckError("업데이트 서버 응답을 해석할 수 없습니다.") from e
    if not isinstance(data, dict):
        raise UpdateCheckError("업데이트 서버 응답 형식이 올바르지 않습니다.")

    tag = str(data.get("tag_name") or "").strip()
    if not tag:
        raise UpdateCheckError("릴리스에 버전 태그가 없습니다.")
    version = tag.lstrip("vV")
    if data.get("draft") or data.get("prerelease"):
        return None
    if not is_newer_version(version, current_version):
        return None

    asset = None
    for a in data.get("assets") or []:
        if isinstance(a, dict) and a.get("name") == EXE_ASSET_NAME:
            asset = a
            break
    if asset is None or not asset.get("browser_download_url"):
        return None  # 실행 파일이 첨부되지 않은 릴리스는 무시

    try:
        size = int(asset.get("size") or 0)
    except (TypeError, ValueError):
        size = 0
    return {
        "version": version,
        "tag": tag,
        "notes": str(data.get("body") or ""),
        "asset_url": asset["browser_download_url"],
        "size": size,
        "digest": str(asset.get("digest") or ""),
        "html_url": str(data.get("html_url") or ""),
    }


def download_update(info: dict, dest_path, progress_cb=None, timeout=60,
                    chunk_size=64 * 1024) -> Path:
    """info["asset_url"]을 dest_path로 내려받는다(임시명 dest.part → 검증 후 rename).
    progress_cb(받은바이트, 전체바이트)를 청크마다 호출한다.
    크기·SHA-256(digest가 'sha256:<hex>'일 때) 불일치 또는 오류 시 파일을 지우고 RuntimeError."""
    dest = Path(dest_path)
    part = dest.with_name(dest.name + ".part")
    expected_size = int(info.get("size") or 0)
    digest = str(info.get("digest") or "").strip()
    expected_sha = ""
    if digest.lower().startswith("sha256:"):
        expected_sha = digest.split(":", 1)[1].strip().lower()

    def _remove(p: Path):
        try:
            if p.exists():
                p.unlink()
        except OSError:
            pass

    sha = hashlib.sha256()
    received = 0
    try:
        req = urllib.request.Request(
            info["asset_url"],
            headers={"User-Agent": UPDATE_USER_AGENT, "Accept": "application/octet-stream"},
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            total = expected_size
            if not total:
                try:
                    total = int(resp.headers.get("Content-Length") or 0)
                except (TypeError, ValueError, AttributeError):
                    total = 0
            with open(part, "wb") as f:
                while True:
                    chunk = resp.read(chunk_size)
                    if not chunk:
                        break
                    f.write(chunk)
                    sha.update(chunk)
                    received += len(chunk)
                    if progress_cb is not None:
                        progress_cb(received, total)
    except Exception as e:
        _remove(part)
        if isinstance(e, RuntimeError):
            raise
        raise RuntimeError(f"업데이트 파일 다운로드 실패: {e}") from e

    if expected_size and received != expected_size:
        _remove(part)
        raise RuntimeError(
            f"다운로드한 파일 크기가 다릅니다 (예상 {expected_size:,}바이트, 실제 {received:,}바이트)."
        )
    if received == 0:
        _remove(part)
        raise RuntimeError("다운로드한 파일이 비어 있습니다.")
    if expected_sha and sha.hexdigest().lower() != expected_sha:
        _remove(part)
        raise RuntimeError("다운로드한 파일의 SHA-256 해시가 일치하지 않습니다 (파일 손상 가능).")

    try:
        os.replace(part, dest)  # 기존 dest가 있으면 덮어씀
    except OSError as e:
        _remove(part)
        raise RuntimeError(f"다운로드한 파일을 저장할 수 없습니다: {e}") from e
    return dest


def old_exe_path(exe_path) -> Path:
    """교체 시 현재 exe를 옮겨 둘 경로: RestaurantFinder.exe → RestaurantFinder.old.exe"""
    p = Path(exe_path)
    return p.with_name(f"{p.stem}.old{p.suffix}")


def new_exe_path(exe_path) -> Path:
    """새 버전을 내려받을 경로: RestaurantFinder.exe → RestaurantFinder.new.exe"""
    p = Path(exe_path)
    return p.with_name(f"{p.stem}.new{p.suffix}")


def _child_env() -> dict:
    """재시작할 새 exe에 넘길 환경 변수.
    PyInstaller onefile 부모의 임시 폴더(_MEIxxxx)를 재사용하지 않고
    독립된 새 인스턴스로 시작하도록 PyInstaller 내부 변수를 정리한다."""
    env = dict(os.environ)
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"  # PyInstaller 6.9+ 공식 방법
    for k in list(env):
        if k.startswith("_PYI_") or k in ("_MEIPASS2",):
            env.pop(k, None)
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        # tkinter 런타임 훅이 설정한 TCL/TK 경로가 부모 임시 폴더를 가리키면 제거
        for k in ("TCL_LIBRARY", "TK_LIBRARY"):
            v = env.get(k, "")
            if v and os.path.normcase(v).startswith(os.path.normcase(meipass)):
                env.pop(k, None)
    return env


def apply_update(current_exe, new_exe, relaunch=True):
    """실행 중인 exe를 새 exe로 교체한다. (Windows는 실행 중인 exe의 이름 변경을 허용)
    1) current → <이름>.old.exe  2) new → current  3) relaunch면 새 exe 실행.
    중간에 실패하면 원래 상태로 롤백을 시도한 뒤 RuntimeError.
    config.json 등 exe 이외의 파일은 건드리지 않는다."""
    cur = Path(current_exe)
    new = Path(new_exe)
    old = old_exe_path(cur)
    if not new.is_file():
        raise RuntimeError(f"새 버전 파일이 없습니다: {new.name}")
    if not cur.is_file():
        raise RuntimeError(f"현재 실행 파일을 찾을 수 없습니다: {cur.name}")

    # 이전 업데이트에서 남은 .old.exe 정리 (아직 실행 중이면 지워지지 않을 수 있음)
    if old.exists():
        try:
            old.unlink()
        except OSError:
            # 지울 수 없으면 고유한 이름으로 비켜 둔다 (다음 시작 시 정리)
            old = cur.with_name(f"{cur.stem}.old-{int(time.time())}{cur.suffix}")

    # 1) 현재 exe를 .old로 이동
    try:
        os.rename(cur, old)
    except OSError as e:
        raise RuntimeError(f"현재 실행 파일 이름을 바꿀 수 없습니다: {e}") from e

    # 2) 새 exe를 현재 경로로 이동
    try:
        os.rename(new, cur)
    except OSError as e:
        try:
            os.rename(old, cur)  # 롤백
        except OSError:
            pass
        raise RuntimeError(f"새 버전으로 교체하지 못했습니다(원래 버전으로 복구): {e}") from e

    # 3) 새 버전 실행
    if relaunch:
        flags = (getattr(subprocess, "DETACHED_PROCESS", 0x00000008)
                 | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0x00000200))
        try:
            subprocess.Popen(
                [str(cur)], close_fds=True, creationflags=flags,
                cwd=str(cur.parent), env=_child_env(),
            )
        except Exception as e:
            # 새 버전을 실행할 수 없음 → 원상복구 (cur→new, old→cur)
            try:
                os.rename(cur, new)
                os.rename(old, cur)
            except OSError:
                pass
            raise RuntimeError(f"새 버전을 실행하지 못했습니다(원래 버전으로 복구): {e}") from e
    return True


def cleanup_old_exe(exe_path=None, include_new=True) -> list:
    """exe 옆에 남은 업데이트 잔여 파일을 지운다(실패는 무시). 지운 파일 목록 반환.
    - <이름>.old.exe, <이름>.old-*.exe : 이전 버전 (이전 프로세스 종료 중이면 실패할 수 있음)
    - include_new=True면 <이름>.new.exe, <이름>.new.exe.part 도 삭제
      (다운로드 결과는 세션 메모리에만 보관하므로, 시작 시 남은 파일은 버리고 다시 확인)
    exe_path가 없으면 frozen(PyInstaller) 실행일 때만 sys.executable 기준으로 동작."""
    if exe_path is None:
        if not getattr(sys, "frozen", False):
            return []
        exe_path = sys.executable
    exe = Path(exe_path)
    targets = [old_exe_path(exe)]
    try:
        targets += sorted(exe.parent.glob(f"{exe.stem}.old-*{exe.suffix}"))
    except OSError:
        pass
    if include_new:
        newp = new_exe_path(exe)
        targets += [newp, newp.with_name(newp.name + ".part")]
    removed = []
    for t in targets:
        try:
            if t.is_file() and t.resolve() != exe.resolve():
                t.unlink()
                removed.append(t)
        except OSError:
            pass
    return removed


# ── GUI ─────────────────────────────────────────
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.rows = []              # 표에 표시 중인(조건 일치) 결과
        self.rows_received = 0      # 서버에서 수신한 건수
        self._row_by_item = {}      # 트리 항목 id → 원본 row(dict) 매핑
        self._sort_field = None     # 현재 정렬 기준 필드
        self._sort_desc = False     # 내림차순 여부
        self._base_headings = {}    # 필드 → 정렬 화살표 없는 원래 헤더 텍스트
        self._timer_id = None
        self._elapsed = 0
        self._cond_desc = ""
        self._page_info = ""
        self._cancel = False
        self._open_cnt = 0
        self._open_only_pref = True  # '영업 중만 표시' 기본값
        self._current_service = "MOIS_I1"
        # _build_main이 키 UI를 mfds 기준으로 초기화하므로 여기선 mfds로 두고,
        # 아래 select_service("MOIS_I1")가 data_go_kr로 전환하도록 한다.
        self._current_provider = "mfds"
        self._key_by_provider = {}        # provider → 입력 중이던 인증키 기억
        self._save_by_provider = {}       # provider → 저장 체크 상태 기억

        root.title("음식점 정보 조회 대시보드 — 식품안전나라 · 행정안전부 공공데이터")
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
        self.select_service("MOIS_I1", initial=True)

        # 자동 업데이트 상태 (이 세션 메모리에만 보관)
        self._upd_state = "idle"      # idle | checking | downloading | ready
        self._upd_info = None         # check_latest_release 결과 dict
        self._upd_path = None         # 검증 완료된 새 exe 경로(Path)
        self._upd_percent = -1
        self._upd_on_close = False    # 닫을 때 적용 예약 여부
        try:
            self._init_updater()
        except Exception:
            pass  # 업데이트 기능 오류가 본 기능을 막지 않도록 격리

    def _sync_ime_font(self):
        """Windows 한글 IME 조합 글자가 입력칸 폰트보다 커 보이는 문제 방지.
        현재 포커스 창의 IME 조합 폰트를 입력칸 폰트(맑은 고딕)에 맞춘다."""
        try:
            import ctypes
            from ctypes import wintypes

            hwnd = ctypes.windll.user32.GetFocus()
            if not hwnd:
                return
            imm = ctypes.windll.imm32
            himc = imm.ImmGetContext(hwnd)
            if not himc:
                return

            class LOGFONTW(ctypes.Structure):
                _fields_ = [
                    ("lfHeight", wintypes.LONG), ("lfWidth", wintypes.LONG),
                    ("lfEscapement", wintypes.LONG), ("lfOrientation", wintypes.LONG),
                    ("lfWeight", wintypes.LONG),
                    ("lfItalic", ctypes.c_byte), ("lfUnderline", ctypes.c_byte),
                    ("lfStrikeOut", ctypes.c_byte), ("lfCharSet", ctypes.c_byte),
                    ("lfOutPrecision", ctypes.c_byte), ("lfClipPrecision", ctypes.c_byte),
                    ("lfQuality", ctypes.c_byte), ("lfPitchAndFamily", ctypes.c_byte),
                    ("lfFaceName", ctypes.c_wchar * 32),
                ]

            dpi = self.root.winfo_fpixels("1i")
            px = int(round(10 * dpi / 72.0))  # 10pt를 현재 DPI 기준 픽셀 높이로
            lf = LOGFONTW()
            lf.lfHeight = -px
            lf.lfWeight = 400
            lf.lfCharSet = 129  # HANGEUL_CHARSET
            lf.lfFaceName = FONT
            imm.ImmSetCompositionFontW(himc, ctypes.byref(lf))
            imm.ImmReleaseContext(hwnd, himc)
        except Exception:
            pass  # IME 폰트 설정 실패해도 입력 자체에는 지장 없음

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
            sb, text="🍽 Restaurant Finder", bg=C["sidebar"], fg="white",
            font=(FONT, 14, "bold"), anchor="w",
        ).pack(fill="x", padx=20, pady=(24, 0))
        tk.Label(
            sb, text="식품안전나라 · 행정안전부\n공공데이터 오픈 API", bg=C["sidebar"],
            fg=C["sidebar_dim"], font=(FONT, 9), anchor="w", justify="left",
        ).pack(fill="x", padx=22, pady=(0, 22))

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
        # 업데이트 확인 버튼 (도움말 버튼과 같은 스타일, 한 단계 작게)
        self.btn_update = tk.Label(
            bottom, text="🔄  업데이트 확인", bg=C["sidebar_hi"], fg=C["sidebar_tx"],
            font=(FONT, 9, "bold"), pady=6, cursor="hand2",
        )
        self.btn_update.pack(fill="x", padx=14, pady=(0, 10))
        self.btn_update.bind("<Button-1>", lambda e: self.check_for_update(manual=True))
        tk.Label(
            bottom, text="식품안전나라 · 행정안전부\n공공데이터 오픈 API",
            bg=C["sidebar"], fg=C["sidebar_dim"], font=(FONT, 8), justify="left",
        ).pack(fill="x", padx=20)
        # 현재 버전 표시
        tk.Label(
            bottom, text=f"v{APP_VERSION}", bg=C["sidebar"], fg=C["sidebar_dim"],
            font=(FONT, 8), anchor="w",
        ).pack(fill="x", padx=20, pady=(4, 0))

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

        # 인증키 (헤더 오른쪽) — 데이터셋 제공처에 따라 라벨/키가 전환됨
        keybox = tk.Frame(header, bg=C["bg"])
        keybox.pack(side="right")
        keyrow = tk.Frame(keybox, bg=C["bg"])
        keyrow.pack(side="top", anchor="e")
        self.var_key_label = tk.StringVar(value="인증키(식품안전나라)")
        tk.Label(
            keyrow, textvariable=self.var_key_label, bg=C["bg"], fg=C["muted"], font=(FONT, 9),
        ).pack(side="left", padx=(0, 6))
        saved_key = load_saved_key("mfds")
        self.var_key = tk.StringVar(value=saved_key or "sample")
        self.ent_key = ttk.Entry(keyrow, textvariable=self.var_key, width=30, show="*")
        self.ent_key.pack(side="left")
        self.var_show = tk.BooleanVar(value=False)
        tk.Checkbutton(
            keyrow, text="표시", variable=self.var_show, bg=C["bg"], font=(FONT, 9),
            command=lambda: self.ent_key.config(show="" if self.var_show.get() else "*"),
        ).pack(side="left", padx=2)
        self.var_save_key = tk.BooleanVar(value=bool(saved_key))
        tk.Checkbutton(
            keyrow, text="저장(암호화)", variable=self.var_save_key, bg=C["bg"], font=(FONT, 9),
        ).pack(side="left")
        # 공공데이터포털 전용 안내(그 외에는 빈 문자열로 숨김)
        self.var_key_hint = tk.StringVar(value="")
        tk.Label(
            keybox, textvariable=self.var_key_hint, bg=C["bg"], fg=C["accent"],
            font=(FONT, 8), anchor="e", justify="right",
        ).pack(side="top", anchor="e")

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
            # 한글 IME 조합이 매끄럽도록 classic tk.Entry 사용 (ttk.Entry는 조합 중
            # 글자가 커 보이는 문제가 있음). 테두리는 ttk 룩에 가깝게 직접 지정.
            ent = tk.Entry(
                cond, textvariable=var_value, width=36, font=(FONT, 10),
                relief="solid", bd=1, highlightthickness=1,
                highlightbackground=C["border"], highlightcolor=C["accent"],
                disabledbackground="#F3F4F6",
            )
            ent.grid(row=1 + i, column=2, sticky="w", padx=8, pady=3, ipady=2)
            # IME가 마지막 글자를 확정한 뒤 검색되도록 after_idle로 지연
            ent.bind("<Return>", lambda e: self.root.after_idle(self.search))
            # 포커스를 받을 때 IME 조합 글자 폰트를 입력칸 폰트에 맞춤(확대 방지)
            ent.bind("<FocusIn>", lambda e: self.root.after(1, self._sync_ime_font))
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
        self.tree.bind("<Button-3>", self._show_cell_menu)  # 우클릭 복사 메뉴
        self.tree.bind("<Control-c>", lambda e: self._copy_row())
        self.tree.tag_configure("odd", background="#F8F9FE")
        self.tree.tag_configure("closed", foreground="#9CA3AF")

        # 우클릭 복사 메뉴
        self._menu = tk.Menu(self.root, tearoff=0)
        self._menu.add_command(label="이 셀 값 복사", command=self._copy_cell)
        self._menu.add_command(label="이 행 전체 복사", command=self._copy_row)
        self._menu.add_command(label="이 열 전체 복사", command=self._copy_column)
        self._menu_target = None  # (item, field)

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
        # 업데이트 상태(다운로드 진행률 등) — 조회 상태 문구와 겹치지 않도록 별도 칸
        self.var_update = tk.StringVar(value="")
        tk.Label(
            bottom, textvariable=self.var_update, bg=C["bg"], fg=C["accent"], font=(FONT, 9),
        ).pack(side="right", padx=(0, 12))

    # ── 동작 ────────────────────────────────
    @property
    def service_id(self) -> str:
        return self._current_service

    def _switch_provider_key(self, new_provider: str):
        """데이터셋 제공처가 바뀌면 헤더 인증키 입력란을 해당 provider용으로 전환.
        전환 전 현재 입력값을 이전 provider용으로 기억하고, 저장 체크 시 암호화 저장."""
        old = self._current_provider
        if new_provider == old:
            return
        # 현재 입력값/저장 상태를 이전 provider용으로 기억
        cur = self.var_key.get().strip()
        self._key_by_provider[old] = cur
        self._save_by_provider[old] = self.var_save_key.get()
        try:
            if self.var_save_key.get() and cur and cur != "sample":
                save_key(cur, old)
        except OSError:
            pass  # 저장 실패해도 전환 자체는 진행

        # 새 provider의 키/저장 상태 복원 (메모리 기억 우선, 없으면 저장분)
        self._current_provider = new_provider
        if new_provider in self._key_by_provider:
            newkey = self._key_by_provider[new_provider]
            newsave = self._save_by_provider.get(new_provider, bool(newkey))
        else:
            newkey = load_saved_key(new_provider)
            newsave = bool(newkey)
        if not newkey and new_provider == "mfds":
            newkey = "sample"  # 식품안전나라 기본 테스트 키
        self.var_key.set(newkey)
        self.var_save_key.set(newsave)

        # 라벨/안내 문구 갱신
        if new_provider == "data_go_kr":
            self.var_key_label.set("인증키(공공데이터포털)")
            self.var_key_hint.set("Decoding 키(일반 인증키)를 입력하세요")
        else:
            self.var_key_label.set("인증키(식품안전나라)")
            self.var_key_hint.set("")

    def select_service(self, sid: str, initial: bool = False):
        self._current_service = sid
        info = SERVICES[sid]
        self._switch_provider_key(info.get("provider", "mfds"))
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

        if supports_open_filter(info):
            self.chk_open.config(state="normal")
            self.var_open_only.set(self._open_only_pref)
        else:
            # 폐업 필터를 지원하지 않는 데이터셋으로 전환: 직전 선택을 기억해 두고 비활성화
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
        self._row_by_item = {}
        self._sort_field = None
        self._sort_desc = False
        self._base_headings = dict(columns)
        self.tree["columns"] = list(columns.keys())
        for field, label in columns.items():
            width = 95 if field.endswith(("_DT", "YMD")) or field == "YEAR" else 130
            if field in ("LOCP_ADDR", "BSSH_NM",
                         "ROAD_NM_ADDR", "LOTNO_ADDR", "BPLC_NM"):
                width = 240
            if field == "PRSDNT_NM":
                width = 85
            # 헤더 클릭 시 해당 필드로 정렬
            self.tree.heading(field, text=label, command=lambda f=field: self._sort_by(f))
            self.tree.column(field, width=width, anchor="w")

    def search(self):
        key = self.var_key.get().strip()
        if not key:
            messagebox.showwarning("확인", "인증키를 입력하세요.")
            return

        service = self.service_id
        info = SERVICES[service]
        provider = info.get("provider", "mfds")
        filters = self.collect_filters()
        self._last_filters = filters

        if not filters:
            if provider == "data_go_kr":
                warn = (
                    "검색 조건이 없습니다.\n"
                    "조건 없이 조회하면 전국 데이터를 100건씩 순서대로 받습니다"
                    "(양이 매우 많을 수 있음).\n계속할까요?"
                )
            else:
                warn = (
                    "검색 조건이 없습니다.\n"
                    "조건 없이 조회하면 처음 1,000건만 표시합니다.\n계속할까요?"
                )
            if not messagebox.askyesno("확인", warn):
                return

        # 키 저장 처리 (DPAPI 암호화) — 현재 provider 필드에 저장
        try:
            if self.var_save_key.get() and key != "sample":
                save_key(key, provider)
            else:
                save_key("", provider)
        except OSError as e:
            messagebox.showwarning("키 저장", str(e))
        # 메모리에도 현재 입력값/저장 상태 반영 (데이터셋 전환 시 유지)
        self._key_by_provider[provider] = key
        self._save_by_provider[provider] = self.var_save_key.get()
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
        if provider == "data_go_kr":
            # 공공데이터포털은 서버가 다중 조건 AND를 지원 → 모든 조건을 전송
            if len(filters) > 1:
                self._cond_desc += " · 서버에서 모든 조건 AND 검색"
        else:
            server_filters = pick_server_filters(filters)
            if len(filters) > 1:
                sf = next(iter(server_filters))
                self._cond_desc += f" · 서버검색: {label_of.get(sf, sf)}, 나머지는 자동 필터링"
        self._page_info = "1페이지 요청 중"
        self.btn_search.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.start_progress()

        open_only = self.var_open_only.get() and supports_open_filter(info)
        threading.Thread(
            target=self._worker, args=(key, service, filters, open_only), daemon=True
        ).start()

    def _worker(self, key: str, service: str, filters: dict, open_only: bool):
        """데이터 수신 작업 스레드. 제공처에 따라 페이징 방식을 분기."""
        if SERVICES[service].get("provider") == "data_go_kr":
            self._worker_datagokr(key, service, filters, open_only)
            return
        self._worker_mfds(key, service, filters, open_only)

    def _worker_mfds(self, key: str, service: str, filters: dict, open_only: bool):
        """식품안전나라: 조건에 맞는 서버 데이터 전체를 1,000건씩 이어서 수신.
        서버에는 가장 선별력 높은 조건 1개만 보내고, 전체 조건 검증은 _on_page에서 수행."""
        server_filters = pick_server_filters(filters)
        start = 1
        total = None
        pages = 1
        page_no = 0
        try:
            while not self._cancel:
                url = build_url(key, service, start, start + PAGE_SIZE - 1, server_filters)
                try:
                    t, rows = api_fetch(url, service)
                except RuntimeError as e:
                    if page_no == 0:
                        raise  # 첫 페이지부터 실패 → 전체 오류로 처리
                    # 이미 일부 페이지를 받은 상태 → 부분 결과 유지하고 경고
                    self.root.after(0, self._on_done, total or 0, pages, str(e))
                    return
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

    def _worker_datagokr(self, key: str, service: str, filters: dict, open_only: bool):
        """공공데이터포털 일반음식점: 100건/페이지 pageNo 기반 페이징.
        서버가 다중 조건 AND를 지원하므로 모든 조건을 cond[...]로 전달한다."""
        server_filters = dict(filters)  # 모든 사용자 조건을 서버로
        if open_only:
            server_filters["SALS_STTS_CD"] = "01"  # 영업/정상만 서버에서 선필터
        page_no = 0
        total = None
        pages = 1
        try:
            while not self._cancel:
                try:
                    t, rows = datagokr_fetch(key, server_filters, page_no + 1)
                except RuntimeError as e:
                    if page_no == 0:
                        raise  # 첫 페이지부터 실패 → 전체 오류로 처리
                    # 이미 일부 페이지를 받은 상태 → 부분 결과 유지하고 경고
                    self.root.after(0, self._on_done, total or 0, pages, str(e))
                    return
                if total is None:
                    total = t
                    pages = max(1, (total + DATAGOKR_PAGE_SIZE - 1) // DATAGOKR_PAGE_SIZE)
                page_no += 1
                self.root.after(0, self._on_page, total, rows, page_no, pages, open_only)

                if not rows or page_no >= pages:
                    break
                # 대량 데이터(2만 건 초과)는 전체 수신 여부를 한 번 확인
                if page_no == 1 and total > DATAGOKR_CONFIRM_COUNT:
                    if not self._ask_main(
                        f"조건과 일치하는 데이터가 {total:,}건({pages}페이지)입니다.\n"
                        f"공공데이터포털은 100건씩 나눠 받으므로 시간이 오래 걸릴 수 있습니다.\n\n"
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
            matched = [r for r in matched if not is_row_closed(info, r)]

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
            if is_row_closed(info, row):
                tags.append("closed")  # 폐업/말소 업소는 회색으로 표시
            item = self.tree.insert("", "end", values=values, tags=tags)
            self._row_by_item[item] = row

        if supports_open_filter(info):
            self._open_cnt += sum(
                1 for r in matched if not is_row_closed(info, r)
            )
        self.rows.extend(matched)

        self.kpi_vars["total"].set(f"{total:,}")
        self.kpi_vars["received"].set(f"{self.rows_received:,}")
        self.kpi_vars["matched"].set(f"{len(self.rows):,}")
        self.kpi_vars["open"].set(
            f"{self._open_cnt:,}" if supports_open_filter(info) else "해당 없음"
        )
        self._page_info = f"페이지 {page_no}/{pages} 수신, 일치 {len(self.rows):,}건"

    def _on_done(self, total: int, pages: int, partial_error: str = ""):
        self.stop_progress()
        self.btn_search.config(state="normal")
        self.btn_stop.config(state="disabled")

        n = len(self.rows)
        if partial_error:
            prefix = "일부만 수신(서버 오류로 중단)"
        elif self._cancel:
            prefix = "중지됨"
        else:
            prefix = "완료"
        msg = f"{prefix}: 서버 일치 {total:,}건 중 {self.rows_received:,}건 수신 → 조건 일치 {n:,}건"
        if n > MAX_TREE_ROWS:
            msg += f" (표에는 처음 {MAX_TREE_ROWS:,}건만 표시, 전체는 CSV/JSON 저장으로 확인)"
        self.var_status.set(msg)

        # 사용자가 직접 정렬하지 않았으면 인허가일자 최근순으로 기본 정렬
        info = SERVICES[self.service_id]
        default_sort = info.get("default_sort")
        if n and self._sort_field is None and default_sort in info["columns"]:
            self._sort_by(default_sort, desc=True)

        if partial_error:
            messagebox.showwarning(
                "일부만 수신됨",
                f"수신 도중 서버 오류가 발생해 중단했습니다.\n"
                f"지금까지 받은 {n:,}건은 표에 그대로 남아 있습니다.\n\n{partial_error}",
            )
            return

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
        row = self._row_by_item.get(sel[0])
        if row is None:
            return
        info = SERVICES[self.service_id]

        win = tk.Toplevel(self.root)
        win.title(row.get("BSSH_NM") or row.get("BPLC_NM") or "상세 정보")
        win.geometry("500x440")
        win.configure(bg=C["card"])
        txt = tk.Text(win, wrap="word", font=(FONT, 10), relief="flat", bg=C["card"])
        txt.pack(fill="both", expand=True, padx=12, pady=12)
        for field, value in row.items():
            label = info["columns"].get(field, field)
            txt.insert("end", f"{label} ({field})\n    {value or '-'}\n\n")
        txt.config(state="disabled")

    # ── 정렬 ────────────────────────────────
    def _sort_by(self, field: str, desc: bool = None):
        """해당 필드 기준 정렬.
        desc=None(헤더 클릭): 같은 필드 재클릭 시 오름/내림 토글.
        desc 지정(기본 정렬 등): 방향을 명시적으로 설정."""
        items = list(self.tree.get_children(""))
        if not items:
            return
        if desc is not None:
            self._sort_field = field
            self._sort_desc = desc
        elif self._sort_field == field:
            self._sort_desc = not self._sort_desc
        else:
            self._sort_field = field
            self._sort_desc = False

        def key(item):
            return (self.tree.set(item, field) or "").strip()

        def as_num(s):
            try:
                return float(s.replace(",", ""))
            except ValueError:
                return None

        nonempty = [it for it in items if key(it) != ""]
        empties = [it for it in items if key(it) == ""]
        numeric = all(as_num(key(it)) is not None for it in nonempty)

        nonempty.sort(
            key=(lambda it: as_num(key(it))) if numeric else (lambda it: key(it)),
            reverse=self._sort_desc,
        )
        # 빈 값은 정렬 방향과 무관하게 항상 뒤로
        ordered = nonempty + empties
        for idx, item in enumerate(ordered):
            self.tree.move(item, "", idx)

        self._restripe()
        self._update_sort_indicator()

    def _restripe(self):
        """정렬 후 줄무늬(odd)와 폐업(closed) 태그를 표시 순서에 맞게 재적용."""
        info = SERVICES[self.service_id]
        for idx, item in enumerate(self.tree.get_children("")):
            tags = ["odd"] if idx % 2 == 1 else []
            row = self._row_by_item.get(item, {})
            if is_row_closed(info, row):
                tags.append("closed")
            self.tree.item(item, tags=tags)

    def _update_sort_indicator(self):
        """헤더 텍스트에 정렬 방향 화살표(▲/▼) 표시."""
        arrow = " ▼" if self._sort_desc else " ▲"
        for field, label in self._base_headings.items():
            text = label + (arrow if field == self._sort_field else "")
            self.tree.heading(field, text=text)

    # ── 복사 ────────────────────────────────
    def _show_cell_menu(self, event):
        """우클릭한 셀을 선택하고 복사 메뉴 표시."""
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)  # '#1', '#2' ...
        if not item or not col:
            return
        self.tree.selection_set(item)
        try:
            field = self.tree["columns"][int(col[1:]) - 1]
        except (ValueError, IndexError):
            return
        self._menu_target = (item, field)
        label = self._base_headings.get(field, field)
        self._menu.entryconfig(0, label=f"'{label}' 셀 값 복사")
        self._menu.entryconfig(2, label=f"'{label}' 열 전체 복사")
        self._menu.tk_popup(event.x_root, event.y_root)

    def _to_clipboard(self, text: str, desc: str):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.var_status.set(f"클립보드에 복사됨: {desc}")

    def _copy_cell(self):
        if not self._menu_target:
            return
        item, field = self._menu_target
        value = self.tree.set(item, field)
        self._to_clipboard(value, f"{self._base_headings.get(field, field)} = {value or '(빈 값)'}")

    def _copy_row(self):
        sel = self.tree.selection()
        if not sel:
            return
        values = self.tree.item(sel[0], "values")
        self._to_clipboard("\t".join(str(v) for v in values), "선택한 행 전체")

    def _copy_column(self):
        if not self._menu_target:
            return
        _item, field = self._menu_target
        col_vals = [self.tree.set(it, field) for it in self.tree.get_children("")]
        self._to_clipboard(
            "\n".join(col_vals),
            f"'{self._base_headings.get(field, field)}' 열 {len(col_vals)}건",
        )

    # ── 자동 업데이트 ─────────────────────────
    @staticmethod
    def _is_frozen() -> bool:
        return bool(getattr(sys, "frozen", False))

    def _is_searching(self) -> bool:
        """조회(수신) 진행 중인지. 조회 중에는 [조회] 버튼이 비활성화된다."""
        try:
            return str(self.btn_search.cget("state")) == "disabled"
        except Exception:
            return False

    def _post(self, fn, *args):
        """작업 스레드 → 메인 스레드로 tkinter 호출 전달 (창이 닫혔으면 무시)."""
        try:
            self.root.after(0, fn, *args)
        except Exception:
            pass

    def _set_update_status(self, text: str):
        try:
            self.var_update.set(text)
        except Exception:
            pass

    def _init_updater(self):
        """시작 시 잔여 파일 정리 + 약 2초 뒤 자동 업데이트 확인 예약."""
        if self._is_frozen():
            cleanup_old_exe(include_new=True)
            # 이전 프로세스가 아직 종료 중이면 .old.exe 삭제가 실패하므로 몇 초 뒤 한 번 더
            # (이때는 진행 중일 수 있는 새 다운로드 파일을 건드리지 않도록 old만 정리)
            self.root.after(5000, lambda: self._safe(cleanup_old_exe, None, False))
        self.root.after(2000, lambda: self.check_for_update(manual=False))

    @staticmethod
    def _safe(fn, *args):
        try:
            return fn(*args)
        except Exception:
            return None

    def check_for_update(self, manual: bool = False):
        """업데이트 확인 시작(백그라운드 스레드). manual=True면 결과를 대화상자로 알림."""
        try:
            state = self._upd_state
            if state == "checking":
                if manual:
                    messagebox.showinfo("업데이트 확인", "업데이트를 확인하는 중입니다. 잠시만 기다려 주세요.")
                return
            if state == "downloading":
                if manual:
                    pct = f" ({self._upd_percent}%)" if self._upd_percent >= 0 else ""
                    messagebox.showinfo("업데이트 확인", f"새 버전을 내려받는 중입니다{pct}.")
                return
            if state == "ready" and self._upd_info and self._upd_path:
                if manual:
                    self._prompt_apply()
                return
            self._upd_state = "checking"
            if manual:
                self._set_update_status("업데이트 확인 중...")
            threading.Thread(
                target=self._update_check_worker, args=(manual,), daemon=True
            ).start()
        except Exception as e:
            self._upd_state = "idle"
            if manual:
                messagebox.showerror("업데이트 확인", f"업데이트 확인을 시작하지 못했습니다.\n{e}")

    def _update_check_worker(self, manual: bool):
        try:
            info = check_latest_release()
        except Exception as e:
            self._post(self._on_update_check_error, e, manual)
            return
        self._post(self._on_update_checked, info, manual)

    def _on_update_check_error(self, err: Exception, manual: bool):
        self._upd_state = "idle"
        self._set_update_status("")
        if not manual:
            return  # 자동 확인 실패는 조용히 무시
        if isinstance(err, UpdateNoReleaseError):
            messagebox.showinfo(
                "업데이트 확인",
                f"아직 게시된 릴리스가 없습니다.\n현재 버전: v{APP_VERSION}",
            )
        elif isinstance(err, UpdateRateLimitError):
            messagebox.showwarning("업데이트 확인", str(err))
        elif isinstance(err, UpdateNetworkError):
            messagebox.showwarning(
                "업데이트 확인",
                "업데이트 서버에 연결할 수 없습니다.\n인터넷 연결을 확인한 뒤 다시 시도하세요.\n\n"
                f"({err})",
            )
        elif isinstance(err, UpdateCheckError):
            messagebox.showwarning("업데이트 확인", str(err))
        else:
            messagebox.showerror("업데이트 확인", f"업데이트 확인 중 오류가 발생했습니다.\n{err}")

    def _on_update_checked(self, info, manual: bool):
        try:
            if not info:
                self._upd_state = "idle"
                self._set_update_status("")
                if manual:
                    messagebox.showinfo("업데이트 확인", f"최신 버전입니다(v{APP_VERSION}).")
                return

            ver = info["version"]
            if not self._is_frozen():
                # 소스 실행: 다운로드/교체하지 않고 안내만
                self._upd_state = "idle"
                self._set_update_status(f"새 버전 v{ver} 있음 (소스 실행 중이라 자동 적용 안 함)")
                if manual:
                    messagebox.showinfo(
                        "업데이트 확인",
                        f"새 버전 v{ver} 이(가) 있습니다. (현재 v{APP_VERSION})\n\n"
                        "소스(.py)로 실행 중이므로 자동 업데이트를 적용하지 않습니다.\n"
                        f"{info.get('html_url') or 'GitHub 릴리스 페이지'}에서 확인하세요.",
                    )
                return

            # 실행 파일: exe 옆 <이름>.new.exe 로 자동 다운로드
            self._upd_info = info
            self._upd_state = "downloading"
            self._upd_percent = -1
            self._set_update_status(f"새 버전 v{ver} 다운로드 준비 중...")
            dest = new_exe_path(sys.executable)
            threading.Thread(
                target=self._update_download_worker, args=(info, dest, manual), daemon=True
            ).start()
        except Exception as e:
            self._upd_state = "idle"
            self._set_update_status("")
            if manual:
                messagebox.showerror("업데이트", f"업데이트 처리 중 오류가 발생했습니다.\n{e}")

    def _update_download_worker(self, info: dict, dest, manual: bool):
        def progress(done, total):
            if total:
                pct = min(100, int(done * 100 / total))
                if pct != self._upd_percent:  # 1% 단위로만 화면 갱신
                    self._upd_percent = pct
                    self._post(self._set_update_status,
                               f"업데이트 v{info['version']} 다운로드 중... {pct}%")
            else:
                self._post(self._set_update_status,
                           f"업데이트 v{info['version']} 다운로드 중... {done // 1024:,}KB")

        try:
            path = download_update(info, dest, progress_cb=progress)
        except Exception as e:
            self._post(self._on_update_download_error, e, manual)
            return
        self._post(self._on_update_downloaded, info, path)

    def _on_update_download_error(self, err: Exception, manual: bool):
        self._upd_state = "idle"
        self._upd_info = None
        self._upd_path = None
        self._set_update_status("업데이트 다운로드 실패" if not manual else "")
        if manual:
            messagebox.showwarning("업데이트", f"새 버전을 내려받지 못했습니다.\n\n{err}")

    def _on_update_downloaded(self, info: dict, path):
        self._upd_state = "ready"
        self._upd_info = info
        self._upd_path = Path(path)
        self._set_update_status(f"새 버전 v{info['version']} 준비 완료")
        self._prompt_apply()

    @staticmethod
    def _notes_excerpt(notes: str, max_lines: int = 8, max_chars: int = 400) -> str:
        text = (notes or "").replace("\r\n", "\n").strip()
        if not text:
            return ""
        lines = text.split("\n")
        cut = "\n".join(lines[:max_lines])
        if len(cut) > max_chars:
            cut = cut[:max_chars].rstrip()
        if cut != text:
            cut += "\n…"
        return cut

    def _prompt_apply(self):
        """다운로드·검증 완료 후 적용 방법을 묻는다."""
        try:
            info = self._upd_info or {}
            ver = info.get("version", "?")
            notes = self._notes_excerpt(info.get("notes", ""))
            msg = f"새 버전 v{ver} 준비 완료 (현재 v{APP_VERSION})\n"
            if notes:
                msg += f"\n[릴리스 노트]\n{notes}\n"
            msg += (
                "\n지금 재시작해서 적용할까요?\n"
                "(아니오: 프로그램을 닫을 때 자동 적용)\n\n"
                "※ 저장된 인증키(config.json)는 그대로 유지됩니다."
            )
            if messagebox.askyesno("업데이트 준비 완료", msg):
                self._apply_now()
            else:
                self._arm_apply_on_close()
        except Exception as e:
            self._arm_apply_on_close()
            messagebox.showwarning("업데이트", f"업데이트 안내 중 오류가 발생했습니다. 닫을 때 적용합니다.\n{e}")

    def _arm_apply_on_close(self):
        """창을 닫을 때(WM_DELETE_WINDOW) 새 버전을 적용하도록 예약."""
        self._upd_on_close = True
        ver = (self._upd_info or {}).get("version", "")
        self._set_update_status(f"새 버전 v{ver}: 프로그램을 닫을 때 적용됩니다")
        try:
            self.root.protocol("WM_DELETE_WINDOW", self._on_close_apply_update)
        except Exception:
            pass

    def _apply_now(self):
        """즉시 교체 후 재시작. 조회 중이면 결과를 잃지 않도록 '닫을 때 적용'으로 전환."""
        if self._is_searching():
            self._arm_apply_on_close()
            messagebox.showinfo(
                "업데이트",
                "조회가 진행 중이라 지금은 재시작하지 않습니다.\n"
                "프로그램을 닫을 때 새 버전이 자동으로 적용됩니다.",
            )
            return
        try:
            apply_update(sys.executable, self._upd_path, relaunch=True)
        except Exception as e:
            self._arm_apply_on_close()
            messagebox.showerror(
                "업데이트 적용 실패",
                f"새 버전을 적용하지 못했습니다. 현재 버전으로 계속 사용합니다.\n"
                f"(프로그램을 닫을 때 한 번 더 적용을 시도합니다)\n\n{e}",
            )
            return
        self._upd_on_close = False
        try:
            self.root.destroy()
        except Exception:
            pass

    def _on_close_apply_update(self):
        """창 닫기 훅: 예약된 업데이트를 적용(재시작 없이)한 뒤 종료."""
        try:
            if (self._upd_on_close and self._upd_path is not None
                    and Path(self._upd_path).is_file() and self._is_frozen()):
                apply_update(sys.executable, self._upd_path, relaunch=False)
        except Exception:
            pass  # 적용 실패해도 종료는 진행 (다음 시작 시 잔여 파일 정리 후 재확인)
        try:
            self.root.destroy()
        except Exception:
            pass

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
