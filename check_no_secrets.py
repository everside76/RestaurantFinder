# -*- coding: utf-8 -*-
"""
릴리스 전 인증키 유출 검사 (표준 라이브러리만 사용)

저장소가 public이므로 커밋·태그·푸시·릴리스 전에 반드시 실행한다.
  python -I check_no_secrets.py

1. dist\\config.json / 저장소 루트 config.json 에 저장된 인증키를
   Windows DPAPI로 복호화해 실제 키 목록을 얻는다 (없으면 생략).
2. 검사 대상
   (a) git이 추적하는 모든 파일 + 아직 추적되지 않았지만 무시되지도 않는 파일
   (b) 스테이징/커밋 예정 변경분 (git diff --cached, git diff)
   (c) 빌드된 dist\\RestaurantFinder.exe 바이트
   각 대상은 UTF-8, UTF-16-LE 두 가지 인코딩으로 검사한다.
3. 찾는 값: 키 원문 전체, (16자 이상이면) 앞 12자·뒤 12자, 암호문(*_enc 값) 앞 40자
4. git이 추적하는 파일 중 config.json, *.new.exe, *.old.exe, *.part 가 있으면 실패

키 값은 절대 출력하지 않는다(건수와 파일 위치만 출력).
하나라도 발견되면 종료 코드 1, 없으면 0.
"""

import base64
import ctypes
import ctypes.wintypes
import fnmatch
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
EXE_PATH = ROOT / "dist" / "RestaurantFinder.exe"
CONFIG_CANDIDATES = [ROOT / "dist" / "config.json", ROOT / "config.json"]

# 설정 파일 필드 → 표시 이름 (키 값 대신 이 이름만 출력)
KEY_FIELDS = {
    "api_key_enc": "식품안전나라 키",
    "datagokr_key_enc": "공공데이터포털 키",
}
LEGACY_PLAIN_FIELD = "api_key"  # 구버전 평문 저장 필드 (식품안전나라)

# git이 추적해서는 안 되는 파일 패턴 (파일 이름 기준)
FORBIDDEN_TRACKED = ["config.json", "*.new.exe", "*.old.exe", "*.old-*.exe", "*.part"]


# ── DPAPI 복호화 ───────────────────────────────
class _DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ("cbData", ctypes.wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_char)),
    ]


def dpapi_decrypt(b64: str) -> str:
    raw = base64.b64decode(b64)
    buf = ctypes.create_string_buffer(raw, len(raw))
    blob_in = _DATA_BLOB(len(raw), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    blob_out = _DATA_BLOB()
    if not ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)
    ):
        raise OSError("DPAPI 복호화 실패")
    data = ctypes.string_at(blob_out.pbData, blob_out.cbData)
    ctypes.windll.kernel32.LocalFree(blob_out.pbData)
    return data.decode("utf-8")


# ── 찾을 값(needle) 수집 ────────────────────────
def collect_needles():
    """{표시 이름: [bytes 패턴, ...]} 를 반환. 키 값 자체는 출력하지 않는다."""
    needles = {}

    def add(label: str, text: str):
        if not text:
            return
        pats = needles.setdefault(label, set())
        variants = [text]
        if len(text) >= 16:
            variants += [text[:12], text[-12:]]
        for v in variants:
            pats.add(v.encode("utf-8"))
            pats.add(v.encode("utf-16-le"))

    found_config = False
    for cfg_path in CONFIG_CANDIDATES:
        if not cfg_path.is_file():
            continue
        found_config = True
        try:
            rel = cfg_path.relative_to(ROOT)
        except ValueError:
            rel = cfg_path.name
        try:
            cfg = json.loads(cfg_path.read_text(encoding="utf-8") or "{}")
        except (OSError, json.JSONDecodeError):
            print(f"  [경고] {rel}: 읽을 수 없어 건너뜀")
            continue
        if not isinstance(cfg, dict):
            continue
        for field, label in KEY_FIELDS.items():
            enc = cfg.get(field)
            if not enc:
                continue
            # 암호문 앞 40자도 유출 대상으로 검사 (복호화 키와 별도 항목)
            head = str(enc)[:40]
            needles.setdefault(f"{label}(암호문)", set()).update(
                {head.encode("ascii", "ignore"), head.encode("utf-16-le")}
            )
            try:
                add(label, dpapi_decrypt(str(enc)))
            except Exception:
                print(f"  [경고] {rel}: {label} 복호화 실패 (다른 PC/계정에서 저장됨?) - 암호문만 검사")
        plain = cfg.get(LEGACY_PLAIN_FIELD)
        if plain and plain != "sample":
            add("식품안전나라 키(평문 저장분)", str(plain))
    if not found_config:
        print("  config.json 없음 - 키 값 검사는 생략하고 파일 이름 검사만 수행")
    # 빈 패턴 제거
    return {k: [p for p in v if p] for k, v in needles.items() if any(v)}


# ── 검사 대상 수집 ──────────────────────────────
def git(*args) -> bytes:
    res = subprocess.run(
        ["git", *args], cwd=str(ROOT), stdout=subprocess.PIPE, stderr=subprocess.PIPE
    )
    if res.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)} 실패 (종료 코드 {res.returncode})")
    return res.stdout


def git_files(*args) -> list:
    out = git("ls-files", "-z", *args)
    return [p for p in out.decode("utf-8", errors="replace").split("\0") if p]


def collect_targets():
    """[(위치 이름, bytes)] 목록"""
    targets = []
    tracked = git_files()
    untracked = git_files("--others", "--exclude-standard")
    for rel in tracked + untracked:
        p = ROOT / rel
        if p.is_file():
            try:
                targets.append((rel, p.read_bytes()))
            except OSError:
                print(f"  [경고] {rel}: 읽을 수 없어 건너뜀")
    targets.append(("git diff --cached (스테이징된 변경)", git("diff", "--cached")))
    targets.append(("git diff (커밋 예정 변경)", git("diff")))
    if EXE_PATH.is_file():
        targets.append((str(EXE_PATH.relative_to(ROOT)), EXE_PATH.read_bytes()))
    else:
        print(f"  {EXE_PATH.relative_to(ROOT)} 없음 - exe 검사 건너뜀")
    return tracked, targets


def main() -> int:
    # 콘솔 코드페이지(cp949 등)에서 출력 오류로 검사가 중단되지 않도록
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(errors="replace")
        except Exception:
            pass
    print("== 인증키 유출 검사 ==")
    failed = False

    try:
        needles = collect_needles()
        tracked, targets = collect_targets()
    except Exception as e:
        print(f"[실패] 검사를 수행할 수 없습니다: {e}")
        return 1

    # 1) 추적되면 안 되는 파일
    bad_files = [
        rel for rel in tracked
        if any(fnmatch.fnmatch(Path(rel).name.lower(), pat) for pat in FORBIDDEN_TRACKED)
    ]
    if bad_files:
        failed = True
        print(f"[실패] git이 추적 중인 금지 파일 {len(bad_files)}개: {', '.join(bad_files)}")
    else:
        print("금지 파일(config.json, *.new.exe, *.old.exe, *.part) 추적: 0건")

    # 2) 키 값/암호문 검색
    print(f"검사 대상: {len(targets)}곳")
    if not needles:
        print("검사할 저장 키 없음")
    for label, pats in needles.items():
        hits = [name for name, data in targets if any(p in data for p in pats)]
        if hits:
            failed = True
            print(f"[실패] {label}: {len(hits)}건 - 위치: {', '.join(hits)}")
        else:
            print(f"{label}: 0건")

    if failed:
        print("\n결과: 실패 - 인증키 유출 가능성이 있습니다. 커밋·푸시·릴리스를 중단하세요.")
        return 1
    print("\n결과: 통과 - 인증키 유출 없음")
    return 0


if __name__ == "__main__":
    sys.exit(main())
