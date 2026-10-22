# -*- coding: utf-8 -*-
"""
식품안전나라(foodsafetykorea.go.kr) 오픈 API를 이용한 음식점 정보 조회 프로그램

API 요청 형식:
    http://openapi.foodsafetykorea.go.kr/api/{keyId}/{serviceId}/{dataType}/{startIdx}/{endIdx}[/{조건변수=값}]

기본 서비스: I1590 (식품모범음식점 — 전국 모범음식점 지정 현황)
기타 음식점 관련 서비스:
    I2847 : 나트륨 줄이기 실천음식점 지정업체 대장

사용 예:
    python restaurant_api.py                              # 1~20건 조회 (sample 키)
    python restaurant_api.py --name 냉면                   # 업소명에 '냉면' 포함 업소 검색
    python restaurant_api.py --region 경주                 # 시군구명에 '경주' 포함 업소 검색
    python restaurant_api.py --start 1 --end 100          # 조회 범위 지정
    python restaurant_api.py --key 발급받은인증키            # 개인 인증키 사용
    python restaurant_api.py --save result.json           # 결과를 JSON 파일로 저장

※ sample 키는 테스트용으로 5건만 반환됩니다.
   실제 사용 시 식품안전나라(https://www.foodsafetykorea.go.kr) 회원가입 후
   [마이페이지 > 오픈API 인증키 발급]에서 인증키를 발급받아 --key 옵션으로 지정하세요.
"""

import argparse
import json
import sys
import urllib.error
import urllib.parse
import urllib.request

# 정식 인증키는 HTTP 호출 시 오류가 발생하므로 반드시 HTTPS 사용
BASE_URL = "https://openapi.foodsafetykorea.go.kr/api"

DEFAULT_KEY = "sample"     # 테스트용 인증키 (5건 제한)
DEFAULT_SERVICE = "I1590"  # 식품모범음식점

# I1590 응답 필드 명세
FIELD_LABELS = {
    "BSSH_NM": "업소명",
    "SIGNGU_NM": "소재지(시군구)",
    "PNCPL_FOOD_NM": "주된 음식",
    "YEAR": "지정년도",
    "APPN_DT": "지정일자",
    "LCNS_NO": "인허가번호",
}


def format_date(value: str) -> str:
    """YYYYMMDD → YYYY-MM-DD 변환"""
    if value and len(value) == 8 and value.isdigit():
        return f"{value[:4]}-{value[4:6]}-{value[6:]}"
    return value


def build_url(key: str, service: str, start: int, end: int, filters: dict) -> str:
    """API 요청 URL 생성"""
    url = f"{BASE_URL}/{key}/{service}/json/{start}/{end}"
    if filters:
        cond = "&".join(
            f"{k}={urllib.parse.quote(str(v))}" for k, v in filters.items()
        )
        url += f"/{cond}"
    return url


def fetch(url: str) -> dict:
    """API 호출 후 JSON 파싱"""
    try:
        with urllib.request.urlopen(url, timeout=15) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        sys.exit(f"[오류] HTTP 오류가 발생했습니다: {e.code} {e.reason}")
    except urllib.error.URLError as e:
        sys.exit(f"[오류] 서버에 연결할 수 없습니다: {e.reason}")
    except (TimeoutError, OSError) as e:
        sys.exit(f"[오류] 네트워크 오류(시간 초과 등): {e}")
    except json.JSONDecodeError:
        sys.exit("[오류] 응답을 JSON으로 해석할 수 없습니다. 서비스 ID나 인증키를 확인하세요.")


def parse_result(data: dict, service: str):
    """응답에서 결과 코드와 목록 추출"""
    body = data.get(service)
    if body is None:
        # 인증키 오류 등은 서비스 ID 없이 RESULT만 반환됨
        result = data.get("RESULT", {})
        sys.exit(
            f"[오류] {result.get('CODE', '알 수 없음')}: "
            f"{result.get('MSG', '응답 형식이 올바르지 않습니다.')}"
        )

    result = body.get("RESULT", {})
    code = result.get("CODE", "")
    if code != "INFO-000":
        sys.exit(f"[안내] {code}: {result.get('MSG', '')}")

    total = int(body.get("total_count", 0))
    rows = body.get("row", [])
    return total, rows


def print_rows(rows: list, total: int):
    """조회 결과를 보기 좋게 출력"""
    print(f"\n총 {total:,}건 중 {len(rows)}건 조회")
    print("=" * 60)
    for i, row in enumerate(rows, 1):
        print(f"\n[{i}] {row.get('BSSH_NM', '(업소명 없음)')}")
        for field, label in FIELD_LABELS.items():
            if field == "BSSH_NM":
                continue
            value = row.get(field)
            if value:
                if field.endswith("_DT"):
                    value = format_date(value)
                print(f"    {label:<9}: {value}")
    print("\n" + "=" * 60)


def main():
    # Windows 콘솔 한글 깨짐 방지
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(
        description="식품안전나라 오픈 API 음식점(모범음식점) 정보 조회 프로그램",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--key", default=DEFAULT_KEY, help="오픈 API 인증키")
    parser.add_argument("--service", default=DEFAULT_SERVICE, help="서비스 ID")
    parser.add_argument("--start", type=int, default=1, help="조회 시작 위치")
    parser.add_argument("--end", type=int, default=20, help="조회 종료 위치 (최대 1000건)")
    parser.add_argument("--name", help="업소명 검색어 (BSSH_NM)")
    parser.add_argument("--region", help="지역(시군구) 검색어 (SIGNGU_NM)")
    parser.add_argument("--save", metavar="파일명", help="결과를 JSON 파일로 저장")
    args = parser.parse_args()

    if args.end - args.start + 1 > 1000:
        sys.exit("[오류] 한 번에 최대 1,000건까지만 조회할 수 있습니다.")

    filters = {}
    if args.name:
        filters["BSSH_NM"] = args.name
    if args.region:
        filters["SIGNGU_NM"] = args.region

    url = build_url(args.key, args.service, args.start, args.end, filters)
    print(f"요청 URL: {url}")

    data = fetch(url)
    total, rows = parse_result(data, args.service)

    if not rows:
        print("조회된 데이터가 없습니다.")
        return

    print_rows(rows, total)

    if args.save:
        with open(args.save, "w", encoding="utf-8") as f:
            json.dump(rows, f, ensure_ascii=False, indent=2)
        print(f"결과 {len(rows)}건을 '{args.save}' 파일로 저장했습니다.")


if __name__ == "__main__":
    main()
