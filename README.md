# RestaurantFinder — 식품안전나라 오픈 API 음식점 정보 조회 대시보드

식품의약품안전처 [식품안전나라](https://www.foodsafetykorea.go.kr) 오픈 API로
전국 음식점 정보를 조회하는 Windows 프로그램입니다.

## 실행

- **실행 파일**: `dist\RestaurantFinder.exe` (Python 설치 불필요, 더블클릭 실행)
- 소스 실행: `python restaurant_api_gui.py`
- 소스 수정 후 재빌드 (아이콘 포함, 릴리스는 아래 `release.ps1` 사용 권장):
  ```
  python -m PyInstaller --onefile --noconsole --clean --name "RestaurantFinder" --icon "app.ico" --add-data "app.ico;." restaurant_api_gui.py
  ```
  결과물은 `dist\RestaurantFinder.exe`. 빌드 후 생기는 `build\` 폴더와 `RestaurantFinder.spec`는 지워도 됩니다.
- 아이콘 다시 생성: `python make_icon.py` (Pillow 필요, `app.ico` 재생성)
- 버전: `restaurant_api_gui.py` 상단의 `APP_VERSION` 이 단일 출처입니다 (사이드바 하단에 `v1.0.0` 형태로 표시).

## 자동 업데이트

GitHub Releases(`everside76/RestaurantFinder`)를 이용해 실행 파일을 자동으로 최신 버전으로 바꿉니다.

- 프로그램 시작 약 2초 뒤 백그라운드에서 최신 릴리스를 확인합니다 (실패해도 조용히 무시, 조회 기능에 영향 없음).
- 새 버전이 있으면 exe 옆에 `RestaurantFinder.new.exe`로 내려받고(하단 상태줄에 진행률 % 표시),
  파일 크기와 SHA-256 해시를 검증합니다.
- 다운로드가 끝나면 **[예] 지금 재시작해서 적용** / **[아니오] 프로그램을 닫을 때 자동 적용** 중에서 고릅니다.
  조회가 진행 중이면 결과를 잃지 않도록 재시작 대신 '닫을 때 적용'으로 전환합니다.
- 교체 방식: 실행 중인 `RestaurantFinder.exe` → `RestaurantFinder.old.exe` 로 이름 변경 후 새 파일을 원래 이름으로 이동.
  중간에 실패하면 원래 파일로 되돌립니다. `.old.exe`·남은 `.new.exe`/`.part`는 다음 시작 시 자동 정리됩니다.
- 업데이트는 **exe 파일만** 교체합니다. `config.json`(암호화된 인증키)은 읽지도 쓰지도 않으므로 그대로 유지됩니다.
- 사이드바 하단 **[🔄 업데이트 확인]** 버튼으로 수동 확인 가능 (최신이면 "최신 버전입니다" 안내).
- 소스(`python restaurant_api_gui.py`)로 실행할 때는 새 버전 안내만 하고 내려받거나 교체하지 않습니다.

## 릴리스 방법 (release.ps1)

저장소 폴더에서 PowerShell로 실행합니다. 사전 조건: `git`, `python`(PyInstaller 설치), `gh`(GitHub CLI, `gh auth login` 완료).

```powershell
.\release.ps1 -Version 1.2.3
.\release.ps1 -Version 1.2.3 -Notes "검색 속도 개선, 오류 메시지 정리"
```

> 실행 정책 때문에 막히면: `powershell -ExecutionPolicy Bypass -File .\release.ps1 -Version 1.2.3`

스크립트가 수행하는 순서 (어느 단계든 실패하면 즉시 중단, 커밋 전 실패 시 버전 변경 자동 복구):

1. 작업 트리가 깨끗한지, 같은 태그가 이미 없는지, `gh` 로그인 상태 확인
2. `restaurant_api_gui.py`의 `APP_VERSION` 을 지정한 버전으로 교체
3. `python -m py_compile` 로 문법 확인
4. PyInstaller로 `dist\RestaurantFinder.exe` 빌드 → `build\`, `RestaurantFinder.spec` 정리
5. **인증키 유출 검사** (`python -I check_no_secrets.py`) — 실패 시 커밋 전에 중단
6. `git commit -a` ("릴리스 vX.Y.Z") → `git tag vX.Y.Z` → `git push` + `git push origin vX.Y.Z`
7. **인증키 유출 재검사** — 실패 시 릴리스를 만들지 않음
8. `gh release create vX.Y.Z dist\RestaurantFinder.exe` (노트를 생략하면 기본 문구 사용)

릴리스가 게시되면 기존 사용자의 프로그램이 다음 실행 시 자동으로 새 버전을 받습니다.

### 릴리스 전 인증키 유출 자동 검사 (check_no_secrets.py)

저장소가 공개(public)이므로 인증키가 절대 올라가지 않도록 릴리스 과정에서 자동으로 두 번 검사합니다.
수동 실행: `python -I check_no_secrets.py` (통과 시 종료 코드 0, 발견 시 1)

- `dist\config.json`, 루트 `config.json`에 저장된 인증키를 이 PC의 DPAPI로 복호화해 실제 키를 얻은 뒤,
- git이 추적하는(또는 추적될 수 있는) 모든 파일, `git diff --cached`/`git diff` 변경분,
  빌드된 `dist\RestaurantFinder.exe`(UTF-8·UTF-16 둘 다)에서
  키 원문, 키 앞 12자·뒤 12자, 암호문 앞 40자가 있는지 찾습니다.
- `config.json`, `*.new.exe`, `*.old.exe`, `*.part` 가 git에 추적되고 있으면 실패합니다 (`.gitignore`에도 등록됨).
- 결과에는 키 값이 절대 출력되지 않고 "식품안전나라 키: 0건"처럼 건수와 파일 위치만 표시됩니다.

## 데이터셋

| 서비스 ID | 이름 | 내용 |
|---|---|---|
| MOIS_I1 (기본) | 일반음식점 (행정안전부) | 공공데이터포털 `general_restaurants` API, 별도 인증키(Decoding 키) |
| I1200 | 식품접객업정보 | 전국 음식점 인허가 약 313만 건 (업소명, 대표자명, 주소, 업종, 전화, 인허가/폐업일자, 좌표) |

## 주요 기능

- **다중 조건 검색** (최대 3개 AND): 예) 소재지 주소=서울특별시 강남구 + 업소명=김밥
- **조건 일치 결과만 표시**: 서버가 다중 조건을 정확히 AND 처리하지 않는 경우가 있어,
  프로그램이 응답을 재검증하여 모든 조건이 포함된 행만 표에 표시 (KPI 카드에서 수신/일치 건수 확인)
- **영업 중인 업소만 표시** 옵션 (폐업일자 있는 업소 제외, 폐업 업소는 회색 표시)
- 행 더블클릭 → 전체 필드 상세 보기 / CSV(엑셀 호환)·JSON 저장
- 조회 중 프로그레스바 + 경과 시간 표시
- 사이드바 [검색 방법 안내] 버튼에 상세 사용법 내장
- GitHub 릴리스 기반 자동 업데이트 (아래 참고)

## 인증키

- 식품안전나라 회원가입 → 마이페이지 → 오픈API 인증키 발급 (데이터셋별 활용 신청 필요)
- 헤더의 인증키 칸에 입력 후 [저장(암호화)] 체크 → **Windows DPAPI로 암호화**되어
  exe 옆 `config.json`에 저장 (같은 PC·같은 Windows 계정에서만 복호화 가능)
- API는 반드시 **HTTPS**로 호출됨 (정식 키는 HTTP 호출 시 오류)
- `sample` 키는 테스트용: 5건 제한, 검색 조건 미적용

## 참고

- 한 번에 최대 1,000건 조회. 조건 일치 결과를 더 찾으려면 조회 범위를 나눠 반복 조회
- 대표자명은 개인정보 마스킹(`김*수`, `이**`) 형태 그대로 검색해야 함
- `restaurant_api.py`: 초기 버전 CLI (참고용)
