# HWP2Excel

한글(HWP/HWPX) 파일의 표를 엑셀(xlsx) 시트로 변환하는 Windows GUI 프로그램. 표 하나가 시트 하나(`표001`, `표002`, …)가 된다.

## 사용법

```bash
pip install -r requirements.txt
python hwp2excel.py
```

1. HWP/HWPX 파일 선택 (저장 위치는 자동으로 `<원본이름>_표변환.xlsx`)
2. 파싱 방식 선택
   - **한글 프로그램 연동 (권장)** — 한글(HWP) 프로그램이 설치된 경우. COM 자동화로 표를 읽는다.
   - **직접 파싱 (한글 미설치)** — HWP 바이너리를 직접 파싱한다.
3. 변환 시작

### 포맷별 처리 방식

| 파일 | 처리 |
| --- | --- |
| `.hwpx` | ZIP+XML 직접 파싱 (파싱 방식 선택과 무관, 한글 설치 불필요, 셀 병합 주소 반영) |
| `.hwp` + 한글 연동 | COM 자동화로 셀 순회 (셀 주소 기반) |
| `.hwp` + 직접 파싱 | OLE 레코드 파싱 (HWPTAG_TABLE/LIST_HEADER/PARA_TEXT, 셀 주소 기반) |

병합 셀은 좌상단 셀에 값이 들어가고 나머지는 빈 칸이 된다.

### 알려진 제한

- 한글 연동(COM) 방식은 보안 승인 창이 뜰 수 있다("접근 허용" 선택). 배포용 보안 모듈(FilePathCheckerModule)이 레지스트리에 등록된 환경에서는 뜨지 않는다.
- 셀 안의 중첩 표는 별도 표로도 추출되고, 바깥 셀 텍스트에도 섞여 들어갈 수 있다.

## exe 빌드

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name HWP2Excel hwp2excel.py
```

## 이력

이 소스는 원본 `.py`가 유실된 상태에서 배포본 `HWP2Excel.exe`(PyInstaller, Python 3.14)의 바이트코드를 디스어셈블하여 복원한 것이다. 최초 커밋이 복원된 원본이며, 이후 커밋에서 다음 버그를 수정했다.

- `.hwpx` 미지원 (COM은 포맷 `"HWP"` 고정 + `Open()` 결과 미확인 → 조용히 "표 없음", 직접 파싱은 OLE 전용이라 실패) → hwpx ZIP+XML 파서 추가, 확장자 기반 자동 라우팅
- OLE 파서가 존재하지 않는 스트림 이름(`Section0001`, 실제는 `Section0`)을 조회해 항상 표 0개 반환
- OLE 파서 태그 오류: `HWPTAG_TABLE`은 77(76은 SHAPE_COMPONENT), 행/열 수는 페이로드 오프셋 4/6
- 셀 텍스트를 문단 수 기반으로 세던 것을 LIST_HEADER의 셀 주소(오프셋 8/10) 기반 그리드 배치로 교체 (병합 셀 대응)
- COM 셀 읽기: 존재하지 않는 `ctrl.GetCell()` API 대신 셀 순회(`ShapeObjTableSelCell`/`TableRightCell`) + `InitScan` 텍스트 스캔으로 교체
