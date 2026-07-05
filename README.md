# HWP2Excel

한글(HWP) 파일의 표를 엑셀(xlsx) 시트로 변환하는 Windows GUI 프로그램. 표 하나가 시트 하나(`표001`, `표002`, …)가 된다.

## 사용법

```bash
pip install -r requirements.txt
python hwp2excel.py
```

1. HWP 파일 선택 (저장 위치는 자동으로 `<원본이름>_표변환.xlsx`)
2. 파싱 방식 선택
   - **한글 프로그램 연동 (권장)** — 한글(HWP) 프로그램이 설치된 경우. COM 자동화로 표를 읽는다.
   - **직접 파싱 (한글 미설치)** — HWP 바이너리를 직접 파싱한다. 셀 병합/복잡한 서식은 단순화될 수 있다.
3. 변환 시작

## exe 빌드

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name HWP2Excel hwp2excel.py
```

## 이력

이 소스는 원본 `.py`가 유실된 상태에서 배포본 `HWP2Excel.exe`(PyInstaller, Python 3.14)의 바이트코드를 디스어셈블하여 복원한 것이다. 최초 커밋이 복원된 원본이며, 이후 커밋에서 버그를 수정한다.
