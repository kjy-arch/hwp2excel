"""
HWP 표 → Excel 변환기
한글 파일(.hwp/.hwpx)의 표를 엑셀 시트로 변환 (시트마다 표 1개)
"""

import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


def extract_tables_via_hwpx(hwpx_path):
    """
    HWPX(ZIP+XML)에서 표를 직접 파싱. 한글 설치 불필요.
    셀 병합은 좌상단 셀에 값이 들어가고 나머지는 빈 칸.
    """
    import re
    import zipfile
    import xml.etree.ElementTree as ET

    NS = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"

    tables = []
    table_index = 0

    with zipfile.ZipFile(hwpx_path) as zf:
        section_names = sorted(
            n for n in zf.namelist()
            if re.fullmatch(r"Contents/section\d+\.xml", n)
        )
        if not section_names:
            raise ValueError("올바른 HWPX 파일이 아닙니다.")

        for name in section_names:
            root = ET.fromstring(zf.read(name))
            for tbl in root.iter(f"{NS}tbl"):
                row_count = int(tbl.get("rowCnt", "0"))
                col_count = int(tbl.get("colCnt", "0"))
                if row_count <= 0 or col_count <= 0:
                    continue

                grid = [["" for _ in range(col_count)] for _ in range(row_count)]
                for tr_idx, tr in enumerate(tbl.findall(f"{NS}tr")):
                    for tc_idx, tc in enumerate(tr.findall(f"{NS}tc")):
                        addr = tc.find(f"{NS}cellAddr")
                        if addr is not None:
                            r = int(addr.get("rowAddr", tr_idx))
                            c = int(addr.get("colAddr", tc_idx))
                        else:
                            r, c = tr_idx, tc_idx
                        paras = [
                            "".join(t.text or "" for t in p.iter(f"{NS}t"))
                            for p in tc.findall(f"{NS}subList/{NS}p")
                        ]
                        text = "\n".join(paras).strip()
                        if 0 <= r < row_count and 0 <= c < col_count:
                            grid[r][c] = text

                table_index += 1
                tables.append({
                    "index": table_index,
                    "rows": grid,
                    "row_count": row_count,
                    "col_count": col_count,
                })

    return tables


def extract_tables_via_com(hwp_path):
    """한글 프로그램 COM 자동화로 표 추출 (가장 정확)"""
    import re
    import win32com.client
    import pythoncom

    def get_selected_text(hwp):
        # 선택 영역 텍스트를 스캔 방식으로 읽기 (클립보드 미사용)
        hwp.InitScan(0, 0xFF)  # 0xFF = 선택 영역만 스캔
        parts = []
        while True:
            state, text = hwp.GetText()
            if state in (0, 1):
                break
            parts.append(text)
        hwp.ReleaseScan()
        return "".join(parts).strip()

    def get_cell_addr(hwp):
        # KeyIndicator 마지막 요소가 "(A1)..." 형태의 셀 주소
        indicator = hwp.KeyIndicator()[-1]
        m = re.match(r"\(([A-Z]+)(\d+)\)", str(indicator))
        if not m:
            return None
        col = 0
        for ch in m.group(1):
            col = col * 26 + (ord(ch) - ord("A") + 1)
        return int(m.group(2)), col  # (row, col) 1-based

    pythoncom.CoInitialize()
    hwp = None
    tables = []
    try:
        hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        # 포맷 ""(자동 감지): hwp/hwpx 모두 확장자에 맞춰 연다
        if not hwp.Open(hwp_path, "", "forceopen:true"):
            raise RuntimeError(
                "한글에서 파일을 열지 못했습니다.\n"
                "파일 형식을 확인하거나 직접 파싱 방식을 사용해보세요.")

        hwp.SetPos(0, 0, 0)
        ctrl = hwp.HeadCtrl

        table_index = 0
        while ctrl is not None:
            if ctrl.CtrlID == "tbl":
                table_index += 1

                # 표 첫 셀로 진입한 뒤 셀을 차례로 순회하며 텍스트 수집
                hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                hwp.FindCtrl()
                hwp.HAction.Run("ShapeObjTableSelCell")

                cells = {}  # (row, col) -> text, 1-based
                while True:
                    addr = get_cell_addr(hwp)
                    hwp.HAction.Run("TableCellBlock")
                    text = get_selected_text(hwp)
                    if addr:
                        cells[addr] = text
                    pos = hwp.GetPos()
                    hwp.HAction.Run("TableRightCell")
                    if hwp.GetPos() == pos:
                        break  # 마지막 셀
                hwp.HAction.Run("Cancel")

                if not cells:
                    ctrl = ctrl.Next
                    continue

                row_count = max(r for r, _ in cells)
                col_count = max(c for _, c in cells)
                rows_data = [
                    [cells.get((r, c), "") for c in range(1, col_count + 1)]
                    for r in range(1, row_count + 1)
                ]

                tables.append({
                    "index": table_index,
                    "rows": rows_data,
                    "row_count": row_count,
                    "col_count": col_count,
                })
            ctrl = ctrl.Next
    finally:
        if hwp:
            try:
                hwp.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    return tables


def extract_tables_via_olefile(hwp_path):
    """
    HWP 바이너리(OLE)에서 표 레코드를 직접 파싱.
    제한: 셀 병합/복잡한 서식은 단순화될 수 있음.
    """
    import olefile
    import zlib
    import struct

    # HWP 5.0 스펙: HWPTAG_BEGIN(16) + 오프셋
    TAG_PARA_TEXT = 67    # HWPTAG_PARA_TEXT
    TAG_LIST_HEADER = 72  # HWPTAG_LIST_HEADER (셀마다 1개)
    TAG_TABLE = 77        # HWPTAG_TABLE

    def iter_records(data):
        i = 0
        while i + 4 <= len(data):
            header = struct.unpack_from("<I", data, i)[0]
            tag_id = header & 0x3FF
            level = (header >> 10) & 0x3FF
            size = (header >> 20) & 0xFFF
            i += 4
            if size == 0xFFF:
                if i + 4 > len(data):
                    return
                size = struct.unpack_from("<I", data, i)[0]
                i += 4
            payload = data[i:i + size]
            i += size
            yield tag_id, level, payload

    def read_section(raw):
        try:
            return zlib.decompress(raw, -15)
        except Exception:
            return raw

    if not olefile.isOleFile(hwp_path):
        raise ValueError("올바른 HWP 파일이 아닙니다.")

    ole = olefile.OleFileIO(hwp_path)
    tables = []
    table_index = 0

    def decode_para_text(payload):
        # PARA_TEXT는 UTF-16LE. 인라인 컨트롤 문자(코드 < 32)는
        # 확장 컨트롤(8워드) 여부에 따라 건너뛴다.
        EXTENDED = {1, 2, 3, 11, 12, 14, 15, 16, 17, 18, 21, 22, 23}
        chars = []
        i = 0
        n = len(payload) // 2
        while i < n:
            code = struct.unpack_from("<H", payload, i * 2)[0]
            if code in EXTENDED:
                i += 8  # 컨트롤 코드 + 부가정보 7워드
                continue
            if code < 32:
                if code in (10, 13):
                    chars.append("\n")
                i += 1
                continue
            chars.append(chr(code))
            i += 1
        return "".join(chars).strip()

    def emit_table(cell_map, row_count, col_count):
        # cell_map: (row, col) -> text. 병합 셀은 좌상단만 채워진다.
        nonlocal table_index
        if row_count <= 0 or col_count <= 0 or not cell_map:
            return
        table_index += 1
        tables.append({
            "index": table_index,
            "rows": [
                [cell_map.get((r, c), "") for c in range(col_count)]
                for r in range(row_count)
            ],
            "row_count": row_count,
            "col_count": col_count,
        })

    section_num = 0
    while True:
        stream_name = f"BodyText/Section{section_num}"
        if not ole.exists(stream_name):
            break
        data = read_section(ole.openstream(stream_name).read())

        # 상태 머신: TABLE 레코드가 표를 열고, 같은 레벨의 LIST_HEADER가
        # 셀 시작(페이로드 오프셋 8/10에 열/행 주소), 그보다 깊은
        # PARA_TEXT가 셀 내용. 레벨이 표보다 얕아지면 표 범위를 벗어난 것.
        in_table = False
        table_level = 0
        cur_addr = None   # 현재 열린 셀의 (row, col)
        cell_texts = []
        cell_map = {}
        row_count = 0
        col_count = 0

        def close_cell():
            nonlocal cur_addr
            if cur_addr is not None:
                cell_map[cur_addr] = "\n".join(cell_texts).strip()
                cur_addr = None
            cell_texts.clear()

        for tag_id, level, payload in iter_records(data):
            if tag_id == TAG_TABLE and (not in_table or level <= table_level):
                # 새 표 시작 — 표 레코드: 속성 UINT32 + 행 UINT16 + 열 UINT16
                if len(payload) >= 8:
                    close_cell()
                    if in_table:
                        emit_table(cell_map, row_count, col_count)
                    row_count = struct.unpack_from("<H", payload, 4)[0]
                    col_count = struct.unpack_from("<H", payload, 6)[0]
                    in_table = True
                    table_level = level
                    cell_map = {}
                continue

            if not in_table:
                continue

            if tag_id == TAG_LIST_HEADER and level == table_level:
                close_cell()
                if len(payload) >= 12:
                    col = struct.unpack_from("<H", payload, 8)[0]
                    row = struct.unpack_from("<H", payload, 10)[0]
                    if row < row_count and col < col_count:
                        cur_addr = (row, col)
            elif tag_id == TAG_PARA_TEXT and cur_addr is not None \
                    and level > table_level:
                try:
                    cell_texts.append(decode_para_text(payload))
                except Exception:
                    pass
            elif level < table_level:
                # 표 범위 밖으로 나옴 — 표 확정
                close_cell()
                emit_table(cell_map, row_count, col_count)
                in_table = False

        # 스트림 끝에서 열린 표 마무리
        if in_table:
            close_cell()
            emit_table(cell_map, row_count, col_count)

        section_num += 1

    ole.close()
    return tables


def convert_cell_value(text):
    """숫자로 보이는 텍스트는 숫자로 변환 (SUM 등 수식이 동작하도록).

    "4,991" → 4991, "18.6" → 18.6, "18.6%" → 0.186(+퍼센트 서식).
    단위가 붙은 값("6,666명")이나 복합 텍스트는 그대로 둔다.
    반환: (값, 숫자서식 또는 None)
    """
    import re

    t = text.strip()
    m = re.fullmatch(r"[+-]?\d{1,3}(?:,\d{3})*(?:\.\d+)?|[+-]?\d+(?:\.\d+)?", t)
    if m:
        num = t.replace(",", "")
        return (float(num) if "." in num else int(num)), None
    m = re.fullmatch(r"([+-]?\d{1,3}(?:,\d{3})*(?:\.\d+)?|[+-]?\d+(?:\.\d+)?)%", t)
    if m:
        num = m.group(1).replace(",", "")
        return float(num) / 100, "0.0%"
    return text, None


def tables_to_excel(tables, output_path, progress_cb=None):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    header_fill = PatternFill(
        start_color="DCE6F1", end_color="DCE6F1", fill_type="solid"
    )

    for i, tbl in enumerate(tables):
        sheet_name = f"표{tbl['index']:03d}"
        ws = wb.create_sheet(title=sheet_name)

        for r_idx, row in enumerate(tbl["rows"], start=1):
            for c_idx, cell_val in enumerate(row, start=1):
                value, num_format = convert_cell_value(cell_val)
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                if num_format:
                    cell.number_format = num_format
                cell.border = border
                cell.alignment = Alignment(
                    wrap_text=True, vertical="center"
                )
                if r_idx == 1:
                    cell.font = Font(bold=True)
                    cell.fill = header_fill

        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for c in col:
                try:
                    max_len = max(max_len, len(str(c.value or "")))
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = min(max_len + 4, 40)

        if progress_cb:
            progress_cb(i + 1, len(tables))

    wb.save(output_path)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HWP 표 → Excel 변환기")
        self.geometry("520x380")
        self.resizable(False, False)
        self.configure(bg="#f5f5f5")

        self._build_ui()

    def _build_ui(self):
        pad = dict(padx=20, pady=8)

        tk.Label(
            self, text="HWP 표 → Excel 변환기",
            font=("맑은 고딕", 16, "bold"), bg="#f5f5f5", fg="#1a3a5c",
        ).pack(pady=(24, 4))

        tk.Label(
            self, text="한글 파일의 표를 엑셀 시트별로 변환합니다",
            font=("맑은 고딕", 10), bg="#f5f5f5", fg="#666",
        ).pack(pady=(0, 16))

        frm1 = tk.Frame(self, bg="#f5f5f5")
        frm1.pack(fill="x", **pad)
        tk.Label(frm1, text="HWP 파일:", font=("맑은 고딕", 10),
                 bg="#f5f5f5", width=10, anchor="w").pack(side="left")
        self.hwp_var = tk.StringVar()
        tk.Entry(frm1, textvariable=self.hwp_var, width=36,
                 font=("맑은 고딕", 10)).pack(side="left", padx=(4, 6))
        tk.Button(frm1, text="찾아보기", command=self._browse_hwp,
                  font=("맑은 고딕", 9)).pack(side="left")

        frm2 = tk.Frame(self, bg="#f5f5f5")
        frm2.pack(fill="x", **pad)
        tk.Label(frm2, text="저장 위치:", font=("맑은 고딕", 10),
                 bg="#f5f5f5", width=10, anchor="w").pack(side="left")
        self.out_var = tk.StringVar()
        tk.Entry(frm2, textvariable=self.out_var, width=36,
                 font=("맑은 고딕", 10)).pack(side="left", padx=(4, 6))
        tk.Button(frm2, text="찾아보기", command=self._browse_out,
                  font=("맑은 고딕", 9)).pack(side="left")

        frm3 = tk.Frame(self, bg="#f5f5f5")
        frm3.pack(fill="x", **pad)
        tk.Label(frm3, text="파싱 방식:", font=("맑은 고딕", 10),
                 bg="#f5f5f5", width=10, anchor="w").pack(side="left")
        self.method_var = tk.StringVar(value="com")
        tk.Radiobutton(frm3, text="한글 프로그램 연동 (권장)",
                       variable=self.method_var, value="com",
                       bg="#f5f5f5", font=("맑은 고딕", 10)).pack(side="left")
        tk.Radiobutton(frm3, text="직접 파싱 (한글 미설치)",
                       variable=self.method_var, value="ole",
                       bg="#f5f5f5", font=("맑은 고딕", 10)).pack(side="left", padx=(12, 0))

        self.progress = ttk.Progressbar(self, length=460, mode="determinate")
        self.progress.pack(**pad)

        self.status_var = tk.StringVar(value="파일을 선택해주세요.")
        tk.Label(self, textvariable=self.status_var,
                 font=("맑은 고딕", 9), bg="#f5f5f5", fg="#555").pack()

        self.btn = tk.Button(
            self, text="변환 시작", command=self._start,
            font=("맑은 고딕", 12, "bold"),
            bg="#1a3a5c", fg="white",
            width=16, height=2, relief="flat", cursor="hand2",
        )
        self.btn.pack(pady=20)

    def _browse_hwp(self):
        path = filedialog.askopenfilename(
            title="HWP 파일 선택",
            filetypes=[("한글 파일", "*.hwp *.hwpx"), ("모든 파일", "*.*")],
        )
        if path:
            self.hwp_var.set(path)
            base = os.path.splitext(path)[0]
            self.out_var.set(base + "_표변환.xlsx")

    def _browse_out(self):
        path = filedialog.asksaveasfilename(
            title="저장 위치",
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
        )
        if path:
            self.out_var.set(path)

    def _start(self):
        hwp = self.hwp_var.get().strip()
        out = self.out_var.get().strip()

        if not hwp:
            messagebox.showwarning("알림", "HWP 파일을 선택해주세요.")
            return
        if not os.path.exists(hwp):
            messagebox.showerror("오류", "선택한 파일이 존재하지 않습니다.")
            return
        if not out:
            messagebox.showwarning("알림", "저장 위치를 지정해주세요.")
            return

        self.btn.config(state="disabled")
        self.progress["value"] = 0
        self.status_var.set("변환 중...")

        def run():
            try:
                method = self.method_var.get()
                self.status_var.set("표 추출 중...")

                ext = os.path.splitext(hwp)[1].lower()
                if ext == ".hwpx":
                    # HWPX는 ZIP+XML 포맷 — 직접 파싱이 가장 정확하고
                    # 한글 설치도 필요 없다
                    tables = extract_tables_via_hwpx(hwp)
                elif method == "com":
                    tables = extract_tables_via_com(hwp)
                else:
                    tables = extract_tables_via_olefile(hwp)

                if not tables:
                    messagebox.showwarning(
                        "알림", "표를 찾을 수 없습니다.\n파싱 방식을 바꿔 시도해보세요.")
                    return

                self.status_var.set(f"표 {len(tables)}개 발견. 엑셀 생성 중...")
                self.progress["maximum"] = len(tables)

                def on_progress(done, total):
                    self.progress["value"] = done
                    self.status_var.set(f"변환 중... {done}/{total}")
                    self.update_idletasks()

                tables_to_excel(tables, out, progress_cb=on_progress)

                messagebox.showinfo(
                    "완료",
                    f"변환 완료!\n\n표 {len(tables)}개 → {len(tables)}개 시트\n\n저장 위치:\n{out}")

                self.status_var.set(f"완료! 표 {len(tables)}개 변환됨")
            except Exception as e:
                messagebox.showerror("오류", f"변환 실패:\n{e}")
                self.status_var.set("오류 발생")
            finally:
                self.btn.config(state="normal")
                self.progress["value"] = 0

        threading.Thread(target=run, daemon=True).start()


if __name__ == "__main__":
    app = App()
    app.mainloop()
