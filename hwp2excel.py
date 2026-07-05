"""
HWP 표 → Excel 변환기
한글 파일의 표를 엑셀 시트로 변환 (시트마다 표 1개)
"""

import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


def extract_tables_via_com(hwp_path):
    """한글 프로그램 COM 자동화로 표 추출 (가장 정확)"""
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    hwp = None
    tables = []
    try:
        hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.Open(hwp_path, "HWP", "forceopen:true")

        hwp.SetPos(0, 0, 0)
        ctrl = hwp.HeadCtrl

        table_index = 0
        while ctrl is not None:
            if ctrl.CtrlID == "tbl":
                table_index += 1
                rows_data = []
                row_count = ctrl.Properties.Item("NumRows").Value
                col_count = ctrl.Properties.Item("NumCols").Value

                for r in range(row_count):
                    row = []
                    for c in range(col_count):
                        try:
                            cell = ctrl.GetCell(r, c)
                            text = cell.Text.strip() if cell else ""
                        except Exception:
                            text = ""
                        row.append(text)
                    rows_data.append(row)

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

    TAG_PARA_HEADER = 66
    TAG_TABLE = 76
    TAG_CELL_TEXT = 73

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

    section_num = 1
    while True:
        stream_name = f"BodyText/Section{section_num:04d}"
        if not ole.exists(stream_name):
            break
        raw = ole.openstream(stream_name).read()
        data = read_section(raw)

        in_table = False
        current_rows = []
        current_row = []
        expected_cols = 0
        row_count = 0
        col_count = 0

        for tag_id, level, payload in iter_records(data):
            if tag_id == TAG_TABLE:
                if len(payload) >= 16:
                    row_count = struct.unpack_from("<H", payload, 8)[0]
                    col_count = struct.unpack_from("<H", payload, 10)[0]
                in_table = True
                current_rows = []
                current_row = []
                expected_cols = col_count
            elif tag_id == 77 and in_table:
                continue
            elif tag_id == 67 and in_table:
                try:
                    text = payload.decode("utf-16-le", errors="replace").strip()
                    text = text.replace("\x00", "").strip()
                except Exception:
                    text = ""
                current_row.append(text)

                if len(current_row) >= max(expected_cols, 1):
                    current_rows.append(current_row[:])
                    current_row = []

                    if len(current_rows) >= row_count and row_count > 0:
                        table_index += 1
                        tables.append({
                            "index": table_index,
                            "rows": current_rows,
                            "row_count": row_count,
                            "col_count": col_count,
                        })
                        in_table = False

        section_num += 1

    ole.close()
    return tables


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
                cell = ws.cell(row=r_idx, column=c_idx, value=cell_val)
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

                if method == "com":
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
