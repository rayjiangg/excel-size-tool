import os
import re
import traceback
from copy import copy
from datetime import datetime
from dataclasses import dataclass
from typing import List, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


APP_TITLE = "Excel 款色尺码拼接工具"
HEADER_SCAN_MAX_ROW = 20
PREVIEW_MAX_ROWS = 50
PREVIEW_SIZE_GROUPS_PER_PAGE = 5
FORMULA_HEADER_FILL = "FFF200"


@dataclass
class SizeColumn:
    col_idx: int
    header_text: str
    size_no: str


@dataclass
class DetectionResult:
    sheet_name: str
    header_row: int
    color_col_idx: int
    color_header: str
    size_columns: List[SizeColumn]


class ExcelTransformError(Exception):
    pass


class WorkbookAnalyzer:
    @staticmethod
    def normalize_text(value) -> str:
        if value is None:
            return ""
        return str(value).strip().replace("\n", " ").replace("\r", " ")

    @staticmethod
    def extract_size_no(value) -> Optional[str]:
        if value is None:
            return None
        if isinstance(value, int):
            return str(value) if value > 0 else None
        if isinstance(value, float) and value.is_integer() and value > 0:
            return str(int(value))

        text = WorkbookAnalyzer.normalize_text(value)
        if not text:
            return None

        compact = text.replace(" ", "")
        m = re.fullmatch(r"尺码(\d+)", compact, flags=re.IGNORECASE)
        if m:
            return m.group(1)
        m = re.fullmatch(r"size(\d+)", compact, flags=re.IGNORECASE)
        if m:
            return m.group(1)
        if re.fullmatch(r"\d+", compact):
            return compact
        return None

    @staticmethod
    def is_color_header(value) -> bool:
        return WorkbookAnalyzer.normalize_text(value) == "款色"

    @staticmethod
    def detect_header_row(ws) -> Optional[int]:
        best_row = None
        best_count = -1
        for row_idx in range(1, min(ws.max_row, HEADER_SCAN_MAX_ROW) + 1):
            color_found = False
            size_count = 0
            for col_idx in range(1, ws.max_column + 1):
                value = ws.cell(row=row_idx, column=col_idx).value
                if WorkbookAnalyzer.is_color_header(value):
                    color_found = True
                elif WorkbookAnalyzer.extract_size_no(value):
                    size_count += 1
            if color_found and size_count > best_count:
                best_row = row_idx
                best_count = size_count
        return best_row

    @staticmethod
    def detect_columns(ws, header_row: int) -> DetectionResult:
        color_col_idx = None
        size_columns: List[SizeColumn] = []

        for col_idx in range(1, ws.max_column + 1):
            value = ws.cell(row=header_row, column=col_idx).value
            if WorkbookAnalyzer.is_color_header(value):
                color_col_idx = col_idx
                continue

            size_no = WorkbookAnalyzer.extract_size_no(value)
            if size_no is not None:
                size_columns.append(
                    SizeColumn(col_idx=col_idx, header_text=f"尺码{size_no}", size_no=size_no)
                )

        if color_col_idx is None:
            raise ExcelTransformError(f"在第 {header_row} 行没有找到“款色”表头。")
        if not size_columns:
            raise ExcelTransformError("在该表头行没有找到尺码列（支持：尺码1 / 1 / size1）。")

        size_columns.sort(key=lambda x: (int(x.size_no), x.col_idx))
        return DetectionResult(
            sheet_name=ws.title,
            header_row=header_row,
            color_col_idx=color_col_idx,
            color_header="款色",
            size_columns=size_columns,
        )

    @staticmethod
    def collect_valid_rows(ws, detect: DetectionResult, max_rows: Optional[int] = None):
        rows = []
        for row_idx in range(detect.header_row + 1, ws.max_row + 1):
            color_value = ws.cell(row=row_idx, column=detect.color_col_idx).value
            color_text = WorkbookAnalyzer.normalize_text(color_value)
            if not color_text:
                continue

            item = {
                "row_idx": row_idx,
                "color": color_text,
                "sizes": [],
            }
            for sc in detect.size_columns:
                item["sizes"].append(
                    {
                        "size_no": sc.size_no,
                        "concat": f"{color_text}{sc.size_no}",
                        "raw": ws.cell(row=row_idx, column=sc.col_idx).value,
                    }
                )
            rows.append(item)
            if max_rows is not None and len(rows) >= max_rows:
                break
        return rows


class WorkbookTransformer:
    @staticmethod
    def safe_sheet_name(wb, base_name: str) -> str:
        name = (base_name or "模型_处理结果")[:31]
        i = 1
        while name in wb.sheetnames:
            suffix = f"_{i}"
            name = f"{base_name[:31-len(suffix)]}{suffix}"
            i += 1
        return name

    @staticmethod
    def clone_style(dst, src, *, fill_override=None, bold_override=None, clear_fill=False):
        dst._style = copy(src._style)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)
        dst.border = copy(src.border)

        if clear_fill:
            dst.fill = PatternFill(fill_type=None)
        elif fill_override is not None:
            dst.fill = copy(fill_override)
        else:
            dst.fill = copy(src.fill)

        if bold_override is not None:
            src_font = src.font
            dst.font = Font(
                name=src_font.name,
                sz=src_font.sz,
                b=bold_override,
                i=src_font.i,
                charset=src_font.charset,
                u=src_font.u,
                strike=src_font.strike,
                color=src_font.color.type == 'rgb' and src_font.color.rgb or None,
                vertAlign=src_font.vertAlign,
                outline=src_font.outline,
                shadow=src_font.shadow,
                condense=src_font.condense,
                extend=src_font.extend,
                family=src_font.family,
                scheme=src_font.scheme,
            )
        else:
            dst.font = copy(src.font)

    @staticmethod
    def set_column_widths(result_ws, source_ws, detect: DetectionResult):
        # A列沿用款色列宽；公式列沿用款色列宽；数字列沿用原尺码列宽
        color_width = source_ws.column_dimensions[get_column_letter(detect.color_col_idx)].width or 12
        result_ws.column_dimensions["A"].width = color_width

        current_col = 2
        for sc in detect.size_columns:
            formula_letter = get_column_letter(current_col)
            raw_letter = get_column_letter(current_col + 1)
            size_width = source_ws.column_dimensions[get_column_letter(sc.col_idx)].width or 10
            result_ws.column_dimensions[formula_letter].width = color_width
            result_ws.column_dimensions[raw_letter].width = size_width
            current_col += 2

    @staticmethod
    def copy_row_heights(result_ws, source_ws, detect: DetectionResult, data_rows_written: int):
        header_height = source_ws.row_dimensions[detect.header_row].height
        if header_height:
            result_ws.row_dimensions[1].height = header_height
            result_ws.row_dimensions[2].height = header_height
        for i in range(data_rows_written):
            src_row = detect.header_row + 1 + i
            dst_row = 3 + i
            h = source_ws.row_dimensions[src_row].height
            if h:
                result_ws.row_dimensions[dst_row].height = h

    @staticmethod
    def create_template_result_sheet(wb, source_ws, detect: DetectionResult, log_cb):
        result_name = WorkbookTransformer.safe_sheet_name(wb, f"{detect.sheet_name}_处理结果")
        ws = wb.create_sheet(result_name)

        source_color_header = source_ws.cell(detect.header_row, detect.color_col_idx)
        sample_size_header = source_ws.cell(detect.header_row, detect.size_columns[0].col_idx)
        sample_data_text = source_ws.cell(detect.header_row + 1, detect.color_col_idx)
        yellow_fill = PatternFill(fill_type="solid", fgColor=FORMULA_HEADER_FILL)

        # 第1行：只在数字列的正上方写“尺码1/2/3...”，不合并，不跨列
        top_blank_cols = [1]
        current_col = 2
        for sc in detect.size_columns:
            top_blank_cols.append(current_col)
            current_col += 2

        max_col = 1 + len(detect.size_columns) * 2
        for col in range(1, max_col + 1):
            cell = ws.cell(1, col)
            WorkbookTransformer.clone_style(cell, source_color_header, clear_fill=True, bold_override=False)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.value = ""

        # 第2行
        a2 = ws.cell(2, 1, value="款色")
        WorkbookTransformer.clone_style(a2, source_color_header)

        current_col = 2
        for sc in detect.size_columns:
            formula_header = ws.cell(2, current_col, value="公式变成")
            WorkbookTransformer.clone_style(formula_header, source_color_header, fill_override=yellow_fill)

            raw_header = ws.cell(2, current_col + 1, value=sc.size_no)
            WorkbookTransformer.clone_style(raw_header, source_ws.cell(detect.header_row, sc.col_idx))

            top_size = ws.cell(1, current_col + 1, value=f"尺码{sc.size_no}")
            WorkbookTransformer.clone_style(top_size, source_ws.cell(detect.header_row, sc.col_idx), clear_fill=True, bold_override=False)
            top_size.alignment = Alignment(horizontal="center", vertical="center")
            current_col += 2

        # 第3行开始数据
        target_row = 3
        written = 0
        valid_rows = WorkbookAnalyzer.collect_valid_rows(source_ws, detect)
        for row in valid_rows:
            src_row = row["row_idx"]
            color_src = source_ws.cell(src_row, detect.color_col_idx)
            dst_color = ws.cell(target_row, 1, value=row["color"])
            WorkbookTransformer.clone_style(dst_color, color_src)

            current_col = 2
            for idx, sc in enumerate(detect.size_columns):
                raw_src = source_ws.cell(src_row, sc.col_idx)

                dst_formula = ws.cell(target_row, current_col, value=f"{row['color']}{sc.size_no}")
                # 公式变成列整体沿用款色列样式，保证字体/边框一致
                WorkbookTransformer.clone_style(dst_formula, color_src)

                dst_raw = ws.cell(target_row, current_col + 1, value=raw_src.value)
                WorkbookTransformer.clone_style(dst_raw, raw_src)
                current_col += 2

            target_row += 1
            written += 1

        WorkbookTransformer.set_column_widths(ws, source_ws, detect)
        WorkbookTransformer.copy_row_heights(ws, source_ws, detect, written)
        ws.freeze_panes = "A3"
        ws.sheet_view.showGridLines = True
        log_cb(f"已生成结果工作表：{result_name}，共写入 {written} 行数据。")
        return result_name

    @staticmethod
    def process_file(input_path: str, output_path: str, sheet_name: str, header_row: int, log_cb):
        wb = load_workbook(input_path)
        if sheet_name not in wb.sheetnames:
            raise ExcelTransformError(f"工作表不存在：{sheet_name}")
        source_ws = wb[sheet_name]
        detect = WorkbookAnalyzer.detect_columns(source_ws, header_row)

        log_cb(f"开始处理工作表：{sheet_name}")
        log_cb(
            f"检测到 款色 列：{get_column_letter(detect.color_col_idx)}；"
            f"尺码列：{', '.join([get_column_letter(s.col_idx) for s in detect.size_columns])}"
        )

        result_name = WorkbookTransformer.create_template_result_sheet(wb, source_ws, detect, log_cb)
        out_dir = os.path.dirname(output_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        wb.save(output_path)
        log_cb(f"文件已保存到：{output_path}")
        return result_name


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1420x900")
        self.minsize(1220, 780)

        self.input_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.header_row_var = tk.IntVar(value=1)
        self.confirm_detect_var = tk.BooleanVar(value=False)
        self.confirm_preview_var = tk.BooleanVar(value=False)

        self.wb_preview = None
        self.detect_result: Optional[DetectionResult] = None
        self.preview_rows = []
        self.preview_page = 0
        self.preview_group_pages = 1

        self._build_ui()

    def log(self, message: str):
        now = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{now}] {message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _build_ui(self):
        top = ttk.Frame(self, padding=12)
        top.pack(fill="x")

        file_frame = ttk.LabelFrame(top, text="文件设置", padding=10)
        file_frame.pack(fill="x")

        ttk.Label(file_frame, text="输入文件：").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(file_frame, textvariable=self.input_path_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(file_frame, text="选择文件", command=self.choose_input_file).grid(row=0, column=2, padx=6)
        ttk.Button(file_frame, text="加载工作簿", command=self.load_workbook_preview).grid(row=0, column=3, padx=6)

        ttk.Label(file_frame, text="输出文件：").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(file_frame, textvariable=self.output_path_var).grid(row=1, column=1, sticky="ew", padx=6)
        ttk.Button(file_frame, text="选择保存位置", command=self.choose_output_file).grid(row=1, column=2, padx=6)
        file_frame.columnconfigure(1, weight=1)

        config_frame = ttk.LabelFrame(top, text="识别与处理设置", padding=10)
        config_frame.pack(fill="x", pady=(10, 0))

        ttk.Label(config_frame, text="工作表：").grid(row=0, column=0, sticky="w", pady=6)
        self.sheet_combo = ttk.Combobox(config_frame, textvariable=self.sheet_var, state="readonly", width=24)
        self.sheet_combo.grid(row=0, column=1, sticky="w", padx=6)

        ttk.Label(config_frame, text="表头行：").grid(row=0, column=2, sticky="w")
        ttk.Spinbox(config_frame, from_=1, to=9999, textvariable=self.header_row_var, width=8).grid(row=0, column=3, sticky="w", padx=6)

        ttk.Label(config_frame, text="输出格式：按原表风格复制，插入“公式变成”列").grid(row=0, column=4, columnspan=2, sticky="w", padx=6)

        btn_frame = ttk.Frame(config_frame)
        btn_frame.grid(row=2, column=0, columnspan=6, sticky="w")
        ttk.Button(btn_frame, text="自动识别表头", command=self.auto_detect).pack(side="left", padx=(0, 6))
        ttk.Button(btn_frame, text="生成预览", command=self.generate_preview).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="开始处理并保存", command=self.process_and_save).pack(side="left", padx=6)

        ttk.Checkbutton(config_frame, text="我已核对识别结果", variable=self.confirm_detect_var).grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Checkbutton(config_frame, text="我已核对预览结果", variable=self.confirm_preview_var).grid(row=3, column=2, columnspan=2, sticky="w", pady=(8, 0))

        middle = ttk.Panedwindow(self, orient="vertical")
        middle.pack(fill="both", expand=True, padx=12, pady=12)

        preview_frame = ttk.LabelFrame(middle, text="识别结果与预览", padding=10)
        middle.add(preview_frame, weight=4)

        self.detect_info = tk.Text(preview_frame, height=5, wrap="word")
        self.detect_info.pack(fill="x")

        note_label = ttk.Label(
            preview_frame,
            text="预览区使用轻量表格展示，导出时再按最终 Excel 格式生成，避免滚动卡顿。",
            foreground="#555555",
        )
        note_label.pack(anchor="w", pady=(6, 4))

        self.preview_notebook = ttk.Notebook(preview_frame)
        self.preview_notebook.pack(fill="both", expand=True)

        # 源表预览
        source_tab = ttk.Frame(self.preview_notebook)
        self.preview_notebook.add(source_tab, text="源表内容预览")
        self.source_tip = ttk.Label(source_tab, text="展示原 Excel 的关键列内容。")
        self.source_tip.pack(anchor="w", padx=6, pady=(6, 2))
        source_frame = ttk.Frame(source_tab)
        source_frame.pack(fill="both", expand=True, padx=6, pady=6)
        self.source_tree = ttk.Treeview(source_frame, show="headings")
        self.source_tree.pack(side="left", fill="both", expand=True)
        sy = ttk.Scrollbar(source_frame, orient="vertical", command=self.source_tree.yview)
        sx = ttk.Scrollbar(source_frame, orient="horizontal", command=self.source_tree.xview)
        self.source_tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        sy.pack(side="right", fill="y")
        sx.pack(side="bottom", fill="x")

        # 结果预览
        result_tab = ttk.Frame(self.preview_notebook)
        self.preview_notebook.add(result_tab, text="结果模板预览")
        top_bar = ttk.Frame(result_tab)
        top_bar.pack(fill="x", padx=6, pady=(6, 2))
        self.preview_page_label = ttk.Label(top_bar, text="预览页：1/1")
        self.preview_page_label.pack(side="left")
        ttk.Button(top_bar, text="上一页", command=self.prev_preview_page).pack(side="right", padx=(6, 0))
        ttk.Button(top_bar, text="下一页", command=self.next_preview_page).pack(side="right")

        self.row1_label = ttk.Label(result_tab, text="第1行：")
        self.row1_label.pack(anchor="w", padx=6)
        self.row2_label = ttk.Label(result_tab, text="第2行：")
        self.row2_label.pack(anchor="w", padx=6)
        self.summary_label = ttk.Label(result_tab, text="")
        self.summary_label.pack(anchor="w", padx=6, pady=(0, 4))

        result_frame = ttk.Frame(result_tab)
        result_frame.pack(fill="both", expand=True, padx=6, pady=6)
        self.result_tree = ttk.Treeview(result_frame, show="headings")
        self.result_tree.pack(side="left", fill="both", expand=True)
        ry = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_tree.yview)
        rx = ttk.Scrollbar(result_frame, orient="horizontal", command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=ry.set, xscrollcommand=rx.set)
        ry.pack(side="right", fill="y")
        rx.pack(side="bottom", fill="x")

        log_frame = ttk.LabelFrame(middle, text="日志", padding=10)
        middle.add(log_frame, weight=2)
        self.log_text = tk.Text(log_frame, height=12, state="disabled")
        self.log_text.pack(fill="both", expand=True)

        self.log("工具已启动。建议流程：加载工作簿 -> 自动识别 -> 生成预览 -> 勾选两项确认 -> 开始处理。")

    def choose_input_file(self):
        path = filedialog.askopenfilename(title="选择 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xltx *.xltm")])
        if not path:
            return
        self.input_path_var.set(path)
        if not self.output_path_var.get().strip():
            base, ext = os.path.splitext(path)
            self.output_path_var.set(f"{base}_处理结果{ext or '.xlsx'}")
        self.log(f"已选择输入文件：{path}")

    def choose_output_file(self):
        path = filedialog.asksaveasfilename(title="选择保存位置", defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if not path:
            return
        self.output_path_var.set(path)
        self.log(f"已选择输出文件：{path}")

    def load_workbook_preview(self):
        input_path = self.input_path_var.get().strip()
        if not input_path:
            messagebox.showwarning(APP_TITLE, "请先选择输入文件。")
            return
        try:
            self.wb_preview = load_workbook(input_path, data_only=False)
            sheets = self.wb_preview.sheetnames
            self.sheet_combo["values"] = sheets
            if sheets:
                self.sheet_var.set(sheets[0])
            self.confirm_detect_var.set(False)
            self.confirm_preview_var.set(False)
            self.detect_result = None
            self.preview_rows = []
            self.preview_page = 0
            self.clear_preview()
            self.log(f"工作簿加载成功，共 {len(sheets)} 个工作表：{', '.join(sheets)}")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"加载工作簿失败：{e}")
            self.log(f"加载工作簿失败：{e}")

    def get_current_ws(self):
        if self.wb_preview is None:
            raise ExcelTransformError("请先加载工作簿。")
        sheet_name = self.sheet_var.get().strip()
        if not sheet_name:
            raise ExcelTransformError("请选择工作表。")
        if sheet_name not in self.wb_preview.sheetnames:
            raise ExcelTransformError(f"工作表不存在：{sheet_name}")
        return self.wb_preview[sheet_name]

    def auto_detect(self):
        try:
            ws = self.get_current_ws()
            header_row = WorkbookAnalyzer.detect_header_row(ws)
            if not header_row:
                raise ExcelTransformError("自动识别失败：前 20 行内没有同时找到“款色”和尺码列。")
            self.header_row_var.set(header_row)
            self.detect_result = WorkbookAnalyzer.detect_columns(ws, header_row)
            self.show_detect_info(self.detect_result)
            self.confirm_detect_var.set(False)
            self.confirm_preview_var.set(False)
            self.log(f"自动识别完成：工作表={ws.title}，表头行={header_row}")
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))
            self.log(f"自动识别失败：{e}")

    def show_detect_info(self, detect: DetectionResult):
        size_desc = "，".join([f"{get_column_letter(s.col_idx)}=尺码{s.size_no}" for s in detect.size_columns])
        text = (
            f"工作表：{detect.sheet_name}\n"
            f"表头行：{detect.header_row}\n"
            f"款色列：{get_column_letter(detect.color_col_idx)}（款色）\n"
            f"识别到 {len(detect.size_columns)} 个尺码列：{size_desc}\n"
            f"输出说明：保留原表字体与单元格风格；在每个尺码列前插入一列“公式变成”；第1行仅在数字列正上方显示“尺码X”。"
        )
        self.detect_info.delete("1.0", "end")
        self.detect_info.insert("1.0", text)

    def clear_preview(self):
        for tree in [self.source_tree, self.result_tree]:
            tree["columns"] = ()
            for item in tree.get_children():
                tree.delete(item)
        self.row1_label.configure(text="第1行：")
        self.row2_label.configure(text="第2行：")
        self.summary_label.configure(text="")
        self.preview_page_label.configure(text="预览页：1/1")

    def generate_preview(self):
        try:
            ws = self.get_current_ws()
            self.detect_result = WorkbookAnalyzer.detect_columns(ws, int(self.header_row_var.get()))
            self.show_detect_info(self.detect_result)
            self.preview_rows = WorkbookAnalyzer.collect_valid_rows(ws, self.detect_result, PREVIEW_MAX_ROWS)
            self.preview_page = 0
            self.preview_group_pages = max(1, (len(self.detect_result.size_columns) + PREVIEW_SIZE_GROUPS_PER_PAGE - 1) // PREVIEW_SIZE_GROUPS_PER_PAGE)
            self.refresh_source_preview()
            self.refresh_result_preview()
            self.confirm_preview_var.set(False)
            self.log(f"预览已生成，共展示前 {len(self.preview_rows)} 行有效数据。")
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))
            self.log(f"生成预览失败：{e}")

    def refresh_source_preview(self):
        detect = self.detect_result
        if detect is None:
            return
        cols = ["源数据行", "款色"] + [f"尺码{s.size_no}" for s in detect.size_columns]
        self.source_tree["columns"] = cols
        for item in self.source_tree.get_children():
            self.source_tree.delete(item)
        for col in cols:
            self.source_tree.heading(col, text=col)
            self.source_tree.column(col, width=130 if col != "源数据行" else 80, anchor="center")
        for row in self.preview_rows:
            values = [row["row_idx"], row["color"]] + [x["raw"] for x in row["sizes"]]
            self.source_tree.insert("", "end", values=values)

    def refresh_result_preview(self):
        detect = self.detect_result
        if detect is None:
            return
        start = self.preview_page * PREVIEW_SIZE_GROUPS_PER_PAGE
        end = min(start + PREVIEW_SIZE_GROUPS_PER_PAGE, len(detect.size_columns))
        groups = detect.size_columns[start:end]

        row1_parts = ["[空]", "[空]"]
        row2_parts = ["款色"]
        cols = ["款色"]
        for sc in groups:
            row1_parts.extend([f"[尺码{sc.size_no}]", ""])
            row2_parts.extend(["公式变成", sc.size_no])
            cols.extend(["公式变成", sc.size_no])

        self.row1_label.configure(text="第1行：  " + "  ".join([p for p in row1_parts if p != ""]))
        self.row2_label.configure(text="第2行：  " + "  ".join(row2_parts))
        self.summary_label.configure(text=f"当前显示分组：尺码{groups[0].size_no} 到 尺码{groups[-1].size_no}。导出到 Excel 时会保持原表风格并插入公式变成列。")
        self.preview_page_label.configure(text=f"预览页：{self.preview_page + 1}/{self.preview_group_pages}")

        self.result_tree["columns"] = cols
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        for col in cols:
            self.result_tree.heading(col, text=col)
            width = 150 if col in ("款色", "公式变成") else 70
            self.result_tree.column(col, width=width, anchor="center")

        for row in self.preview_rows:
            values = [row["color"]]
            for sc in groups:
                size_item = next(x for x in row["sizes"] if x["size_no"] == sc.size_no)
                values.extend([size_item["concat"], size_item["raw"]])
            self.result_tree.insert("", "end", values=values)

    def next_preview_page(self):
        if self.detect_result is None:
            return
        if self.preview_page < self.preview_group_pages - 1:
            self.preview_page += 1
            self.refresh_result_preview()

    def prev_preview_page(self):
        if self.detect_result is None:
            return
        if self.preview_page > 0:
            self.preview_page -= 1
            self.refresh_result_preview()

    def process_and_save(self):
        input_path = self.input_path_var.get().strip()
        output_path = self.output_path_var.get().strip()
        if not input_path:
            messagebox.showwarning(APP_TITLE, "请先选择输入文件。")
            return
        if not output_path:
            messagebox.showwarning(APP_TITLE, "请选择输出文件保存位置。")
            return
        if self.detect_result is None:
            messagebox.showwarning(APP_TITLE, "请先完成自动识别或生成预览。")
            return
        if not self.confirm_detect_var.get():
            messagebox.showwarning(APP_TITLE, "请先勾选“我已核对识别结果”。")
            return
        if not self.confirm_preview_var.get():
            messagebox.showwarning(APP_TITLE, "请先勾选“我已核对预览结果”。")
            return

        if os.path.abspath(input_path) == os.path.abspath(output_path):
            if not messagebox.askyesno(APP_TITLE, "输入文件和输出文件相同。继续会覆盖原文件，是否继续？"):
                return

        detect = self.detect_result
        summary = (
            f"即将处理：\n"
            f"- 文件：{os.path.basename(input_path)}\n"
            f"- 工作表：{detect.sheet_name}\n"
            f"- 表头行：{detect.header_row}\n"
            f"- 款色列：{get_column_letter(detect.color_col_idx)}\n"
            f"- 尺码列：{', '.join([f'尺码{s.size_no}' for s in detect.size_columns])}\n"
            f"- 输出格式：复制原表风格，在每个尺码列前插入“公式变成”列\n"
            f"- 保存到：{output_path}"
        )
        if not messagebox.askyesno(APP_TITLE, summary):
            self.log("用户取消了执行确认。")
            return

        try:
            result_sheet_name = WorkbookTransformer.process_file(
                input_path=input_path,
                output_path=output_path,
                sheet_name=detect.sheet_name,
                header_row=detect.header_row,
                log_cb=self.log,
            )
            messagebox.showinfo(APP_TITLE, f"处理完成。\n新增结果工作表：{result_sheet_name}\n输出文件：{output_path}")
        except Exception as e:
            self.log(f"处理失败：{e}")
            self.log(traceback.format_exc())
            messagebox.showerror(APP_TITLE, f"处理失败：{e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
