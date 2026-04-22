import math
import os
import re
import traceback
from copy import copy
from datetime import datetime
from dataclasses import dataclass
from typing import List, Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


APP_TITLE = "Excel 款色尺码拼接工具"
HEADER_SCAN_MAX_ROW = 20
PREVIEW_MAX_ROWS = 12
PREVIEW_SIZE_GROUPS_PER_PAGE = 5


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
        text = str(value).strip()
        return text.replace("\n", " ").replace("\r", " ")

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
        color_header = "款色"
        size_columns: List[SizeColumn] = []

        for col_idx in range(1, ws.max_column + 1):
            value = ws.cell(row=header_row, column=col_idx).value
            text = WorkbookAnalyzer.normalize_text(value)
            if WorkbookAnalyzer.is_color_header(value):
                color_col_idx = col_idx
                color_header = text or "款色"
                continue

            size_no = WorkbookAnalyzer.extract_size_no(value)
            if size_no is not None:
                size_columns.append(
                    SizeColumn(col_idx=col_idx, header_text=text or f"尺码{size_no}", size_no=size_no)
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
            color_header=color_header,
            size_columns=size_columns,
        )

    @staticmethod
    def build_preview(ws, result: DetectionResult, max_rows: int = PREVIEW_MAX_ROWS):
        preview_rows = []
        for row_idx in range(result.header_row + 1, ws.max_row + 1):
            color_text = WorkbookAnalyzer.normalize_text(ws.cell(row=row_idx, column=result.color_col_idx).value)
            if not color_text:
                continue

            row_item = {"row_idx": row_idx, "款色": color_text, "values": []}
            for size_col in result.size_columns:
                original_value = ws.cell(row=row_idx, column=size_col.col_idx).value
                row_item["values"].append(
                    {
                        "size_label": f"尺码{size_col.size_no}",
                        "size_no": size_col.size_no,
                        "concat_value": f"{color_text}{size_col.size_no}",
                        "original_value": original_value,
                    }
                )
            preview_rows.append(row_item)
            if len(preview_rows) >= max_rows:
                break
        return preview_rows


class WorkbookTransformer:
    BLUE_FILL = PatternFill("solid", fgColor="D9E2F3")
    YELLOW_FILL = PatternFill("solid", fgColor="FFF200")
    WHITE_FILL = PatternFill("solid", fgColor="FFFFFF")
    THIN_BORDER = Border(
        left=Side(style="thin", color="B7C3D0"),
        right=Side(style="thin", color="B7C3D0"),
        top=Side(style="thin", color="B7C3D0"),
        bottom=Side(style="thin", color="B7C3D0"),
    )

    @staticmethod
    def safe_sheet_name(wb, base_name: str) -> str:
        name = (base_name or "模型_处理结果")[:31]
        n = 1
        while name in wb.sheetnames:
            suffix = f"_{n}"
            name = f"{base_name[:31-len(suffix)]}{suffix}"
            n += 1
        return name

    @staticmethod
    def apply_cell_style(cell, *, fill=None, bold=False, center=True, font_color=None):
        cell.border = copy(WorkbookTransformer.THIN_BORDER)
        cell.font = Font(bold=bold, color=font_color)
        if fill is not None:
            cell.fill = copy(fill)
        cell.alignment = Alignment(horizontal="center" if center else "left", vertical="center")

    @staticmethod
    def autosize_columns(ws):
        for col_idx in range(1, ws.max_column + 1):
            if col_idx == 1:
                width = 18
            elif col_idx % 2 == 0:
                width = 18
            else:
                width = 10
            ws.column_dimensions[get_column_letter(col_idx)].width = width

    @staticmethod
    def create_template_result_sheet(wb, source_ws, detect: DetectionResult, log_cb):
        sheet_name = WorkbookTransformer.safe_sheet_name(wb, "模型_处理结果")
        ws = wb.create_sheet(sheet_name)

        # 第1行：空 + merged 尺码标题
        ws.cell(row=1, column=1, value="")
        WorkbookTransformer.apply_cell_style(ws.cell(row=1, column=1), fill=WorkbookTransformer.WHITE_FILL)

        current_col = 2
        for sc in detect.size_columns:
            ws.merge_cells(start_row=1, start_column=current_col, end_row=1, end_column=current_col + 1)
            top_left = ws.cell(row=1, column=current_col, value=f"尺码{sc.size_no}")
            WorkbookTransformer.apply_cell_style(top_left, fill=WorkbookTransformer.WHITE_FILL, bold=False)
            right_cell = ws.cell(row=1, column=current_col + 1, value="")
            WorkbookTransformer.apply_cell_style(right_cell, fill=WorkbookTransformer.WHITE_FILL, bold=False)
            current_col += 2

        # 第2行：款色 / 公式变成 / 1 / 公式变成 / 2 ...
        ws.cell(row=2, column=1, value="款色")
        WorkbookTransformer.apply_cell_style(ws.cell(row=2, column=1), fill=WorkbookTransformer.BLUE_FILL, bold=True)

        current_col = 2
        for sc in detect.size_columns:
            ws.cell(row=2, column=current_col, value="公式变成")
            WorkbookTransformer.apply_cell_style(ws.cell(row=2, column=current_col), fill=WorkbookTransformer.YELLOW_FILL, bold=True)
            ws.cell(row=2, column=current_col + 1, value=sc.size_no)
            WorkbookTransformer.apply_cell_style(ws.cell(row=2, column=current_col + 1), fill=WorkbookTransformer.BLUE_FILL, bold=True)
            current_col += 2

        # 第3行开始：数据
        target_row = 3
        data_count = 0
        for row_idx in range(detect.header_row + 1, source_ws.max_row + 1):
            color_text = WorkbookAnalyzer.normalize_text(source_ws.cell(row=row_idx, column=detect.color_col_idx).value)
            if not color_text:
                continue

            ws.cell(row=target_row, column=1, value=color_text)
            WorkbookTransformer.apply_cell_style(ws.cell(row=target_row, column=1), fill=WorkbookTransformer.WHITE_FILL)

            current_col = 2
            for sc in detect.size_columns:
                concat_cell = ws.cell(row=target_row, column=current_col, value=f"{color_text}{sc.size_no}")
                raw_cell = ws.cell(row=target_row, column=current_col + 1, value=source_ws.cell(row=row_idx, column=sc.col_idx).value)
                WorkbookTransformer.apply_cell_style(concat_cell, fill=WorkbookTransformer.WHITE_FILL)
                WorkbookTransformer.apply_cell_style(raw_cell, fill=WorkbookTransformer.WHITE_FILL)
                current_col += 2

            data_count += 1
            target_row += 1

        ws.freeze_panes = "A3"
        ws.sheet_view.showGridLines = True
        WorkbookTransformer.autosize_columns(ws)
        log_cb(f"已生成模板结果工作表：{sheet_name}，共写入 {data_count} 行数据")
        return sheet_name

    @staticmethod
    def process_file(input_path: str, output_path: str, sheet_name: str, header_row: int, log_cb):
        wb = load_workbook(input_path)
        if sheet_name not in wb.sheetnames:
            raise ExcelTransformError(f"工作表不存在：{sheet_name}")

        source_ws = wb[sheet_name]
        detect = WorkbookAnalyzer.detect_columns(source_ws, header_row)
        log_cb(f"开始处理工作表：{sheet_name}")
        log_cb(f"检测到 款色 列：{get_column_letter(detect.color_col_idx)}；尺码列数量：{len(detect.size_columns)}")

        result_sheet_name = WorkbookTransformer.create_template_result_sheet(wb, source_ws, detect, log_cb)

        out_dir = os.path.dirname(output_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        wb.save(output_path)

        log_cb(f"文件已保存到：{output_path}")
        log_cb(f"新增工作表：{result_sheet_name}")
        return result_sheet_name


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1380x880")
        self.minsize(1180, 780)

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
        ttk.Spinbox(config_frame, from_=1, to=9999, textvariable=self.header_row_var, width=8).grid(
            row=0, column=3, sticky="w", padx=6
        )

        ttk.Label(config_frame, text="输出格式：固定为模板样式（按模型_处理结果格式）").grid(
            row=0, column=4, columnspan=2, sticky="w", padx=6
        )

        btn_frame = ttk.Frame(config_frame)
        btn_frame.grid(row=2, column=0, columnspan=6, sticky="w")
        ttk.Button(btn_frame, text="自动识别表头", command=self.auto_detect).pack(side="left", padx=(0, 6))
        ttk.Button(btn_frame, text="生成预览", command=self.generate_preview).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="开始处理并保存", command=self.process_and_save).pack(side="left", padx=6)

        ttk.Checkbutton(config_frame, text="我已核对识别结果", variable=self.confirm_detect_var).grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(8, 0)
        )
        ttk.Checkbutton(config_frame, text="我已核对预览结果", variable=self.confirm_preview_var).grid(
            row=3, column=2, columnspan=2, sticky="w", pady=(8, 0)
        )

        middle = ttk.Panedwindow(self, orient="vertical")
        middle.pack(fill="both", expand=True, padx=12, pady=12)

        preview_frame = ttk.LabelFrame(middle, text="识别结果与预览", padding=10)
        middle.add(preview_frame, weight=4)

        info_frame = ttk.Frame(preview_frame)
        info_frame.pack(fill="x")
        self.detect_info = tk.Text(info_frame, height=5, wrap="word")
        self.detect_info.pack(fill="x")

        note_label = ttk.Label(
            preview_frame,
            text="预览区改为直接显示 Excel 内容和结果模板，不再用一整行长文本横向展示。",
            foreground="#444444",
        )
        note_label.pack(anchor="w", pady=(6, 4))

        self.preview_notebook = ttk.Notebook(preview_frame)
        self.preview_notebook.pack(fill="both", expand=True)

        self.source_preview_tab = ttk.Frame(self.preview_notebook)
        self.preview_notebook.add(self.source_preview_tab, text="源表内容预览")

        source_top = ttk.Frame(self.source_preview_tab)
        source_top.pack(fill="x", padx=6, pady=(6, 2))
        self.source_preview_tip = ttk.Label(source_top, text="这里直接展示 Excel 源表的关键列内容。")
        self.source_preview_tip.pack(anchor="w")

        source_tree_frame = ttk.Frame(self.source_preview_tab)
        source_tree_frame.pack(fill="both", expand=True, padx=6, pady=6)
        self.source_tree = ttk.Treeview(source_tree_frame, show="headings")
        self.source_tree.pack(side="left", fill="both", expand=True)
        source_y = ttk.Scrollbar(source_tree_frame, orient="vertical", command=self.source_tree.yview)
        source_x = ttk.Scrollbar(source_tree_frame, orient="horizontal", command=self.source_tree.xview)
        source_y.pack(side="right", fill="y")
        source_x.pack(side="bottom", fill="x")
        self.source_tree.configure(yscrollcommand=source_y.set, xscrollcommand=source_x.set)

        self.result_preview_tab = ttk.Frame(self.preview_notebook)
        self.preview_notebook.add(self.result_preview_tab, text="结果模板预览")

        result_top = ttk.Frame(self.result_preview_tab)
        result_top.pack(fill="x", padx=6, pady=(6, 2))
        self.preview_page_label = ttk.Label(result_top, text="预览页：1/1")
        self.preview_page_label.pack(side="left")
        ttk.Button(result_top, text="上一页", command=self.prev_preview_page).pack(side="right", padx=(6, 0))
        ttk.Button(result_top, text="下一页", command=self.next_preview_page).pack(side="right")

        self.result_header_row1 = ttk.Label(self.result_preview_tab, text="第1行：", foreground="#444444")
        self.result_header_row1.pack(anchor="w", padx=6)
        self.result_header_row2 = ttk.Label(self.result_preview_tab, text="第2行：", foreground="#444444")
        self.result_header_row2.pack(anchor="w", padx=6, pady=(0, 4))
        self.result_header_summary = ttk.Label(self.result_preview_tab, text="", foreground="#444444")
        self.result_header_summary.pack(anchor="w", padx=6, pady=(0, 4))

        result_tree_frame = ttk.Frame(self.result_preview_tab)
        result_tree_frame.pack(fill="both", expand=True, padx=6, pady=6)
        self.result_tree = ttk.Treeview(result_tree_frame, show="headings")
        self.result_tree.pack(side="left", fill="both", expand=True)
        result_y = ttk.Scrollbar(result_tree_frame, orient="vertical", command=self.result_tree.yview)
        result_x = ttk.Scrollbar(result_tree_frame, orient="horizontal", command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=result_y.set, xscrollcommand=result_x.set)
        result_y.pack(side="right", fill="y")
        result_x.pack(side="bottom", fill="x")

        log_frame = ttk.LabelFrame(middle, text="日志", padding=10)
        middle.add(log_frame, weight=2)
        self.log_text = tk.Text(log_frame, height=12, state="disabled")
        self.log_text.pack(fill="both", expand=True)

        self.log("工具已启动。建议流程：加载工作簿 -> 自动识别 -> 生成预览 -> 勾选两项确认 -> 开始处理。")

    def choose_input_file(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xltx *.xltm")],
        )
        if not path:
            return
        self.input_path_var.set(path)
        if not self.output_path_var.get().strip():
            base, ext = os.path.splitext(path)
            self.output_path_var.set(f"{base}_处理结果{ext or '.xlsx'}")
        self.log(f"已选择输入文件：{path}")

    def choose_output_file(self):
        path = filedialog.asksaveasfilename(
            title="选择保存位置",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
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
        color_col = get_column_letter(detect.color_col_idx)
        size_desc = "，".join([f"{get_column_letter(s.col_idx)}=尺码{s.size_no}" for s in detect.size_columns])
        text = (
            f"工作表：{detect.sheet_name}\n"
            f"表头行：{detect.header_row}\n"
            f"款色列：{color_col}（{detect.color_header}）\n"
            f"识别到 {len(detect.size_columns)} 个尺码列：{size_desc}\n"
            f"输出说明：第1行为空+尺码标题；第2行为 款色 / 公式变成 / 尺码号；第3行开始为数据。"
        )
        self.detect_info.delete("1.0", "end")
        self.detect_info.insert("1.0", text)

    def clear_preview(self):
        self.detect_info.delete("1.0", "end")
        for tree in (getattr(self, "source_tree", None), getattr(self, "result_tree", None)):
            if tree is not None:
                tree["columns"] = ()
                for item in tree.get_children():
                    tree.delete(item)
        self.preview_page_label.configure(text="预览页：1/1")
        self.result_header_row1.configure(text="第1行：")
        self.result_header_row2.configure(text="第2行：")
        self.result_header_summary.configure(text="")

    def build_source_preview(self, detect: DetectionResult):
        columns = ["源数据行", "款色"] + [f"尺码{s.size_no}" for s in detect.size_columns]
        self.source_tree["columns"] = columns
        for item in self.source_tree.get_children():
            self.source_tree.delete(item)
        for col in columns:
            self.source_tree.heading(col, text=col)
            self.source_tree.column(col, width=120 if col != "款色" else 160, anchor="center")

        for row in self.preview_rows:
            values = [row["row_idx"], row["款色"]] + [v["original_value"] for v in row["values"]]
            self.source_tree.insert("", "end", values=values)

    def build_result_preview(self, detect: DetectionResult):
        size_groups = detect.size_columns
        self.preview_group_pages = max(1, math.ceil(len(size_groups) / PREVIEW_SIZE_GROUPS_PER_PAGE))
        self.preview_page = min(self.preview_page, self.preview_group_pages - 1)
        self.preview_page_label.configure(text=f"预览页：{self.preview_page + 1}/{self.preview_group_pages}")

        start_idx = self.preview_page * PREVIEW_SIZE_GROUPS_PER_PAGE
        end_idx = min(len(size_groups), start_idx + PREVIEW_SIZE_GROUPS_PER_PAGE)
        page_groups = size_groups[start_idx:end_idx]

        row1_parts = ["[空]"] + [f"[尺码{sc.size_no}]" for sc in page_groups]
        row2_parts = ["款色"]
        for sc in page_groups:
            row2_parts.extend(["公式变成", sc.size_no])

        self.result_header_row1.configure(text="第1行：  " + "   ".join(row1_parts))
        self.result_header_row2.configure(text="第2行：  " + "   ".join(str(x) for x in row2_parts))
        self.result_header_summary.configure(
            text=f"当前显示分组：尺码{page_groups[0].size_no} 到 尺码{page_groups[-1].size_no}。导出到 Excel 时会严格按两层表头生成：第1行空+尺码标题，第2行款色/公式变成/尺码号。"
        )

        columns = ["款色"]
        for sc in page_groups:
            columns.extend([f"公式变成/尺码{sc.size_no}", sc.size_no])

        self.result_tree["columns"] = columns
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        for col in columns:
            self.result_tree.heading(col, text=col)
            self.result_tree.column(col, width=160 if "公式变成" in col or col == "款色" else 90, anchor="center")

        for row in self.preview_rows:
            row_values = [row["款色"]]
            for v in row["values"][start_idx:end_idx]:
                row_values.extend([v["concat_value"], v["original_value"]])
            self.result_tree.insert("", "end", values=row_values)

    def generate_preview(self):
        try:
            ws = self.get_current_ws()
            header_row = int(self.header_row_var.get())
            self.detect_result = WorkbookAnalyzer.detect_columns(ws, header_row)
            self.show_detect_info(self.detect_result)
            self.preview_rows = WorkbookAnalyzer.build_preview(ws, self.detect_result)
            self.preview_page = 0

            self.build_source_preview(self.detect_result)
            self.build_result_preview(self.detect_result)
            self.confirm_preview_var.set(False)
            self.log(f"预览已生成，共展示前 {len(self.preview_rows)} 行有效数据。")
            if not self.preview_rows:
                self.log("提示：没有可预览的有效数据行，请检查表头行或数据内容。")
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))
            self.log(f"生成预览失败：{e}")

    def prev_preview_page(self):
        if self.preview_page > 0 and self.detect_result is not None:
            self.preview_page -= 1
            self.build_result_preview(self.detect_result)

    def next_preview_page(self):
        if self.detect_result is None:
            return
        if self.preview_page + 1 < self.preview_group_pages:
            self.preview_page += 1
            self.build_result_preview(self.detect_result)

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

        detect = self.detect_result
        size_text = "、".join([f"尺码{s.size_no}" for s in detect.size_columns])
        summary = (
            f"即将处理：\n"
            f"- 文件：{os.path.basename(input_path)}\n"
            f"- 工作表：{detect.sheet_name}\n"
            f"- 表头行：{detect.header_row}\n"
            f"- 款色列：{get_column_letter(detect.color_col_idx)}\n"
            f"- 尺码列：{size_text}\n"
            f"- 结果格式：\n"
            f"  第1行：空 + 尺码1/尺码2/...\n"
            f"  第2行：款色 + 公式变成 + 1/2/...\n"
            f"  第3行开始：款色+尺码号 与原值\n"
            f"- 保存到：{output_path}\n\n"
            f"是否继续？"
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
