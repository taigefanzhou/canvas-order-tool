"""
丽群帆布纺织电商统计系统
功能：读取帆布订单原始数据Excel，按尺寸分组排序，生成带小计和汇总的明细表
"""

import re
import sys
import os
import platform
import subprocess
import threading
from datetime import datetime
from collections import OrderedDict

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, StringVar
from PIL import Image, ImageTk


def extract_size(spec_name):
    """从规格名称中提取尺寸，如 '2米*3米'，提取不到则返回 '定制'"""
    match = re.search(r'(\d+(?:\.\d+)?)\s*米?\s*\*\s*(\d+(?:\.\d+)?)\s*米', spec_name)
    if match:
        w = float(match.group(1))
        h = float(match.group(2))
        w_str = f"{int(w)}米" if w == int(w) else f"{w}米"
        h_str = f"{int(h)}米" if h == int(h) else f"{h}米"
        return f"{w_str}*{h_str}"
    return "定制"


def parse_size_area(size_str):
    """从尺寸字符串计算面积（平方米），定制返回0"""
    if size_str == "定制":
        return 0.0
    match = re.match(r'(\d+(?:\.\d+)?)米\*(\d+(?:\.\d+)?)米', size_str)
    if match:
        return float(match.group(1)) * float(match.group(2))
    return 0.0


def size_sort_key(size_str):
    """尺寸排序：按宽*高数值排序，定制排最后"""
    if size_str == "定制":
        return (9999, 9999)
    match = re.match(r'(\d+(?:\.\d+)?)米\*(\d+(?:\.\d+)?)米', size_str)
    if match:
        return (float(match.group(1)), float(match.group(2)))
    return (9999, 9999)


def process_orders(input_path, output_dir):
    """处理订单数据"""
    wb = openpyxl.load_workbook(input_path)
    ws = wb.active

    # 读取表头，自动匹配列位置
    header_row = [str(cell.value).strip() if cell.value else '' for cell in ws[1]]

    col_map = {}
    for idx, name in enumerate(header_row):
        if name in ('订单号', '订单编号'):
            col_map['order_no'] = idx
        elif name in ('规格名称', '商品规格', '规格'):
            col_map['spec_name'] = idx
        elif name in ('规格编码', '编码'):
            col_map['spec_code'] = idx
        elif name in ('数量', '购买数量', '订购数量', '商品数量'):
            col_map['qty'] = idx
        elif name in ('备注', '买家备注', '卖家备注'):
            col_map['remark'] = idx
        elif name in ('买家留言',):
            col_map['buyer_msg'] = idx
        elif name in ('快递单号', '运单号', '物流单号'):
            col_map['tracking_no'] = idx

    missing = [k for k in ('order_no', 'spec_name', 'qty') if k not in col_map]
    if missing:
        col_names = {'order_no': '订单号', 'spec_name': '规格名称', 'qty': '数量'}
        raise ValueError(
            f"Excel表头中未找到必要的列：{', '.join(col_names[k] for k in missing)}\n"
            f"当前表头：{header_row}\n"
            f"请确认Excel文件格式是否正确"
        )

    orders = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        order_no = row[col_map['order_no']]
        spec_name = row[col_map['spec_name']]
        spec_code = row[col_map.get('spec_code', -1)] if col_map.get('spec_code') is not None and col_map.get('spec_code') < len(row) else ''
        qty = row[col_map['qty']]
        remark = row[col_map.get('remark', -1)] if col_map.get('remark') is not None and col_map.get('remark') < len(row) else ''
        buyer_msg = row[col_map.get('buyer_msg', -1)] if col_map.get('buyer_msg') is not None and col_map.get('buyer_msg') < len(row) else ''
        tracking_no = row[col_map.get('tracking_no', -1)] if col_map.get('tracking_no') is not None and col_map.get('tracking_no') < len(row) else ''

        if not order_no or not spec_name:
            continue

        # 跳过非帆布商品（如补收差价等）
        spec_str = str(spec_name)
        if '差价' in spec_str or '补' in spec_str:
            continue

        size = extract_size(spec_str)

        # 合并备注和买家留言
        remark_parts = []
        if remark:
            remark_parts.append(str(remark))
        if buyer_msg:
            remark_parts.append(str(buyer_msg))
        remark_text = ' | '.join(remark_parts)
        try:
            qty_val = int(qty) if qty else 0
        except (ValueError, TypeError):
            qty_val = 0

        orders.append({
            'order_no': str(order_no),
            'spec_name': str(spec_name),
            'spec_code': str(spec_code) if spec_code else '',
            'qty': qty_val,
            'tracking_no': str(tracking_no) if tracking_no else '',
            'remark': remark_text,
            'size': size,
        })

    # 按尺寸分组
    grouped = OrderedDict()
    for order in orders:
        size = order['size']
        if size not in grouped:
            grouped[size] = []
        grouped[size].append(order)

    sorted_sizes = sorted(grouped.keys(), key=size_sort_key)

    # 按面积拆分为小件组（≤10m²）和大件组（>10m²）
    small_sizes = [s for s in sorted_sizes if parse_size_area(s) <= 10]
    large_sizes = [s for s in sorted_sizes if parse_size_area(s) > 10]

    # 创建输出工作簿
    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "Sheet1"

    header_font = Font(bold=True, size=11)
    subtotal_font = Font(bold=True, size=11, color="000000")
    subtotal_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
    section_font = Font(bold=True, size=12, color="FFFFFF")
    section_fill_small = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    section_fill_large = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
    block_total_font = Font(bold=True, size=11, color="FFFFFF")
    block_total_fill_small = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
    block_total_fill_large = PatternFill(start_color="E04040", end_color="E04040", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    # 表头
    headers = ['序号', '订单号', '规格名称', '规格编码', '数量', '快递单号', '备注', '', '尺寸', '总数量', '总平方数']
    for col, h in enumerate(headers, 1):
        cell = out_ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center')
        if col <= 7 or col >= 9:
            cell.border = thin_border

    current_row = 2
    seq = 1
    total_qty = 0
    total_area = 0.0
    summary_data_small = []
    summary_data_large = []

    def write_section_title(ws, row, title, fill):
        """写分区标题行"""
        c = ws.cell(row=row, column=1, value=title)
        c.font = section_font
        c.fill = fill
        c.alignment = Alignment(horizontal='left')
        for col in range(1, 8):
            ws.cell(row=row, column=col).fill = fill
            ws.cell(row=row, column=col).font = section_font
            ws.cell(row=row, column=col).border = thin_border
        return row + 1

    def write_block_total(ws, row, label, qty, area, font, fill):
        """写区块合计行"""
        c1 = ws.cell(row=row, column=1, value=label)
        c1.font = font
        c1.fill = fill
        for col in range(1, 8):
            ws.cell(row=row, column=col).fill = fill
            ws.cell(row=row, column=col).border = thin_border
            ws.cell(row=row, column=col).font = font
        c5 = ws.cell(row=row, column=5, value=qty)
        c5.font = font
        c5.fill = fill
        c5.border = thin_border
        c5.alignment = Alignment(horizontal='center')
        return row + 1

    def write_size_group(ws, row, seq_val, size, group, grouped_data):
        """写一个尺寸的明细和小计，返回 (new_row, new_seq, group_qty, group_area)"""
        group_qty = sum(o['qty'] for o in group)
        area_per = parse_size_area(size)
        group_area = area_per * group_qty

        for order in group:
            ws.cell(row=row, column=1, value=seq_val).border = thin_border
            ws.cell(row=row, column=2, value=order['order_no']).border = thin_border
            ws.cell(row=row, column=3, value=order['spec_name']).border = thin_border
            ws.cell(row=row, column=4, value=order['spec_code']).border = thin_border
            c5 = ws.cell(row=row, column=5, value=order['qty'])
            c5.border = thin_border
            c5.alignment = Alignment(horizontal='center')
            ws.cell(row=row, column=6, value=order['tracking_no']).border = thin_border
            ws.cell(row=row, column=7, value=order['remark']).border = thin_border
            seq_val += 1
            row += 1

        subtotal_label = f"【{size}】小计"
        c1 = ws.cell(row=row, column=1, value=subtotal_label)
        c1.font = subtotal_font
        c1.fill = subtotal_fill
        for col in range(1, 8):
            ws.cell(row=row, column=col).fill = subtotal_fill
            ws.cell(row=row, column=col).border = thin_border
            ws.cell(row=row, column=col).font = subtotal_font
        c5 = ws.cell(row=row, column=5, value=group_qty)
        c5.font = subtotal_font
        c5.fill = subtotal_fill
        c5.border = thin_border
        c5.alignment = Alignment(horizontal='center')
        row += 1

        return row, seq_val, group_qty, group_area

    # ── 小件区块（≤10平方米）──
    block_small_qty = 0
    block_small_area = 0.0

    if small_sizes:
        current_row = write_section_title(out_ws, current_row, "≤ 10平方米（含）", section_fill_small)

        for size in small_sizes:
            group = grouped[size]
            current_row, seq, gq, ga = write_size_group(out_ws, current_row, seq, size, group, None)
            block_small_qty += gq
            block_small_area += ga
            summary_data_small.append((size, gq, ga))

        current_row = write_block_total(
            out_ws, current_row,
            f"10平方米以下（含）合计",
            block_small_qty, block_small_area,
            block_total_font, block_total_fill_small
        )
        current_row += 1  # 空行分隔

    # ── 大件区块（>10平方米）──
    block_large_qty = 0
    block_large_area = 0.0

    if large_sizes:
        current_row = write_section_title(out_ws, current_row, "> 10平方米", section_fill_large)

        for size in large_sizes:
            group = grouped[size]
            current_row, seq, gq, ga = write_size_group(out_ws, current_row, seq, size, group, None)
            block_large_qty += gq
            block_large_area += ga
            summary_data_large.append((size, gq, ga))

        current_row = write_block_total(
            out_ws, current_row,
            f"10平方米以上合计",
            block_large_qty, block_large_area,
            block_total_font, block_total_fill_large
        )

    total_qty = block_small_qty + block_large_qty
    total_area = block_small_area + block_large_area

    # 总计行
    current_row += 1
    out_ws.cell(row=current_row, column=1, value="总计").font = Font(bold=True, size=12)
    c5 = out_ws.cell(row=current_row, column=5, value=total_qty)
    c5.font = Font(bold=True, size=12)
    c5.alignment = Alignment(horizontal='center')

    # 右侧汇总表（I-K列）—— 分两段
    summary_row = 2

    # 小件汇总
    if summary_data_small:
        c_title = out_ws.cell(row=summary_row, column=9, value="≤10m²")
        c_title.font = Font(bold=True, color="FFFFFF")
        c_title.fill = section_fill_small
        c_title.alignment = Alignment(horizontal='center')
        c_title.border = thin_border
        for ci in (10, 11):
            out_ws.cell(row=summary_row, column=ci).fill = section_fill_small
            out_ws.cell(row=summary_row, column=ci).border = thin_border
        summary_row += 1

        for size, qty, area in summary_data_small:
            for col_idx, val in [(9, size), (10, qty), (11, round(area, 2))]:
                c = out_ws.cell(row=summary_row, column=col_idx, value=val)
                c.border = thin_border
                c.alignment = Alignment(horizontal='center')
            summary_row += 1

        # 小件小计
        for col_idx, val in [(9, "小计"), (10, block_small_qty), (11, round(block_small_area, 2))]:
            c = out_ws.cell(row=summary_row, column=col_idx, value=val)
            c.font = Font(bold=True)
            c.fill = block_total_fill_small
            c.font = Font(bold=True, color="FFFFFF")
            c.border = thin_border
            c.alignment = Alignment(horizontal='center')
        summary_row += 2  # 空行分隔

    # 大件汇总
    if summary_data_large:
        c_title = out_ws.cell(row=summary_row, column=9, value=">10m²")
        c_title.font = Font(bold=True, color="FFFFFF")
        c_title.fill = section_fill_large
        c_title.alignment = Alignment(horizontal='center')
        c_title.border = thin_border
        for ci in (10, 11):
            out_ws.cell(row=summary_row, column=ci).fill = section_fill_large
            out_ws.cell(row=summary_row, column=ci).border = thin_border
        summary_row += 1

        for size, qty, area in summary_data_large:
            for col_idx, val in [(9, size), (10, qty), (11, round(area, 2))]:
                c = out_ws.cell(row=summary_row, column=col_idx, value=val)
                c.border = thin_border
                c.alignment = Alignment(horizontal='center')
            summary_row += 1

        # 大件小计
        for col_idx, val in [(9, "小计"), (10, block_large_qty), (11, round(block_large_area, 2))]:
            c = out_ws.cell(row=summary_row, column=col_idx, value=val)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = block_total_fill_large
            c.border = thin_border
            c.alignment = Alignment(horizontal='center')
        summary_row += 2

    # 汇总总计行
    for col_idx, val in [(9, "总计"), (10, total_qty), (11, round(total_area, 2))]:
        c = out_ws.cell(row=summary_row, column=col_idx, value=val)
        c.font = Font(bold=True)
        c.border = thin_border
        c.alignment = Alignment(horizontal='center')

    # 列宽
    out_ws.column_dimensions['A'].width = 16
    out_ws.column_dimensions['B'].width = 28
    out_ws.column_dimensions['C'].width = 55
    out_ws.column_dimensions['D'].width = 12
    out_ws.column_dimensions['E'].width = 8
    out_ws.column_dimensions['F'].width = 22
    out_ws.column_dimensions['G'].width = 35
    out_ws.column_dimensions['H'].width = 3
    out_ws.column_dimensions['I'].width = 14
    out_ws.column_dimensions['J'].width = 10
    out_ws.column_dimensions['K'].width = 12

    today = datetime.now().strftime("%Y%m%d")
    base_name = f"帆布订单明细_{today}"
    output_path = os.path.join(output_dir, f"{base_name}.xlsx")

    # 如果文件被占用（Excel打开中），自动加编号避免冲突
    counter = 2
    while os.path.exists(output_path):
        try:
            with open(output_path, 'a'):
                break  # 文件没被占用，可以覆盖
        except PermissionError:
            output_path = os.path.join(output_dir, f"{base_name}_{counter}.xlsx")
            counter += 1

    out_wb.save(output_path)
    return output_path, len(orders), total_qty, len(sorted_sizes), round(total_area, 2)


def open_folder(path):
    """跨平台打开文件夹"""
    system = platform.system()
    if system == "Windows":
        os.startfile(path)
    elif system == "Darwin":
        subprocess.call(["open", path])
    else:
        subprocess.call(["xdg-open", path])


def resource_path(filename):
    """获取资源文件路径（兼容 PyInstaller 打包）"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)


class OrderApp:
    # 配色方案
    BG = "#f0f4f8"
    CARD_BG = "#ffffff"
    PRIMARY = "#4a90d9"
    SUCCESS = "#28a745"
    DANGER = "#dc3545"
    TEXT = "#333333"
    TEXT_LIGHT = "#888888"
    BORDER = "#d0d7de"

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("丽群帆布纺织电商统计系统")
        self.root.configure(bg=self.BG)
        self.root.resizable(False, False)

        # 设置窗口图标
        try:
            ico_path = resource_path("logo.ico")
            self.root.iconbitmap(ico_path)
        except Exception:
            pass

        w, h = 600, 580
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

        self.input_path = StringVar()
        self.output_dir = StringVar()
        self.output_path = None

        self._setup_styles()
        self._build_ui()
        self.root.mainloop()

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Title.TLabel", background=self.BG, foreground=self.PRIMARY,
                        font=("Microsoft YaHei", 18, "bold"))
        style.configure("Card.TFrame", background=self.CARD_BG)
        style.configure("CardTitle.TLabel", background=self.CARD_BG, foreground=self.TEXT,
                        font=("Microsoft YaHei", 10, "bold"))
        style.configure("Path.TEntry", font=("Microsoft YaHei", 9))

        style.configure("Primary.TButton", font=("Microsoft YaHei", 9),
                        background=self.PRIMARY, foreground="white")
        style.map("Primary.TButton",
                  background=[("active", "#3a7bc8"), ("disabled", "#a0b4c8")])

        style.configure("Success.TButton", font=("Microsoft YaHei", 11, "bold"),
                        background=self.SUCCESS, foreground="white", padding=(20, 10))
        style.map("Success.TButton",
                  background=[("active", "#218838"), ("disabled", "#a0c8a0")])

        style.configure("Info.TButton", font=("Microsoft YaHei", 9),
                        background="#17a2b8", foreground="white")
        style.map("Info.TButton",
                  background=[("active", "#138496"), ("disabled", "#a0c8d0")])

        style.configure("Result.TLabel", background=self.CARD_BG, foreground=self.TEXT_LIGHT,
                        font=("Microsoft YaHei", 11), wraplength=520, justify="left")

        style.configure("green.Horizontal.TProgressbar", troughcolor="#e0e0e0",
                        background=self.SUCCESS, thickness=8)

    def _make_card(self, parent, title_text, pady=(0, 8)):
        outer = tk.Frame(parent, bg=self.BG)
        outer.pack(fill="x", padx=24, pady=pady)

        title = ttk.Label(outer, text=title_text, style="CardTitle.TLabel")
        title.configure(background=self.BG)
        title.pack(anchor="w", pady=(0, 4))

        card = tk.Frame(outer, bg=self.CARD_BG, highlightbackground=self.BORDER,
                        highlightthickness=1, padx=12, pady=10)
        card.pack(fill="x")
        return card

    def _build_ui(self):
        # 顶部标题区域：logo + 文字
        header = tk.Frame(self.root, bg=self.BG)
        header.pack(pady=(16, 10))

        try:
            logo_img = Image.open(resource_path("logo.png"))
            logo_img = logo_img.resize((48, 48), Image.LANCZOS)
            self._logo_photo = ImageTk.PhotoImage(logo_img)
            tk.Label(header, image=self._logo_photo, bg=self.BG).pack(side="left", padx=(0, 10))
        except Exception:
            pass

        ttk.Label(header, text="丽群帆布纺织电商统计系统", style="Title.TLabel").pack(side="left")

        # 选择文件卡片
        file_card = self._make_card(self.root, "原始数据文件", pady=(0, 10))
        file_row = tk.Frame(file_card, bg=self.CARD_BG)
        file_row.pack(fill="x")
        self.file_entry = ttk.Entry(file_row, textvariable=self.input_path,
                                    state="readonly", style="Path.TEntry")
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(file_row, text="选择文件", style="Primary.TButton",
                   command=self._select_file, width=10).pack(side="right")

        # 保存位置卡片
        save_card = self._make_card(self.root, "保存位置", pady=(0, 10))
        save_row = tk.Frame(save_card, bg=self.CARD_BG)
        save_row.pack(fill="x")
        self.save_entry = ttk.Entry(save_row, textvariable=self.output_dir,
                                    state="readonly", style="Path.TEntry")
        self.save_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(save_row, text="选择文件夹", style="Primary.TButton",
                   command=self._select_output_dir, width=10).pack(side="right")

        # 开始处理按钮
        self.process_btn = ttk.Button(self.root, text="开始处理", style="Success.TButton",
                                      command=self._start_process, width=18)
        self.process_btn.pack(pady=(8, 5))

        # 进度条
        self.progress = ttk.Progressbar(self.root, mode="indeterminate",
                                        style="green.Horizontal.TProgressbar", length=540)
        self.progress.pack(pady=(0, 8))

        # 结果卡片
        result_card = self._make_card(self.root, "处理结果", pady=(0, 10))
        self.result_label = ttk.Label(result_card, text="等待处理...", style="Result.TLabel")
        self.result_label.pack(fill="x")

        # 打开文件夹按钮
        self.open_btn = ttk.Button(self.root, text="打开文件夹", style="Info.TButton",
                                   command=self._open_output_folder, width=14, state="disabled")
        self.open_btn.pack(pady=(0, 15))

    def _select_file(self):
        path = filedialog.askopenfilename(
            title="选择帆布订单原始数据",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if path:
            self.input_path.set(path)
            if not self.output_dir.get():
                self.output_dir.set(os.path.dirname(path))

    def _select_output_dir(self):
        initial = self.output_dir.get() or os.path.expanduser("~")
        path = filedialog.askdirectory(title="选择保存文件夹", initialdir=initial)
        if path:
            self.output_dir.set(path)

    def _start_process(self):
        if not self.input_path.get():
            messagebox.showwarning("提示", "请先选择原始数据文件")
            return
        if not self.output_dir.get():
            messagebox.showwarning("提示", "请先选择保存位置")
            return

        self.process_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.result_label.config(text="正在处理中...", foreground="gray")
        self.progress.start(15)

        thread = threading.Thread(target=self._do_process, daemon=True)
        thread.start()

    def _do_process(self):
        try:
            result = process_orders(self.input_path.get(), self.output_dir.get())
            self.output_path = result[0]
            self.root.after(0, self._on_success, result)
        except Exception as e:
            self.root.after(0, self._on_error, str(e))

    def _on_success(self, result):
        output_path, order_count, total_qty, size_count, total_area = result
        self.progress.stop()
        self.result_label.config(
            text=(
                f"处理完成！\n"
                f"共 {order_count} 条订单    |    共 {size_count} 种尺寸\n"
                f"总数量：{total_qty}    |    总平方数：{total_area} m\u00b2\n"
                f"已保存到：{output_path}"
            ),
        )
        self.result_label.configure(foreground=self.SUCCESS)
        self.process_btn.config(state="normal")
        self.open_btn.config(state="normal")

    def _on_error(self, msg):
        self.progress.stop()
        self.result_label.config(text=f"处理失败：\n{msg}")
        self.result_label.configure(foreground=self.DANGER)
        self.process_btn.config(state="normal")

    def _open_output_folder(self):
        if self.output_path:
            open_folder(os.path.dirname(self.output_path))


def main():
    OrderApp()


if __name__ == "__main__":
    main()
