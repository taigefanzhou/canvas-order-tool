"""
帆布订单整理工具
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
        elif name in ('数量', '购买数量', '订购数量'):
            col_map['qty'] = idx
        elif name in ('备注', '买家备注', '卖家备注'):
            col_map['remark'] = idx
        elif name in ('买家留言',):
            col_map['buyer_msg'] = idx

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

    # 创建输出工作簿
    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "Sheet1"

    header_font = Font(bold=True, size=11)
    subtotal_font = Font(bold=True, size=11, color="000000")
    subtotal_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    # 表头增加"平方数"列
    headers = ['序号', '订单号', '规格名称', '规格编码', '数量', '备注', '', '尺寸', '总数量', '总平方数']
    for col, h in enumerate(headers, 1):
        cell = out_ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center')
        if col <= 6 or col >= 8:
            cell.border = thin_border

    current_row = 2
    seq = 1
    total_qty = 0
    total_area = 0.0
    summary_data = []

    for size in sorted_sizes:
        group = grouped[size]
        group_qty = sum(o['qty'] for o in group)
        area_per = parse_size_area(size)
        group_area = area_per * group_qty
        total_qty += group_qty
        total_area += group_area
        summary_data.append((size, group_qty, group_area))

        for order in group:
            out_ws.cell(row=current_row, column=1, value=seq).border = thin_border
            out_ws.cell(row=current_row, column=2, value=order['order_no']).border = thin_border
            out_ws.cell(row=current_row, column=3, value=order['spec_name']).border = thin_border
            out_ws.cell(row=current_row, column=4, value=order['spec_code']).border = thin_border
            c5 = out_ws.cell(row=current_row, column=5, value=order['qty'])
            c5.border = thin_border
            c5.alignment = Alignment(horizontal='center')
            out_ws.cell(row=current_row, column=6, value=order['remark']).border = thin_border
            seq += 1
            current_row += 1

        subtotal_label = f"【{size}】小计"
        c1 = out_ws.cell(row=current_row, column=1, value=subtotal_label)
        c1.font = subtotal_font
        c1.fill = subtotal_fill
        for col in range(1, 7):
            out_ws.cell(row=current_row, column=col).fill = subtotal_fill
            out_ws.cell(row=current_row, column=col).border = thin_border
            out_ws.cell(row=current_row, column=col).font = subtotal_font
        c5 = out_ws.cell(row=current_row, column=5, value=group_qty)
        c5.font = subtotal_font
        c5.fill = subtotal_fill
        c5.border = thin_border
        c5.alignment = Alignment(horizontal='center')
        current_row += 1

    # 总计行
    current_row += 1
    out_ws.cell(row=current_row, column=1, value="总计").font = Font(bold=True, size=12)
    c5 = out_ws.cell(row=current_row, column=5, value=total_qty)
    c5.font = Font(bold=True, size=12)
    c5.alignment = Alignment(horizontal='center')

    # 右侧汇总表（H-J列）
    for i, (size, qty, area) in enumerate(summary_data):
        r = i + 2
        c8 = out_ws.cell(row=r, column=8, value=size)
        c8.border = thin_border
        c8.alignment = Alignment(horizontal='center')
        c9 = out_ws.cell(row=r, column=9, value=qty)
        c9.border = thin_border
        c9.alignment = Alignment(horizontal='center')
        c10 = out_ws.cell(row=r, column=10, value=round(area, 2))
        c10.border = thin_border
        c10.alignment = Alignment(horizontal='center')

    # 汇总总计行
    summary_total_row = len(summary_data) + 2
    for col_idx, val in [(8, "总计"), (9, total_qty), (10, round(total_area, 2))]:
        c = out_ws.cell(row=summary_total_row, column=col_idx, value=val)
        c.font = Font(bold=True)
        c.border = thin_border
        c.alignment = Alignment(horizontal='center')

    # 列宽
    out_ws.column_dimensions['A'].width = 16
    out_ws.column_dimensions['B'].width = 28
    out_ws.column_dimensions['C'].width = 55
    out_ws.column_dimensions['D'].width = 12
    out_ws.column_dimensions['E'].width = 8
    out_ws.column_dimensions['F'].width = 35
    out_ws.column_dimensions['G'].width = 3
    out_ws.column_dimensions['H'].width = 14
    out_ws.column_dimensions['I'].width = 10
    out_ws.column_dimensions['J'].width = 12

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
        self.root.title("帆布订单整理工具")
        self.root.configure(bg=self.BG)
        self.root.resizable(False, False)

        w, h = 600, 530
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
        # 标题
        ttk.Label(self.root, text="帆布订单整理工具", style="Title.TLabel").pack(pady=(20, 12))

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
