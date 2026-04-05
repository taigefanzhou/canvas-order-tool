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

try:
    import ttkbootstrap as ttk
    from ttkbootstrap.constants import *
    from ttkbootstrap.dialogs import Messagebox
except ImportError:
    print("请先安装 ttkbootstrap: pip install ttkbootstrap")
    sys.exit(1)

from tkinter import filedialog, StringVar


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

        if not order_no or not spec_name:
            continue

        size = extract_size(str(spec_name))
        try:
            qty_val = int(qty) if qty else 0
        except (ValueError, TypeError):
            qty_val = 0

        orders.append({
            'order_no': str(order_no),
            'spec_name': str(spec_name),
            'spec_code': str(spec_code) if spec_code else '',
            'qty': qty_val,
            'remark': str(remark) if remark else '',
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
    output_path = os.path.join(output_dir, f"帆布订单明细_{today}.xlsx")
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
    def __init__(self):
        self.root = ttk.Window(
            title="帆布订单整理工具",
            themename="cosmo",
            size=(620, 560),
            resizable=(False, False),
        )
        self._center_window()
        self.input_path = StringVar()
        self.output_dir = StringVar()
        self.output_path = None
        self._build_ui()
        self.root.mainloop()

    def _center_window(self):
        self.root.update_idletasks()
        w, h = 620, 560
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        # 标题
        title = ttk.Label(
            self.root, text="帆布订单整理工具",
            font=("Microsoft YaHei", 18, "bold"),
            bootstyle="primary",
        )
        title.pack(pady=(20, 10))

        # 选择文件区域
        file_frame = ttk.LabelFrame(self.root, text="  选择原始数据文件  ", padding=12)
        file_frame.pack(fill="x", padx=24, pady=(5, 8))

        file_row = ttk.Frame(file_frame)
        file_row.pack(fill="x")
        self.file_entry = ttk.Entry(file_row, textvariable=self.input_path, state="readonly")
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(
            file_row, text="选择文件", bootstyle="outline-primary",
            command=self._select_file, width=10,
        ).pack(side="right")

        # 保存位置区域
        save_frame = ttk.LabelFrame(self.root, text="  选择保存位置  ", padding=12)
        save_frame.pack(fill="x", padx=24, pady=(0, 8))

        save_row = ttk.Frame(save_frame)
        save_row.pack(fill="x")
        self.save_entry = ttk.Entry(save_row, textvariable=self.output_dir, state="readonly")
        self.save_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(
            save_row, text="选择文件夹", bootstyle="outline-primary",
            command=self._select_output_dir, width=10,
        ).pack(side="right")

        # 开始处理按钮
        self.process_btn = ttk.Button(
            self.root, text="开始处理", bootstyle="success",
            command=self._start_process, width=20,
            padding=(10, 8),
        )
        self.process_btn.pack(pady=(10, 5))

        # 进度条
        self.progress = ttk.Progressbar(
            self.root, mode="indeterminate", bootstyle="success-striped",
        )
        self.progress.pack(fill="x", padx=24, pady=(0, 8))

        # 结果区域
        result_frame = ttk.LabelFrame(self.root, text="  处理结果  ", padding=12)
        result_frame.pack(fill="x", padx=24, pady=(0, 8))

        self.result_label = ttk.Label(
            result_frame, text="等待处理...",
            font=("Microsoft YaHei", 11),
            foreground="gray",
            wraplength=540, justify="left",
        )
        self.result_label.pack(fill="x")

        # 打开文件夹按钮
        self.open_btn = ttk.Button(
            self.root, text="打开文件夹", bootstyle="info-outline",
            command=self._open_output_folder, width=16, state="disabled",
        )
        self.open_btn.pack(pady=(5, 15))

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
            Messagebox.show_warning("请先选择原始数据文件", title="提示")
            return
        if not self.output_dir.get():
            Messagebox.show_warning("请先选择保存位置", title="提示")
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
                f"总数量：{total_qty}    |    总平方数：{total_area} m²\n"
                f"已保存到：{output_path}"
            ),
            foreground="#28a745",
        )
        self.process_btn.config(state="normal")
        self.open_btn.config(state="normal")

    def _on_error(self, msg):
        self.progress.stop()
        self.result_label.config(text=f"处理失败：\n{msg}", foreground="#dc3545")
        self.process_btn.config(state="normal")


def main():
    OrderApp()


if __name__ == "__main__":
    main()
