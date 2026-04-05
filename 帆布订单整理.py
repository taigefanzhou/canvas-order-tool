"""
帆布订单整理工具
功能：读取帆布订单原始数据Excel，按尺寸分组排序，生成带小计和汇总的明细表
"""

import re
import sys
import os
from datetime import datetime
from collections import OrderedDict

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from tkinter import Tk, filedialog, messagebox


def extract_size(spec_name):
    """从规格名称中提取尺寸，如 '2米*3米'，提取不到则返回 '定制'"""
    match = re.search(r'(\d+(?:\.\d+)?)\s*米\s*\*\s*(\d+(?:\.\d+)?)\s*米', spec_name)
    if match:
        w = float(match.group(1))
        h = float(match.group(2))
        # 格式化：整数不带小数点
        w_str = f"{int(w)}米" if w == int(w) else f"{w}米"
        h_str = f"{int(h)}米" if h == int(h) else f"{h}米"
        return f"{w_str}*{h_str}"
    return "定制"


def size_sort_key(size_str):
    """尺寸排序：按宽*高数值排序，定制排最后"""
    if size_str == "定制":
        return (9999, 9999)
    match = re.match(r'(\d+(?:\.\d+)?)米\*(\d+(?:\.\d+)?)米', size_str)
    if match:
        return (float(match.group(1)), float(match.group(2)))
    return (9999, 9999)


def process_orders(input_path):
    """处理订单数据"""
    wb = openpyxl.load_workbook(input_path)
    ws = wb.active

    # 读取数据（跳过表头，第1行是序号空，第2行开始是数据）
    orders = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        # 列: 序号, 订单号, 规格名称, 规格编码, 数量, 备注
        order_no = row[1]
        spec_name = row[2]
        spec_code = row[3]
        qty = row[4]
        remark = row[5]

        if not order_no or not spec_name:
            continue

        size = extract_size(str(spec_name))
        orders.append({
            'order_no': str(order_no),
            'spec_name': str(spec_name),
            'spec_code': spec_code or '',
            'qty': int(qty) if qty else 0,
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

    # 按尺寸排序
    sorted_sizes = sorted(grouped.keys(), key=size_sort_key)

    # 创建输出工作簿
    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "Sheet1"

    # 样式定义
    header_font = Font(bold=True, size=11)
    subtotal_font = Font(bold=True, size=11, color="000000")
    subtotal_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
    summary_header_font = Font(bold=True, size=11)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # 写表头
    headers = ['序号', '订单号', '规格名称', '规格编码', '数量', '备注', '', '尺寸', '总数量']
    for col, h in enumerate(headers, 1):
        cell = out_ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center')
        if col <= 6 or col >= 8:
            cell.border = thin_border

    # 写数据
    current_row = 2
    seq = 1
    total_qty = 0

    # 右侧汇总数据
    summary_data = []

    for size in sorted_sizes:
        group = grouped[size]
        group_qty = sum(o['qty'] for o in group)
        total_qty += group_qty
        summary_data.append((size, group_qty))

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

        # 小计行
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

    # 空行 + 总计行
    current_row += 1
    out_ws.cell(row=current_row, column=1, value="总计").font = Font(bold=True, size=12)
    c5 = out_ws.cell(row=current_row, column=5, value=total_qty)
    c5.font = Font(bold=True, size=12)
    c5.alignment = Alignment(horizontal='center')

    # 右侧汇总表（H-I列，从第2行开始）
    for i, (size, qty) in enumerate(summary_data):
        r = i + 2
        c8 = out_ws.cell(row=r, column=8, value=size)
        c8.border = thin_border
        c8.alignment = Alignment(horizontal='center')
        c9 = out_ws.cell(row=r, column=9, value=qty)
        c9.border = thin_border
        c9.alignment = Alignment(horizontal='center')

    # 汇总表总计
    summary_total_row = len(summary_data) + 2
    c8 = out_ws.cell(row=summary_total_row, column=8, value="总计")
    c8.font = Font(bold=True)
    c8.border = thin_border
    c8.alignment = Alignment(horizontal='center')
    c9 = out_ws.cell(row=summary_total_row, column=9, value=total_qty)
    c9.font = Font(bold=True)
    c9.border = thin_border
    c9.alignment = Alignment(horizontal='center')

    # 设置列宽
    out_ws.column_dimensions['A'].width = 16
    out_ws.column_dimensions['B'].width = 28
    out_ws.column_dimensions['C'].width = 55
    out_ws.column_dimensions['D'].width = 12
    out_ws.column_dimensions['E'].width = 8
    out_ws.column_dimensions['F'].width = 35
    out_ws.column_dimensions['G'].width = 3
    out_ws.column_dimensions['H'].width = 14
    out_ws.column_dimensions['I'].width = 10

    # 生成输出文件名
    today = datetime.now().strftime("%Y%m%d")
    dir_name = os.path.dirname(input_path)
    output_path = os.path.join(dir_name, f"帆布订单明细_{today}.xlsx")

    out_wb.save(output_path)
    return output_path, len(orders), total_qty, len(sorted_sizes)


def main():
    root = Tk()
    root.withdraw()

    messagebox.showinfo("帆布订单整理工具", "请选择帆布订单原始数据Excel文件")

    input_path = filedialog.askopenfilename(
        title="选择帆布订单原始数据",
        filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
    )

    if not input_path:
        messagebox.showwarning("提示", "未选择文件，程序退出")
        return

    try:
        output_path, order_count, total_qty, size_count = process_orders(input_path)
        messagebox.showinfo(
            "处理完成",
            f"订单整理完成！\n\n"
            f"共 {order_count} 条订单\n"
            f"共 {size_count} 种尺寸\n"
            f"总数量：{total_qty}\n\n"
            f"已保存到：\n{output_path}"
        )
        # 打开输出文件所在目录
        os.startfile(os.path.dirname(output_path))
    except Exception as e:
        messagebox.showerror("错误", f"处理失败：\n{str(e)}")

    root.destroy()


if __name__ == "__main__":
    main()
