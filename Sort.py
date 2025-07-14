#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string


def sort_sheet(filepath, savepath, sheet_name, group_by_unit=False):
    # 读取带缓存值的 workbook
    wb_values = load_workbook(filepath, data_only=True)
    ws_values = wb_values[sheet_name]
    # 读取保留公式的 workbook
    wb = load_workbook(filepath)
    ws = wb[sheet_name]

    # 不同表头行
    HDR_ROW = 6 if sheet_name != "主要领导" else 4
    COL_A      = 1  # 序号列 A
    COL_B      = column_index_from_string("B")
    COL_UNIT   = column_index_from_string("C")
    COL_SCORE  = column_index_from_string("H")
    COL_NOTE   = column_index_from_string("I") if sheet_name != "主要领导" else None

    merged_ranges = ws.merged_cells.ranges
    def is_merged_non_anchor(r, c):
        for m in merged_ranges:
            if m.min_row <= r <= m.max_row and m.min_col <= c <= m.max_col:
                return not (r == m.min_row and c == m.min_col)
        return False

    # 找到最后一行：根据总分列是否有值
    last_data_row = HDR_ROW
    for r in range(ws.max_row, HDR_ROW, -1):
        if ws.cell(r, COL_SCORE).value not in (None, ""):
            last_data_row = r
            break

    # 收集所有行数据
    rows = []  # list of tuples (原行号, row_values, note, score)
    for rr in range(HDR_ROW + 1, last_data_row + 1):
        # 读取整行原始值
        raw_vals = [ws.cell(rr, c).value for c in range(1, ws.max_column + 1)]
        # 对于非正职公务员，补齐 B/C 合并单元格
        if group_by_unit:
            for col in (COL_B, COL_UNIT):
                if raw_vals[col-1] in (None, ""):
                    for m in merged_ranges:
                        if m.min_col == col and m.min_row <= rr <= m.max_row:
                            raw_vals[col-1] = ws.cell(m.min_row, m.min_col).value
                            break
        # 获取备注（仅非正职公务员）
        note = raw_vals[COL_NOTE - 1] if sheet_name != "主要领导" else ""
        # 获取分数缓存值
        raw_score = ws_values.cell(rr, COL_SCORE).value
        try:
            score = float(raw_score)
        except (TypeError, ValueError):
            score = 0.0
        rows.append((rr, raw_vals, note, score))

    # 排序逻辑
    if group_by_unit:
        # 对非正职公务员按单位分组后排序
        sorted_rows = []
        groups = []  # list of (start, end) index in rows
        start = 0
        last_unit = None
        for i, (_, raw_vals, _, _) in enumerate(rows):
            unit = raw_vals[COL_UNIT-1]
            if unit and unit != last_unit and last_unit is not None:
                groups.append((start, i))
                start = i
            last_unit = unit or last_unit
        groups.append((start, len(rows)))
        # 排序各组
        for s, e in groups:
            grp = rows[s:e]
            grp.sort(key=lambda x: (bool(x[2]), -x[3]))
            sorted_rows.extend(grp)
    else:
        # 主要领导，整体按总分降序
        sorted_rows = sorted(rows, key=lambda x: -x[3])

    # 写回排序结果
    for idx, (_orig_rr, vals, _, _) in enumerate(sorted_rows, start=HDR_ROW+1):
        for c, v in enumerate(vals, start=1):
            if not is_merged_non_anchor(idx, c):
                ws.cell(idx, c).value = v

    # 重写分数公式 & 序号
    for i, (_orig_rr, _vals, _note, _score) in enumerate(sorted_rows, start=1):
        r = HDR_ROW + i
        # 序号
        if not is_merged_non_anchor(r, COL_A):
            ws.cell(r, COL_A).value = i
        # 公式重写
        if sheet_name == "主要领导":
            # D35% + E25% + F20% + G20%
            ws.cell(r, COL_SCORE).value = f"=D{r}*35%+E{r}*25%+F{r}*20%+G{r}*20%"
        else:
            # E35% + F30% + G35%
            ws.cell(r, COL_SCORE).value = f"=SUM(E{r}*35%+F{r}*30%+G{r}*35%)"

    wb.save(savepath)


def choose_input():
    path = filedialog.askopenfilename(
        title="选择源文件",
        filetypes=[("Excel 文件", "*.xlsx;*.xlsm;*.xltx;*.xltm")]
    )
    if path:
        input_var.set(path)


def choose_output():
    path = filedialog.asksaveasfilename(
        title="选择保存路径",
        defaultextension=".xlsx",
        filetypes=[("Excel 文件", "*.xlsx")]
    )
    if path:
        output_var.set(path)


def run_sort():
    in_path = input_var.get()
    out_path = output_var.get()
    sheet = sheet_var.get()
    if not in_path or not out_path:
        messagebox.showwarning("提示", "请先选择输入文件和保存路径。")
        return
    try:
        if sheet in ("非正职公务员", "全部"):
            sort_sheet(in_path, out_path, "非正职公务员", group_by_unit=True)
        if sheet in ("主要领导", "全部"):
            sort_sheet(in_path, out_path, "主要领导", group_by_unit=False)
        messagebox.showinfo("成功", "排序并保存完成！")
    except Exception as e:
        messagebox.showerror("错误", f"处理文件时出错：{e}")

# 主界面
root = tk.Tk()
root.title("排序工具")
root.geometry("540x180")

input_var = tk.StringVar()
output_var = tk.StringVar()
sheet_var = tk.StringVar(value="非正职公务员")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(fill=tk.BOTH, expand=True)

# 输入文件
tk.Label(frame, text="源文件:").grid(row=0, column=0, sticky=tk.W)
tk.Entry(frame, textvariable=input_var, width=48).grid(row=0, column=1)
tk.Button(frame, text="选择...", command=choose_input).grid(row=0, column=2)

# 保存路径
tk.Label(frame, text="保存为:").grid(row=1, column=0, sticky=tk.W)
tk.Entry(frame, textvariable=output_var, width=48).grid(row=1, column=1)
tk.Button(frame, text="选择...", command=choose_output).grid(row=1, column=2)

# 工作表选择
tk.Label(frame, text="选择要排序的工作表:").grid(row=2, column=0, sticky=tk.W)
tk.OptionMenu(frame, sheet_var, "非正职公务员", "主要领导", "全部").grid(row=2, column=1, sticky=tk.W)

# 执行按钮
tk.Button(frame, text="开始排序", command=run_sort, width=20).grid(row=3, column=1, pady=20)

root.mainloop()
