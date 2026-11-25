#基本函数定义
import numpy as np
from openpyxl import Workbook
from openpyxl.utils import column_index_from_string
from openpyxl.utils import get_column_letter
def write_row_start_at(ws, start_cell: str | tuple, data_list: list):
    if isinstance(start_cell, str):
        col_start = column_index_from_string(start_cell[:1])
        row_start = int(start_cell[1:])
    else:
        row_start, col_start = start_cell
    for offset, val in enumerate(data_list):
        ws.cell(row=row_start, column=col_start + offset, value=val)
import openpyxl
from openpyxl.utils import range_boundaries

def sum_excel_area(ws, area_str: str) -> float:
    """
    传入Excel区域字符串如"A1:C10"，自动求和
    """
    min_col, min_row, max_col, max_row = range_boundaries(area_str)
    total = 0.0
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            val = ws.cell(row=r, column=c).value
            if isinstance(val, (int, float)):
                total += val
    return total
from openpyxl.utils import column_index_from_string, get_column_letter
import re
def CopyFormula(src_cell: str, direction: str, steps: int, ws):
    """
    复刻Excel公式拖拽填充逻辑，自动修正相对引用，$绝对引用固定不变
    :param src_cell: 源单元格字符串，例："B6"
    :param direction: 拖拽方向，可选 right / left / down / up
    :param steps: 向右/左/上/下填充几步
    :param ws: openpyxl工作表对象
    """
    # 拆分源单元格：列字母、行数字
    col_part = "".join(c for c in src_cell if c.isalpha())
    row_part = int("".join(c for c in src_cell if c.isdigit()))
    src_col = column_index_from_string(col_part)
    src_formula = ws[src_cell].value

    # 校验源单元格是公式
    if not isinstance(src_formula, str) or not src_formula.startswith("="):
        raise Exception(f"源单元格 {src_cell} 没有Excel公式，无法复制拖拽")

    # 循环填充每一格
    for step in range(1, steps + 1):
        dst_col, dst_row = src_col, row_part

        # 目标单元格行列偏移
        if direction == "right":
            dst_col += step
        elif direction == "left":
            dst_col -= step
        elif direction == "down":
            dst_row += step
        elif direction == "up":
            dst_row -= step
        else:
            raise Exception('方向仅支持: right / left / down / up')

        dst_cell = f"{get_column_letter(dst_col)}{dst_row}"
        new_formula = src_formula

        # 正则匹配【无$的相对引用单元格】，自动偏移行列
        # 规则：$A1、A$1 、$A$1 都不改动；仅A1这种纯相对引用偏移
        match_rule = re.compile(r'(?<!\$)([A-Z]+)(?<!\$)(\d+)')
        ref_list = match_rule.findall(new_formula)

        for col_str, row_str in ref_list:
            old_c = column_index_from_string(col_str)
            old_r = int(row_str)
            new_c, new_r = old_c, old_r

            # 横向拖拽：只变列、行号固定
            if direction in ("right", "left"):
                delta = step if direction == "right" else -step
                new_c += delta
            # 纵向拖拽：只变行、列字母固定
            else:
                delta = step if direction == "down" else -step
                new_r += delta

            new_col_str = get_column_letter(new_c)
            old_ref = f"{col_str}{row_str}"
            new_ref = f"{new_col_str}{new_r}"
            new_formula = new_formula.replace(old_ref, new_ref, 1)

        # 写入目标单元格公式
        ws[dst_cell] = new_formula
def FinalColumn(start_cell, cmd_list, ws):
    # 拆分列标、起始行号
    tar_col = start_cell[0]
    start_r = int(start_cell[1:])
    # 指令公式映射表
    rule_map = {
        "s": lambda r: f"=SUM(B{r}:E{r})",
        "b": lambda r: f"=B{r}",
        "e": lambda r: f"=D{r}"
    }
    # 逐行批量写入
    for offset, cmd in enumerate(cmd_list):
        row = start_r + offset
        ws[f"{tar_col}{row}"].value = rule_map[cmd](row)
def FinalColumn2(start_cell, cmd_str: str, ws):
    # 拆分列标、起始行号
    tar_col = start_cell[0]
    start_r = int(start_cell[1:])
    # 指令公式映射表（和原版完全一致）
    rule_map = {
        "s": lambda r: f"=SUM(B{r}:E{r})",
        "b": lambda r: f"=B{r}",
        "e": lambda r: f"=D{r}"
    }
    # 字符串转字符序列，逐行批量写入
    for offset, cmd in enumerate(cmd_str):
        row = start_r + offset
        ws[f"{tar_col}{row}"].value = rule_map[cmd](row)
def insert_calc_row_at(ws, start_cell: str, row_idx_list: list, copy_cols: int, operator: str):
    """
    自定义指定若干行进行四则运算，向右批量生成相对引用公式
    :param ws: 工作表对象
    :param start_cell: 起始单元格坐标，如"B4"
    :param row_idx_list: 参与运算的行号列表，如[2,3]、[2,3,5]
                         +/*：列表长度≥2任意；-//：仅允许长度严格=2
    :param copy_cols: 向右额外复制生成的单元格数量（不含起始单元格）
    :param operator: 运算符，可选 "+" "-" "*" "/"
    """
    valid_ops = ("+", "-", "*", "/")
    if operator not in valid_ops:
        raise ValueError(f"运算符仅支持：{valid_ops}")

    len_rows = len(row_idx_list)
    if len_rows < 2:
        raise ValueError("参与运算的行号列表至少需要包含2个行号")

    # 减法、除法强制只能两个单元格运算
    if operator in ("-", "/") and len_rows != 2:
        raise ValueError(f"减法、除法仅支持2个单元格运算，列表长度必须=2，当前传入长度{len_rows}")

    # 解析起始单元格
    col_str = ""
    row_str = ""
    for ch in start_cell:
        if ch.isalpha():
            col_str += ch
        else:
            row_str += ch
    target_row = int(row_str)
    start_col_idx = column_index_from_string(col_str)

    # 逐列循环批量生成公式
    for col_offset in range(copy_cols + 1):
        curr_col_idx = start_col_idx + col_offset
        curr_col_let = get_column_letter(curr_col_idx)
        target_cell_addr = f"{curr_col_let}{target_row}"

        # 拼装参与运算的单元格地址
        cell_addrs = [f"{curr_col_let}{r}" for r in row_idx_list]

        if operator == "+":
            # SUM 区域求和
            min_r, max_r = min(row_idx_list), max(row_idx_list)
            formula = f"=SUM({curr_col_let}{min_r}:{curr_col_let}{max_r})"
        elif operator == "*":
            # PRODUCT 区域连乘
            min_r, max_r = min(row_idx_list), max(row_idx_list)
            formula = f"=PRODUCT({curr_col_let}{min_r}:{curr_col_let}{max_r})"
        elif operator in ("-", "/"):
            # 两格直接运算符拼接：第一个单元格 运算符 第二个单元格
            formula = f"={cell_addrs[0]}{operator}{cell_addrs[1]}"

        ws[target_cell_addr].value = formula
def insert_calc_row_at2(ws, start_cell: str, expr_str: str, copy_cols: int):
    """
    自定义行号四则运算表达式，向右批量生成相对引用Excel公式
    :param ws: 工作表对象
    :param start_cell: 起始单元格坐标，如"B4"
    :param expr_str: 行号运算表达式字符串，形如"2+3"、"2+3-4"、"2*5"、"7-2"
                     支持混合运算：仅同级依次运算，不支持括号优先级，严格从左往右计算
                     连续加号自动转为SUM区域，连续乘号自动转为PRODUCT区域
    :param copy_cols: 向右额外复制生成的单元格数量（不含起始单元格）
    """
    # 正则拆分：拆分出行号 + 运算符交替列表
    # 匹配数字和+-*/，提取token
    tokens = re.findall(r'(\d+|[+\-*/])', expr_str.strip())
    if not tokens:
        raise ValueError("运算表达式不能为空")
    
    # 校验首尾不能是运算符
    if tokens[0] in "+-*/" or tokens[-1] in "+-*/":
        raise ValueError(f"表达式{expr_str}首尾不能是运算符，格式非法")
    
    # 分离行号列表、运算符列表
    row_nums = []
    ops = []
    for idx, token in enumerate(tokens):
        if idx % 2 == 0:
            # 偶数位：行号，转整数
            if not token.isdigit():
                raise ValueError(f"位置{idx+1}必须是行数字，非法字符：{token}")
            row_nums.append(int(token))
        else:
            # 奇数位：运算符
            if token not in ("+", "-", "*", "/"):
                raise ValueError(f"仅支持 + - * / 运算符，非法符号：{token}")
            ops.append(token)
    
    if len(row_nums) < 2:
        raise ValueError("表达式至少需要2个行号参与运算，例如'2+3'")
    
    # 解析起始单元格
    col_str = ""
    row_str = ""
    for ch in start_cell:
        if ch.isalpha():
            col_str += ch
        else:
            row_str += ch
    target_row = int(row_str)
    start_col_idx = column_index_from_string(col_str)

    # 逐列循环批量生成公式
    for col_offset in range(copy_cols + 1):
        curr_col_idx = start_col_idx + col_offset
        curr_col_let = get_column_letter(curr_col_idx)
        target_cell_addr = f"{curr_col_let}{target_row}"
        
        # 组装当前列所有行单元格地址
        cell_addrs = [f"{curr_col_let}{r}" for r in row_nums]
        
        # 场景1：全部运算符都是+，用SUM区域公式
        if all(op == "+" for op in ops):
            min_r, max_r = min(row_nums), max(row_nums)
            formula = f"=SUM({curr_col_let}{min_r}:{curr_col_let}{max_r})"
        # 场景2：全部运算符都是*，用PRODUCT区域公式
        elif all(op == "*" for op in ops):
            min_r, max_r = min(row_nums), max(row_nums)
            formula = f"=PRODUCT({curr_col_let}{min_r}:{curr_col_let}{max_r})"
        # 场景3：混合+-*/，逐单元格拼接公式，从左到右依次运算
        else:
            formula_parts = [cell_addrs[0]]
            for op, cell in zip(ops, cell_addrs[1:]):
                formula_parts.append(op)
                formula_parts.append(cell)
            formula = f"={' '.join(formula_parts)}"
        
        ws[target_cell_addr].value = formula

row_store = {}

def insert_calc_row_at3(ws, start_cell: str, expr_str: str, copy_cols: int, alias: str = None):
    """
    自定义行号四则运算表达式，向右批量生成相对引用Excel公式
    :param ws: 工作表对象
    :param start_cell: 起始单元格坐标，如"B4"
    :param expr_str: 行号运算表达式字符串，形如"2+3"、"2+3-4"、"2*5"、"7-2"
                     支持混合运算：仅同级依次运算，不支持括号优先级，严格从左往右计算
                     连续加号自动转为SUM区域，连续乘号自动转为PRODUCT区域
    :param copy_cols: 向右额外复制生成的单元格数量（不含起始单元格）
    :param alias: 自定义别名，填入后本行单元格地址列表存入row_store[别名]，后续直接调取
    :return: 生成的单元格地址列表
    """
    tokens = re.findall(r'(\d+|[+\-*/])', expr_str.strip())
    if not tokens:
        raise ValueError("运算表达式不能为空")

    if tokens[0] in "+-*/" or tokens[-1] in "+-*/":
        raise ValueError(f"表达式{expr_str}首尾不能是运算符，格式非法")

    row_nums = []
    ops = []
    for idx, token in enumerate(tokens):
        if idx % 2 == 0:
            if not token.isdigit():
                raise ValueError(f"位置{idx+1}必须是行数字，非法字符：{token}")
            row_nums.append(int(token))
        else:
            if token not in ("+", "-", "*", "/"):
                raise ValueError(f"仅支持 + - * / 运算符，非法符号：{token}")
            ops.append(token)

    if len(row_nums) < 2:
        raise ValueError("表达式至少需要2个行号参与运算，例如'2+3'")

    # 解析起始单元格
    col_str = ""
    row_str = ""
    for ch in start_cell:
        if ch.isalpha():
            col_str += ch
        else:
            row_str += ch
    target_row = int(row_str)
    start_col_idx = column_index_from_string(col_str)

    cell_addr_list = []
    # 逐列批量写入公式
    for col_offset in range(copy_cols + 1):
        curr_col_idx = start_col_idx + col_offset
        curr_col_let = get_column_letter(curr_col_idx)
        target_cell_addr = f"{curr_col_let}{target_row}"
        cell_addr_list.append(target_cell_addr)

        cell_addrs = [f"{curr_col_let}{r}" for r in row_nums]

        if all(op == "+" for op in ops):
            min_r, max_r = min(row_nums), max(row_nums)
            formula = f"=SUM({curr_col_let}{min_r}:{curr_col_let}{max_r})"
        elif all(op == "*" for op in ops):
            min_r, max_r = min(row_nums), max(row_nums)
            formula = f"=PRODUCT({curr_col_let}{min_r}:{curr_col_let}{max_r})"
        else:
            formula_parts = [cell_addrs[0]]
            for op, cell in zip(ops, cell_addrs[1:]):
                formula_parts.append(op)
                formula_parts.append(cell)
            formula = f"={''.join(formula_parts)}"

        ws[target_cell_addr].value = formula

    # 存入托管字典
    if alias is not None:
        if alias in row_store:
            raise KeyError(f"别名「{alias}」已存在，禁止重复覆盖")
        row_store[alias] = cell_addr_list

    return cell_addr_list
def write_row_start_at2(ws, data_dict: dict):
    """
    从指定起始单元格向右横向逐单元格写入列表数据
    :param ws: openpyxl工作表worksheet对象
    :param data_dict: 字典，key为起始单元格字符串如"B2"，value为一维列表，列表元素依次向右填充
    """
    for start_cell, value_list in data_dict.items():
        # 拆分起始单元格列字母、行号
        col_chars = ""
        row_num_str = ""
        for ch in start_cell:
            if ch.isalpha():
                col_chars += ch
            else:
                row_num_str += ch
        start_col = column_index_from_string(col_chars)
        target_row = int(row_num_str)

        # 遍历列表，逐个向右写入
        for offset, val in enumerate(value_list):
            curr_col = start_col + offset
            curr_cell_addr = f"{get_column_letter(curr_col)}{target_row}"
            ws[curr_cell_addr].value = val
import re
from openpyxl.utils import column_index_from_string, get_column_letter

import re
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

def insert_calc_row_at4(ws, start_cell: str, expr_str: str, copy_cols: int, temp_save_path: str = "temp_calc_temp.xlsx"):
    """
    自定义行号四则运算表达式，向右批量生成相对引用Excel公式
    默认直接返回公式计算完成后的数值列表，无需额外传参
    :param ws: 工作表对象
    :param start_cell: 起始单元格坐标，如"B4"
    :param expr_str: 行号运算表达式字符串，形如"2+3"、"2+3-4"、"2*5"、"7-2"
                     支持混合运算：仅同级依次运算，不支持括号优先级，严格从左往右计算
                     连续加号自动转为SUM区域，连续乘号自动转为PRODUCT区域
    :param copy_cols: 向右额外复制生成的单元格数量（不含起始单元格）
    :param temp_save_path: 临时Excel中转文件路径，不用修改
    :return: 一行单元格Excel公式计算后的纯数值列表（float/int）
    """
    tokens = re.findall(r'(\d+|[+\-*/])', expr_str.strip())
    if not tokens:
        raise ValueError("运算表达式不能为空")

    if tokens[0] in "+-*/" or tokens[-1] in "+-*/":
        raise ValueError(f"表达式{expr_str}首尾不能是运算符，格式非法")

    row_nums = []
    ops = []
    for idx, token in enumerate(tokens):
        if idx % 2 == 0:
            if not token.isdigit():
                raise ValueError(f"位置{idx+1}必须是行数字，非法字符：{token}")
            row_nums.append(int(token))
        else:
            if token not in ("+", "-", "*", "/"):
                raise ValueError(f"仅支持 + - * / 运算符，非法符号：{token}")
            ops.append(token)

    if len(row_nums) < 2:
        raise ValueError("表达式至少需要2个行号参与运算，例如'2+3'")

    # 解析起始单元格行列
    col_str = ""
    row_str = ""
    for ch in start_cell:
        if ch.isalpha():
            col_str += ch
        else:
            row_str += ch
    target_row = int(row_str)
    start_col_idx = column_index_from_string(col_str)

    cell_addr_list = []
    # 批量写入公式
    for col_offset in range(copy_cols + 1):
        curr_col_idx = start_col_idx + col_offset
        curr_col_let = get_column_letter(curr_col_idx)
        target_cell_addr = f"{curr_col_let}{target_row}"
        cell_addr_list.append(target_cell_addr)

        cell_addrs = [f"{curr_col_let}{r}" for r in row_nums]

        if all(op == "+" for op in ops):
            min_r, max_r = min(row_nums), max(row_nums)
            formula = f"=SUM({curr_col_let}{min_r}:{curr_col_let}{max_r})"
        elif all(op == "*" for op in ops):
            min_r, max_r = min(row_nums), max(row_nums)
            formula = f"=PRODUCT({curr_col_let}{min_r}:{curr_col_let}{max_r})"
        else:
            formula_parts = [cell_addrs[0]]
            for op, cell in zip(ops, cell_addrs[1:]):
                formula_parts.append(op)
                formula_parts.append(cell)
            formula = f"={''.join(formula_parts)}"

        ws[target_cell_addr].value = formula

    # 保存工作簿，重载读取计算后的数值
    wb = ws.parent
    wb.save(temp_save_path)
    wb_read = load_workbook(temp_save_path, data_only=True)
    ws_read = wb_read[ws.title]

    value_result = [ws_read[addr].value for addr in cell_addr_list]
    return value_result
import re
import numpy as np
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

def insert_calc_row_at5(ws, start_cell: str, expr_str: str, copy_cols: int):
    """
    升级版：支持同列运算、跨工作表!引用、常量数字混合运算；自动向右批量生成公式，默认返回numpy数组
    :param ws: 当前工作表worksheet对象
    :param start_cell: 写入公式的起始单元格，例"B4"
    :param expr_str: 运算表达式，支持多种格式：
        1. 同行号运算："2+3-4" 同列，取当前列第2、3、4行
        2. 跨表引用："销售预算!B2+3" 固定取「销售预算」工作表B2单元格 + 当前列第3行
        3. 常量混合："2+100" 当前列2行 + 固定数值100
    :param copy_cols: 向右额外生成单元格数量（不含起始单元格）
    :return: np.ndarray 数值向量，可直接arr1+arr2*arr3向量运算
    """
    # 拆分表达式 token：单元格/表!单元格 / 数字常量 / 运算符
    token_pat = re.compile(r'(?P<ref>[^+\-*/]+)|(?P<op>[+\-*/])')
    raw_tokens = token_pat.findall(expr_str.strip())
    tokens = []
    for ref_part, op_part in raw_tokens:
        if ref_part:
            tokens.append(ref_part.strip())
        if op_part:
            tokens.append(op_part.strip())

    if not tokens:
        raise ValueError("运算表达式不能为空")
    if tokens[0] in "+-*/" or tokens[-1] in "+-*/":
        raise ValueError(f"表达式{expr_str}首尾不能是运算符")

    elem_list = []
    op_list = []
    for idx, item in enumerate(tokens):
        if idx % 2 == 0:
            elem_list.append(item)
        else:
            if item not in "+-*/":
                raise ValueError(f"非法运算符 {item}")
            op_list.append(item)

    if len(elem_list) < 2:
        raise ValueError("至少2个运算项，示例：'2+3'、'表1!B2+5'")

    # 解析目标写入单元格位置
    col_str, row_str = "", ""
    for ch in start_cell:
        if ch.isalpha():
            col_str += ch
        else:
            row_str += ch
    target_row = int(row_str)
    start_col_idx = column_index_from_string(col_str)

    cell_addr_list = []

    # 逐列向右批量生成公式
    for col_offset in range(copy_cols + 1):
        curr_col_idx = start_col_idx + col_offset
        curr_col_let = get_column_letter(curr_col_idx)
        write_cell = f"{curr_col_let}{target_row}"
        cell_addr_list.append(write_cell)

        formula_items = []
        for elem in elem_list:
            if "!" in elem:
                # 跨工作表：表名!单元格，固定单元格，不随列偏移
                sheet_name, cell_loc = elem.split("!", 1)
                formula_items.append(f"'{sheet_name}'!{cell_loc}")
            elif elem.isdigit():
                # 纯行号：当前偏移列 + 该行号
                r = int(elem)
                formula_items.append(f"{curr_col_let}{r}")
            else:
                # 纯数字常量
                try:
                    float(elem)
                    formula_items.append(elem)
                except ValueError:
                    raise ValueError(f"无法解析表达式项：{elem}")

        # 拼接完整Excel公式
        full_formula_parts = []
        full_formula_parts.append(formula_items[0])
        for op, it in zip(op_list, formula_items[1:]):
            full_formula_parts.append(op)
            full_formula_parts.append(it)
        formula = f"={''.join(full_formula_parts)}"
        ws[write_cell].value = formula

    # 临时保存读取计算结果
    wb = ws.parent
    temp_path = "_temp_calc.xlsx"
    wb.save(temp_path)
    wb_new = load_workbook(temp_path, data_only=True)
    ws_new = wb_new[ws.title]

    val_list = [ws_new[addr].value for addr in cell_addr_list]
    return np.array(val_list)
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import os

# 1、案例原始数值数据
case_data = {
    "receivables0": 52150,
    "Sales_amount": [22500, 50000, 30000, 25000, 15000, 28000],
    "price": [7.5, 7.5, 7.5, 7.5, 7.5, 7.5],
    "cash_sales_ratio": 0.7,
    "credit_sales_ratio": 0.25,

    "final_inventory_ratio": 0.2,
    "begin_inventory0": 3600,

    "unit_mat_consume": [5, 5, 5, 5, 5, 5],
    "Final_mat_ratio": 0.1,
    "begin_mat0": 13000,
    "unit_mat_price": [0.4, 0.4, 0.4, 0.4, 0.4, 0.4],
    "Payable0": 14200,
    "cash_purchase_ratio": 0.5,
    "credit_purchase_ratio": 0.5,

    "labor_hour_per_unit": [0.05, 0.05, 0.05, 0.05],
    "direct_labor_wage_rate": [10, 10, 10, 10],

    "var_mfg_rate": [20, 20, 20, 20],
    "other_fix_mfg": [9250, 9250, 8750, 9045],
    "mfg_fixed_asset0": 400000,
    "depreciation_rate": 0.01,
    "new": 50000,

    "sr": [0.5, 0.5, 0.5, 0.5],
    "other_fix_selling": [12500, 12500, 12500, 12500],
    "selling_asset": [52500, 52500, 52500, 52500],

    "begin_inventory_value": 18000,
    "tax_rate": 0.25,
    "begin_cash": 55000,
    "APR": 0.16,
    "long": [500000, 500000, 500000, 500000],
    "coInvest": [0, 0, 30000, 0],
    "coDividend": [84450, 0, 0, 0],
    "Sinterest_rate": 0.01,
    "kill_line": 30000,
    "excess_line": 80000,
    "coTax": [0, 0, 20000, 15000],
    "coCapitalEx": [0, 50000, 0, 0],
    "inv_mgmt_rate": [0.2, 0.2, 0.2, 0.2],
    "mgmt_asset": [122500, 122500, 122500, 122500],
    "other_fixed_mgmt": [20000, 20000, 20000, 20000],
    "other_fixed_cost2": [20000, 20000, 20000, 20000],

    "current_assets0": 130350,
    "cash_cash_equivalents0": 55000,
    "accounts_receivable0": 52150,
    "raw_material_inventory0": 5200,
    "finished_goods_inventory0": 18000,
    "non_current_assets0": 745000,
    "land_use_rights0": 150000,
    "fixed_assets0": 575000,
    "production_fixed_assets0": 400000,
    "sales_admin_fixed_assets0": 175000,
    "long_term_equity_investment0": 20000,
    "total_assets0": 875350,
    "current_liabilities0": 29200,
    "accounts_payable0": 14200,
    "taxes_payable0": 15000,
    "short_term_loan0": 0,
    "non_current_liabilities0": 500000,
    "long_term_loan0": 500000,
    "total_equity0": 346150,
    "common_stock0": 200000,
    "retained_earnings0": 146150
}

# 变量-中文注释映射表
var_cn_map = {
    "receivables0": "期初应收账款",
    "Sales_amount": "各期销售额",
    "price": "产品单价",
    "cash_sales_ratio": "现销比例",
    "credit_sales_ratio": "赊销比例",
    "final_inventory_ratio": "期末产成品库存占销量比例",
    "begin_inventory0": "期初产成品库存量",
    "unit_mat_consume": "单位产品材料消耗量",
    "Final_mat_ratio": "期末原材料库存占耗用量比例",
    "begin_mat0": "期初原材料库存量",
    "unit_mat_price": "原材料单价",
    "Payable0": "期初应付账款",
    "cash_purchase_ratio": "现款采购比例",
    "credit_purchase_ratio": "赊购比例",
    "labor_hour_per_unit": "单位产品工时",
    "direct_labor_wage_rate": "直接人工小时工资率",
    "var_mfg_rate": "变动制造费用分配率",
    "other_fix_mfg": "其他固定制造费用",
    "mfg_fixed_asset0": "期初生产固定资产原值",
    "depreciation_rate": "折旧率",
    "new": "本期新增固定资产",
    "sr": "变动销售费用率",
    "other_fix_selling": "其他固定销售费用",
    "selling_asset": "销售部门固定资产",
    "begin_inventory_value": "期初产成品库存金额",
    "tax_rate": "所得税率",
    "begin_cash": "期初现金余额",
    "APR": "短期借款年利率",
    "long": "长期借款余额",
    "coInvest": "资本性投资支出",
    "coDividend": "现金股利发放额",
    "Sinterest_rate": "闲置资金理财收益率",
    "kill_line": "现金最低安全线",
    "excess_line": "现金闲置上限",
    "coTax": "实际缴纳所得税",
    "coCapitalEx": "资本开支",
    "inv_mgmt_rate": "变动管理费用率",
    "mgmt_asset": "管理用固定资产",
    "other_fixed_mgmt": "其他固定管理费用",
    "other_fixed_cost2": "其他固定经营费用",
    "current_assets0": "期初流动资产合计",
    "cash_cash_equivalents0": "期初货币资金",
    "accounts_receivable0": "期初应收账款",
    "raw_material_inventory0": "期初原材料库存金额",
    "finished_goods_inventory0": "期初产成品库存金额",
    "non_current_assets0": "期初非流动资产合计",
    "land_use_rights0": "土地使用权",
    "fixed_assets0": "固定资产合计原值",
    "production_fixed_assets0": "生产用固定资产",
    "sales_admin_fixed_assets0": "销售管理用固定资产",
    "long_term_equity_investment0": "长期股权投资",
    "total_assets0": "期初总资产",
    "current_liabilities0": "期初流动负债合计",
    "accounts_payable0": "期初应付账款",
    "taxes_payable0": "期初应交税费",
    "short_term_loan0": "期初短期借款",
    "non_current_liabilities0": "期初非流动负债合计",
    "long_term_loan0": "期初长期借款",
    "total_equity0": "期初所有者权益合计",
    "common_stock0": "实收资本/普通股股本",
    "retained_earnings0": "期初留存收益"
}

# 季度表头定义
q6_labels = ["Q1", "Q2", "Q3", "Q4", "次年Q1", "次年Q2"]
q4_labels = ["Q1", "Q2", "Q3", "Q4"]

# 全局变量：存储读取后的全部数据 + Excel保存文件夹路径
global_input_data = None
excel_save_folder_path = ""

class BudgetInputGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("全面预算编制Master Budgeting")
        self.root.geometry("1450x850")
        self.var_widgets = {}   # 存储每个变量对应的输入框列表/单个输入框
        # 新增状态标记
        self.data_verified = False    # 数据是否核验通过
        self.path_selected = False    # 是否选定保存路径

        # ========== 顶部操作指南滚动文本框 ==========
        guide_frame = ttk.LabelFrame(root, text="欢迎使用 & 操作使用指南（可自行编辑修改文字）")
        guide_frame.pack(fill="x", padx=12, pady=6)
        self.guide_text = scrolledtext.ScrolledText(guide_frame, height=10, wrap=tk.WORD)
        self.guide_text.pack(fill="both", expand=True, padx=5, pady=5)
        default_guide = """欢迎使用本全面预算编制系统！
【操作步骤】
1、点击【使用兴腾公司模板数据】，一键自动填充老师案例里的基本数据，如有兴趣可在此基础上修改；
2、全部参数修改完毕后，点击【核验数据并选择路径】，校验数据并选择报表文件保存路径；
3、校验和路径选择完成后，点击【点击生成报表】，程序自动关闭录入窗口并批量生成全套Excel报表；
"""
        self.guide_text.insert(tk.END, default_guide)

        # 顶部按钮栏（修改按钮文字+新增第三个按钮）
        btn_frame = ttk.Frame(root)
        btn_frame.pack(fill="x", padx=12, pady=6)
        ttk.Button(btn_frame, text="使用兴腾公司模板数据", command=self.fill_case_data).pack(side="left", padx=10)
        # 第二个按钮改名
        self.btn_check = ttk.Button(btn_frame, text="核验数据并选择路径", command=self.read_all_data)
        self.btn_check.pack(side="left", padx=10)
        # 新增第三个生成报表按钮
        self.btn_gen_report = ttk.Button(btn_frame, text="点击生成报表", command=self.gen_report_btn_click)
        self.btn_gen_report.pack(side="left", padx=10)

        # 滚动画布
        canvas = tk.Canvas(root)
        scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(12,0))
        scrollbar.pack(side="right", fill="y")
        self.main_frame = scroll_frame

        # 模块分组
        self.groups = [
            ("一、销售&应收模块", ["receivables0", "Sales_amount", "price", "cash_sales_ratio", "credit_sales_ratio"]),
            ("二、产成品库存模块", ["final_inventory_ratio", "begin_inventory0"]),
            ("三、原材料&应付模块", ["unit_mat_consume", "Final_mat_ratio", "begin_mat0", "unit_mat_price", "Payable0", "cash_purchase_ratio", "credit_purchase_ratio"]),
            ("四、直接人工模块", ["labor_hour_per_unit", "direct_labor_wage_rate"]),
            ("五、制造费用&折旧模块", ["var_mfg_rate", "other_fix_mfg", "mfg_fixed_asset0", "depreciation_rate", "new"]),
            ("六、销售费用模块", ["sr", "other_fix_selling", "selling_asset"]),
            ("七、税费&现金流参数", ["begin_inventory_value", "tax_rate", "begin_cash", "APR", "long", "coInvest", "coDividend", "Sinterest_rate", "kill_line", "excess_line", "coTax", "coCapitalEx", "inv_mgmt_rate", "mgmt_asset", "other_fixed_mgmt", "other_fixed_cost2"]),
            ("八、期初资产负债表数据", ["current_assets0", "cash_cash_equivalents0", "accounts_receivable0", "raw_material_inventory0", "finished_goods_inventory0", "non_current_assets0", "land_use_rights0", "fixed_assets0", "production_fixed_assets0", "sales_admin_fixed_assets0", "long_term_equity_investment0", "total_assets0", "current_liabilities0", "accounts_payable0", "taxes_payable0", "short_term_loan0", "non_current_liabilities0", "long_term_loan0", "total_equity0", "common_stock0", "retained_earnings0"])
        ]
        self.build_all_groups()

    def build_all_groups(self):
        row_idx = 0
        for group_title, var_list in self.groups:
            group_box = ttk.LabelFrame(self.main_frame, text=group_title)
            group_box.grid(row=row_idx, column=0, columnspan=12, sticky="ew", padx=8, pady=8)
            inner_row = 0
            row_idx += 1

            for var_name in var_list:
                val_sample = case_data[var_name]
                disp_label = f"{var_name}【{var_cn_map[var_name]}】"

                if not isinstance(val_sample, list):
                    ttk.Label(group_box, text=disp_label, width=30).grid(row=inner_row, column=0, sticky="w", padx=5, pady=3)
                    ent = ttk.Entry(group_box, width=18)
                    ent.grid(row=inner_row, column=1, padx=5, pady=3)
                    self.var_widgets[var_name] = ent
                    inner_row += 1
                    continue

                length = len(val_sample)
                ent_list = []
                if length == 6:
                    q_labels = q6_labels
                elif length == 4:
                    q_labels = q4_labels
                else:
                    ttk.Label(group_box, text=disp_label, width=30).grid(row=inner_row, column=0, sticky="w", padx=5, pady=3)
                    ent = ttk.Entry(group_box, width=18)
                    ent.grid(row=inner_row, column=1, padx=5, pady=3)
                    self.var_widgets[var_name] = ent
                    inner_row += 1
                    continue

                ttk.Label(group_box, text=disp_label, width=30).grid(row=inner_row, column=0, sticky="w", padx=5, pady=3)
                for col, qtxt in enumerate(q_labels):
                    ttk.Label(group_box, text=qtxt, width=12).grid(row=inner_row, column=col+1, padx=2, pady=2)
                inner_row += 1
                for col in range(length):
                    ent_q = ttk.Entry(group_box, width=12)
                    ent_q.grid(row=inner_row, column=col+1, padx=2, pady=3)
                    ent_list.append(ent_q)
                self.var_widgets[var_name] = ent_list
                inner_row += 1

    def fill_case_data(self):
        """一键填充案例数据到对应季度输入框"""
        for var_name, val in case_data.items():
            widget = self.var_widgets[var_name]
            if isinstance(widget, list):
                for idx, ent in enumerate(widget):
                    ent.delete(0, tk.END)
                    ent.insert(0, str(val[idx]))
            else:
                widget.delete(0, tk.END)
                widget.insert(0, str(val))
        messagebox.showinfo("填充完成", "案例数据已自动填入各季度输入框！")
        # 重新填充数据后重置核验标记
        self.data_verified = False
        self.path_selected = False

    def read_all_data(self):
        global global_input_data, excel_save_folder_path
        out_data = {}
        try:
            for var_name, widget in self.var_widgets.items():
                if isinstance(widget, list):
                    num_list = []
                    for ent in widget:
                        txt = ent.get().strip()
                        if not txt:
                            raise Exception(f"{var_name} 存在空季度输入")
                        num_list.append(float(txt))
                    out_data[var_name] = num_list
                else:
                    txt = widget.get().strip()
                    if not txt:
                        raise Exception(f"{var_name} 不能为空")
                    out_data[var_name] = float(txt)
        except Exception as e:
            messagebox.showerror("输入校验失败", f"解析错误：{str(e)}")
            self.data_verified = False
            self.path_selected = False
            return

        # 全局注入变量
        global_input_data = out_data
        globals().update(out_data)
        self.data_verified = True

        # 选择保存文件夹
        selected_folder = filedialog.askdirectory(title="请选择预算Excel报表存放文件夹")
        if not selected_folder:
            messagebox.showwarning("提示", "未选择保存目录，后续无法导出Excel")
            excel_save_folder_path = ""
            self.path_selected = False
            return
        excel_save_folder_path = selected_folder
        self.path_selected = True

        msg = f"数据校验全部通过！\n选定保存路径：{excel_save_folder_path}\n\n可点击【点击生成报表】按钮自动生成全套报表！"
        messagebox.showinfo("操作完成", msg)
        print("=== 变量校验成功 ===")
        print("Sales_amount =", Sales_amount)
        print("Excel保存目录：", excel_save_folder_path)

    def gen_report_btn_click(self):
        """第三个按钮：生成报表点击逻辑"""
        # 校验前置条件
        if not self.data_verified:
            messagebox.showerror("操作拦截", "尚未核验数据！请先点击【核验数据并选择路径】完成数据校验！")
            return
        if not self.path_selected:
            messagebox.showerror("操作拦截", "数据已核验，但未选择Excel保存文件夹！请重新执行【核验数据并选择路径】选定目录！")
            return
        # 全部正常，关闭GUI窗口，进入后续报表生成
        messagebox.showinfo("即将生成报表", "录入窗口即将关闭，自动开始生成全套预算Excel报表！")
        self.root.destroy()

# ==================== 主程序入口 ====================
if __name__ == "__main__":
    root = tk.Tk()
    app = BudgetInputGUI(root)
    root.mainloop()

    # 窗口关闭后自动执行你的Excel生成代码
    if excel_save_folder_path and global_input_data is not None:
        print("\n录入窗口已关闭，开始自动生成Excel报表……")
        # ========== 你的原有wb保存代码直接粘贴到下方 ==========
        # full_path = os.path.join(excel_save_folder_path, "销售预算.xlsx")
        # wb.save(full_path)
        # ======================================================
        messagebox.showinfo("报表生成完毕", f"Excel报表已保存至：\n{excel_save_folder_path}")
#基本数据总结
# receivables0=52150
# Sales_amount=[22500,50000,30000,25000,15000,28000]
# price=[7.5,7.5,7.5,7.5,7.5,7.5]
# cash_sales_ratio = 0.7
# credit_sales_ratio = 0.25
# #-----------------------------------------------
# final_inventory_ratio=0.2
# begin_inventory0=3600
# #---------------------------------------
# unit_mat_consume = [5, 5, 5, 5, 5,5]
# Final_mat_ratio=0.1
# begin_mat0=13000
# unit_mat_price=[0.4,0.4,0.4,0.4,0.4,0.4]
# Payable0=14200
# cash_purchase_ratio=0.5
# credit_purchase_ratio=0.5
# #------------------------------------------
# labor_hour_per_unit = [0.05, 0.05, 0.05, 0.05]
# direct_labor_wage_rate = [10, 10, 10, 10]
# #--------------------------------------------
# var_mfg_rate = [20, 20, 20, 20]
# other_fix_mfg = [9250, 9250, 8750, 9045]
# mfg_fixed_asset0=400000
# depreciation_rate=0.01
# new=50000
# #----------------------------------------------
# sr=[0.5, 0.5, 0.5, 0.5]#变动销售费用率
# other_fix_selling=[12500,12500,12500,12500]
# selling_asset=[52500,52500,52500,52500]
# #-----------------------------------------
# begin_inventory_value=18000
# tax_rate=0.25
# begin_cash=55000
# APR=0.16
# long=[500000,500000,500000,500000]
# coInvest=[0,0,30000,0]
# coDividend=[84450,0,0,0]
# Sinterest_rate=0.01
# kill_line=30000
# excess_line=80000
# coTax=[0,0,20000,15000]
# coCapitalEx=[0,50000,0,0]
# inv_mgmt_rate = [0.2, 0.2, 0.2, 0.2]
# mgmt_asset=[122500,122500,122500,122500]
# other_fixed_mgmt=[20000,20000,20000,20000]
# other_fixed_cost2 = [20000, 20000, 20000, 20000]
# #------------------------------------------------

# current_assets0 = 130350
# cash_cash_equivalents0 = 55000
# accounts_receivable0 = 52150
# raw_material_inventory0 = 5200
# finished_goods_inventory0 = 18000
# non_current_assets0 = 745000
# land_use_rights0 = 150000
# fixed_assets0 = 575000
# production_fixed_assets0 = 400000
# sales_admin_fixed_assets0 = 175000
# long_term_equity_investment0 = 20000
# total_assets0 = 875350
# current_liabilities0 = 29200
# accounts_payable0 = 14200
# taxes_payable0 = 15000
# short_term_loan0 = 0
# non_current_liabilities0 = 500000
# long_term_loan0 = 500000
# total_equity0 = 346150
# common_stock0 = 200000
# retained_earnings0 = 146150
#Sales Budget
wb = Workbook()
ws = wb.active
ws.title = "销售预算表"
#表的格式
# 顶部年份大标题合并
ws.merge_cells("B1:E1")   # 2025年
ws.merge_cells("F1:G1")   # 2026年
ws.merge_cells("H1:H2")   # 2025合计
ws.merge_cells("A1:A2")   # A列标题区域
# 写入一级表头文字
ws["B1"] = "2025年"
ws["F1"] = "2026年"
ws["H1"] = "2025合计"
# 二级季度列表头（第2行）
col_headers = [
    "第1季度",
    "第2季度",
    "第3季度",
    "第4季度",
    "第1季度",
    "第2季度"
]
col_pos = ["B2","C2","D2","E2","F2","G2"]
for pos, txt in zip(col_pos, col_headers):
    ws[pos] = txt
# ---------------------- A列所有行项目文字（逐行对应截图） ----------------------
row_texts = [
    "销售数量",
    "销售价格",
    "营业收入",
    "销售活动现金预算",
    "上年末应收账款",
    "第1季度销售现金流入",
    "第1季度末应收账款",
    "第2季度销售现金流入",
    "第2季度末应收账款",
    "第3季度销售现金流入",
    "第3季度末应收账款",
    "第4季度销售现金流入",
    "第4季度末应收账款",
    "下年第1季度销售现金流入",
    "下年第1季度末应收账款",
    "下年第2季度销售现金流入",
    "销售活动现金流入合计"
]
# 从第3行开始依次填充A列行标题
start_row = 3
for idx, text in enumerate(row_texts):
    ws.cell(row=start_row + idx, column=1, value=text)
#数据运算——销售预算
Sales_revenue = (np.array(Sales_amount)*np.array(price)).tolist()
Sales_cash_inflows_current_period=((np.array(Sales_revenue))*cash_sales_ratio).tolist()
Sales_cash_inflows_last_period=(np.array([receivables0/0.25] + Sales_revenue[:-1])*credit_sales_ratio).tolist()
#写入
write_row_start_at2(ws,{"B3":Sales_amount,"B4":price,"B5":Sales_revenue})
start_row=7
start_col = 2
for i, last in enumerate(Sales_cash_inflows_last_period):
    ws.cell(row=start_row + i*2, column=start_col + i, value=last)
for j, cash in enumerate(Sales_cash_inflows_current_period):
    ws.cell(row=start_row + 1 + j*2, column=start_col + j, value=cash)
insert_calc_row_at2(ws,"B19","7+8+9+10+11+12+13+14+15+16+17+18",5)
cash_inflow=(np.array(Sales_cash_inflows_current_period)+np.array(Sales_cash_inflows_last_period)).tolist()
import os
FinalColumn2("H3","sbsbsssssssssssss",ws)
wb.save(os.path.join(excel_save_folder_path, "销售预算.xlsx"))
# --------------------------------------------
wb2 = Workbook()
ProductionBudget = wb2.active
ProductionBudget.title = "销售预算表"
# 第1行表头
header_row1 = ["", "2025年第1季度", "2025年第2季度", "2025年第3季度", "2025年第4季度", "2026年第1季度", "2026年第2季度", "2025年合计"]
ProductionBudget.append(header_row1)
# 第2~6行行标题
row_titles = [
    "",
    "销售数量",
    "加：期末产成品库存",
    "产成品需求总量",
    "减：期初产成品库存",
    "本月生产数量"
]
# 逐行写入，只填文字，数值单元格留空
for idx, title in enumerate(row_titles, start=2):
    ProductionBudget.cell(row=idx, column=1, value=title)
#Calculation
Final_inventory=(np.array(Sales_amount[1:]+[0])*final_inventory_ratio).tolist()
Initial_inventory=np.array([begin_inventory0] + Final_inventory[:-1]).tolist()
#Write
write_row_start_at2(ProductionBudget,{"B3":Sales_amount,"B4":Final_inventory,"B6":Initial_inventory})
insert_calc_row_at2(ProductionBudget,"B5",'3+4',5)
Finished_Goods_total_demand=(np.array(Sales_amount)+np.array(Final_inventory)).tolist()
# production_for_this_month=insert_calc_row_at4(ProductionBudget,"B7",'5-6',5)
Production_for_this_month=(np.array(Finished_Goods_total_demand)-np.array(Initial_inventory)).tolist()
write_row_start_at(ProductionBudget,"B7",Production_for_this_month)

FinalColumn2("H3",'sesbs',ProductionBudget)
wb2.save(os.path.join(excel_save_folder_path, "生产预算.xlsx"))

#-------------------------------------------------------------------
wb3 = Workbook()
MaterialBudget = wb3.active
MaterialBudget.title = "原材料预算表"

MaterialBudget["B1"] = "2025年第1季度"
MaterialBudget["C1"] = "2025年第2季度"
MaterialBudget["D1"] = "2025年第3季度"
MaterialBudget["E1"] = "2025年第4季度"
MaterialBudget["F1"] = "2026年第1季度"
MaterialBudget["G1"] = "2025年合计"

# A列行项目标题（严格对应截图行顺序，从第2行开始）
row_title_list = [
    "",
    "生产数量",
    "单位产品材料消耗",
    "材料消耗总量",
    "加：期末材料库存",
    "材料需求总量",
    "减：期初材料库存",
    "材料采购总量",
    "材料单位价格",
    "材料采购总成本",
    "材料采购现金预算",
    "2024年第4季度末应付账款",
    "2025年第1季度现金支出",
    "2025年第1季度末应付账款",
    "2025年第2季度现金支出",
    "2025年第2季度末应付账款",
    "2025年第3季度现金支出",
    "2025年第3季度末应付账款",
    "2025年第4季度现金支出",
    "2025年第4季度末应付账款",
    "材料采购现金支出合计"
]

for row_idx, title in enumerate(row_title_list, start=2):
    MaterialBudget.cell(row=row_idx, column=1, value=title)


total_mat_consume = (np.array(Production_for_this_month)*np.array(unit_mat_consume)).tolist()
Final_mat=(np.array(total_mat_consume[1:]+[0])*Final_mat_ratio).tolist()
total_mat_demand=(np.array(total_mat_consume)+np.array(Final_mat)).tolist()
Initial_mat=np.array([begin_mat0] + Final_mat[:-1]).tolist()
total_mat_purchase=(np.array(total_mat_demand)-np.array(Initial_mat)).tolist()
total_mat_purchase_cost=(np.array(total_mat_purchase)*np.array(unit_mat_price)).tolist()
Payable_from_current_period=(np.array(total_mat_purchase_cost)*cash_purchase_ratio).tolist()
Payable_from_last_period=(np.array([Payable0/0.5] + total_mat_purchase_cost[:-1])*credit_purchase_ratio).tolist()
cash_out_flows_due_to_mat_purchase=(np.array(Payable_from_last_period)+np.array(Payable_from_current_period))
#write
write_row_start_at2(MaterialBudget,{"B3":Production_for_this_month,
                                    "B4":unit_mat_consume,
                                    "B5":total_mat_consume,
                                    "B6":Final_mat,
                                    "B7":total_mat_demand,
                                    "B8":Initial_mat,
                                    "B9":total_mat_purchase,
                                    "B10":unit_mat_price,
                                    "B11":total_mat_purchase_cost,
                                    "B22":cash_out_flows_due_to_mat_purchase
                                    })
start_row=13
start_col = 2
for i, last in enumerate(Payable_from_last_period[:-1]):
    MaterialBudget.cell(row=start_row + i*2, column=start_col + i, value=last)
for j, cash in enumerate(Payable_from_current_period[:-1]):
    MaterialBudget.cell(row=start_row + 1 + j*2, column=start_col + j, value=cash)
FinalColumn2("G3",'sbsesbsbsbssssssssss',MaterialBudget)


wb3.save(os.path.join(excel_save_folder_path, "原材料预算.xlsx"))



wb4 = Workbook()
DirectLabor = wb4.active
DirectLabor.title = "直接人工预算表"
# 写入列标题行（第1行）
DirectLabor.append(["", "2025年第1季度", "2025年第2季度", "2025年第3季度", "2025年第4季度", "2025年合计"])

# 写入行标题行（第2-7行，数值列全留空）
row_headers = [
    "生产数量",
    "单位产品人工工时",
    "直接人工工时总量",
    "直接人工工资率",
    "直接人工工资总额",
    "直接人工现金支出"
]
for header in row_headers:
    DirectLabor.append([header, "", "", "", "", ""])

# 基础列宽适配（格式优化，无样式高亮）
DirectLabor.column_dimensions["A"].width = 24
for col in ["B", "C", "D", "E", "F"]:
    DirectLabor.column_dimensions[col].width = 20



total_labor_hour=(np.array(Production_for_this_month[:-2])*np.array(labor_hour_per_unit)).tolist()
direct_labor_total_cost=(np.array(total_labor_hour)*np.array(direct_labor_wage_rate)).tolist()

FinalColumn2("F2",'sbsbss',DirectLabor)
write_row_start_at2(DirectLabor,{"B2":Production_for_this_month[:-2],
                                    "B3":labor_hour_per_unit,
                                    "B4":total_labor_hour,
                                    "B5":direct_labor_wage_rate,
                                    "B6":direct_labor_total_cost,
                                    "B7":direct_labor_total_cost
                                    })


wb4.save(os.path.join(excel_save_folder_path, "直接人工预算.xlsx"))

wb5 = Workbook()
ME= wb5.active
ME.title = "直接人工预算表"

header = [
    "",
    "2025年第1季度",
    "2025年第2季度",
    "2025年第3季度",
    "2025年第4季度",
    "2025年合计"
]
ME.append(header)

# 写入行项目（第2-13行），数值列全留空、无高亮、无数字
item_rows = [
    ["生产工时", "", "", "", "", ""],
    ["变动制造费用率", "", "", "", "", ""],
    ["变动制造费用总额", "", "", "", "", ""],
    ["固定制造费用总额", "", "", "", "", ""],
    ["其中：折旧", "", "", "", "", ""],
    ["其他", "", "", "", "", ""],
    ["制造费用总金额", "", "", "", "", ""],
    ["变动制造费用分配率", "", "", "", "", ""],
    ["制造费用现金预算", "", "", "", "", ""],
    ["制造费用总额", "", "", "", "", ""],
    ["减：非付现制造费用", "", "", "", "", ""],
    ["制造费用现金支出", "", "", "", "", ""]
]
for row in item_rows:
    ME.append(row)
    #---------------

fixed_asset=[mfg_fixed_asset0,mfg_fixed_asset0+new,mfg_fixed_asset0+new,mfg_fixed_asset0+new]
#--------
var_mfg_cost=(np.array(var_mfg_rate)*np.array(total_labor_hour)).tolist()
depreciation_for_mfg_asset=(np.array(fixed_asset)*depreciation_rate*3).tolist()
fix_mfg_total = (np.array(depreciation_for_mfg_asset)+np.array(other_fix_mfg)).tolist()
total_mfg_cost=(np.array(fix_mfg_total)+np.array(var_mfg_cost)).tolist()
cash_out_flows_due_to_mfg=(np.array(total_mfg_cost)-np.array(depreciation_for_mfg_asset)).tolist()
FinalColumn2("F2",'sbsssssbbsss',ME)
write_row_start_at2(ME,{"B2":total_labor_hour,
                                    "B3":var_mfg_rate,
                                    "B4":var_mfg_cost,
                                    "B5":fix_mfg_total,
                                    "B6":depreciation_for_mfg_asset,
                                    "B7":other_fix_mfg,
                                    "B8":total_mfg_cost,
                                    "B11":total_mfg_cost,
                                    "B12":depreciation_for_mfg_asset,
                                    "B13":cash_out_flows_due_to_mfg
                                    })

wb5.save(os.path.join(excel_save_folder_path, "制造费用预算.xlsx"))

#----------------------------------------------------------------
wb6 = Workbook()
UC= wb6.active
UC.title = "单位成本"
UC["A1"] = ""
UC["B1"] = "资源消耗量"
UC["C1"] = "单位资源成本"
UC["D1"] = "单项成本"
UC["A2"] = "直接材料"
UC["A3"] = "直接人工"
UC["A4"] = "变动制造费用"
UC["A5"] = "合计"
UC["B2"]=unit_mat_consume[0]
UC["C2"]=unit_mat_price[0]
UC["D2"]="=B2*C2"
UC["B3"]=labor_hour_per_unit[0]
UC["C3"]=direct_labor_wage_rate[0]
UC["D3"]="=B3*C3"
UC["B4"]=labor_hour_per_unit[0]
UC["C4"]=var_mfg_rate[0]
UC["D4"]="=B4*C4"
unit_inventory_cost=unit_mat_consume[0]*unit_mat_price[0]+labor_hour_per_unit[0]*direct_labor_wage_rate[0]+labor_hour_per_unit[0]*var_mfg_rate[0]
UC["D5"]=unit_inventory_cost
UC["A10"] = "单位标准成本"
UC["B10"] = "5.2"
wb6.save(os.path.join(excel_save_folder_path, "单位成本预算.xlsx"))
#---------------------------------------------------------



wb7 = Workbook()
FI= wb7.active
FI.title = "单位成本"
num=Final_mat[3]
unit_cost=unit_mat_price[0]
NUM=Final_inventory[3]
UNIT_COST=unit_inventory_cost
FI.append(["", "数量", "单位成本", "总成本"])
FI.append(["原材料", num, unit_cost, num*unit_cost])
FI.append(["产成品", NUM, UNIT_COST, NUM*UNIT_COST])
cost_of_final_inventory=num*unit_cost+NUM*UNIT_COST
FI.append(["合计", None, None, cost_of_final_inventory])
wb7.save(os.path.join(excel_save_folder_path, "期末存货预算.xlsx"))


#--------------------------------------------------------
wb8 = Workbook()
SE= wb8.active
SE.title = "销售费用预算"
# 第1行表头
SE["A1"] = ""
SE["B1"] = "2025年第1季度"
SE["C1"] = "2025年第2季度"
SE["D1"] = "2025年第3季度"
SE["E1"] = "2025年第4季度"
SE["F1"] = "2025年合计"

# A列行标题，仅写入文字，B-F列全部留白，无数字、无高亮
SE["A2"] = "销售数量"
SE["A3"] = "变动销售费用率"
SE["A4"] = "变动销售费用总额"
SE["A5"] = "固定销售费用总额"
SE["A6"] = "其中：折旧"
SE["A7"] = "其他"
SE["A8"] = "销售费用总金额"
SE["A9"] = "销售费用现金预算"
SE["A10"] = "销售费用总额"
SE["A11"] = "减：非付现销售费用"
SE["A12"] = "销售费用现金支出"

var_selling_expense=(np.array(sr)*np.array(Sales_amount[:-2])).tolist()

depreciation_for_selling=(np.array(selling_asset)*depreciation_rate*3).tolist()

fix_selling_total=(np.array(other_fix_selling)+np.array(depreciation_for_selling)).tolist()
total_selling_expense=(np.array(fix_selling_total)+np.array(var_selling_expense)).tolist()
cash_out_flows_due_to_selling=(np.array(total_selling_expense)-np.array(depreciation_for_selling)).tolist()
FinalColumn2("F2",'sbsssssbsss',SE)
write_row_start_at2(SE,{"B2":Sales_amount[:-2],
                                    "B3":sr,
                                    "B4":var_selling_expense,
                                    "B5":fix_selling_total,
                                    "B6":depreciation_for_selling,
                                    "B7":other_fix_selling,
                                    "B8":total_selling_expense,
                                    "B10":total_selling_expense,
                                    "B11":depreciation_for_selling,
                                    "B12":cash_out_flows_due_to_selling
                                    })
wb8.save(os.path.join(excel_save_folder_path, "销售费用预算.xlsx"))



wb9 = Workbook()
AE= wb9.active
AE.title = "管理费用预算"
from openpyxl import Workbook
wb9 = Workbook()
AE = wb9.active
AE.title = "管理费用预算"
# 硬写入A列行标题
AE['A1'] = ""
AE['A2'] = "原材料存货价值"
AE['A3'] = "产成品存货价值"
AE['A4'] = "存货总价值"
AE['A5'] = "存货管理费用率"
AE['A6'] = "变动管理费用总额"
AE['A7'] = "固定管理费用总额"
AE['A8'] = "其中：折旧"
AE['A9'] = "坏账"
AE['A10'] = "其他"
AE['A11'] = "管理费用总金额"
AE['A12'] = "管理费用现金预算"
AE['A13'] = "管理费用总额"
AE['A14'] = "减：非付现管理费用"
AE['A15'] = "减：坏账"
AE['A16'] = "管理费用现金支出"
# 硬写入第1行列标题
AE['B1'] = "2025年第1季度"
AE['C1'] = "2025年第2季度"
AE['D1'] = "2025年第3季度"
AE['E1'] = "2025年第4季度"
AE['F1'] = "2025年合计"

bad_debt_ratio=1-cash_sales_ratio-credit_sales_ratio
bad_debt = (np.array(Sales_revenue[:-2])*bad_debt_ratio).tolist()
raw_material_value=(np.array(Final_mat[:-2])*np.array(unit_mat_price[:-2])).tolist()
inventory_value=(np.array(Final_inventory[:-2])*np.array(unit_inventory_cost)).tolist()
total_inventory_value=(np.array(raw_material_value)+np.array(inventory_value)).tolist()
var_mgmt_cost = (np.array(inv_mgmt_rate)*total_inventory_value).tolist()
depreciation_for_mgmt=(np.array(mgmt_asset)*depreciation_rate*3).tolist()
fix_mgmt_cost=(np.array(depreciation_for_mgmt)+np.array(bad_debt)+np.array(other_fixed_mgmt)).tolist()
mgmt_total=(np.array(var_mgmt_cost)+np.array(fix_mgmt_cost)).tolist()
cash_out_flows_due_to_mgmt=(np.array(mgmt_total)-np.array(depreciation_for_mgmt)-np.array(bad_debt))
write_row_start_at2(AE,{"B2":raw_material_value,
                                    "B3":inventory_value,
                                    "B4":total_inventory_value,
                                    "B5":inv_mgmt_rate,
                                    "B6":var_mgmt_cost,
                                    "B7":fix_mgmt_cost,
                                    "B8":depreciation_for_mgmt,
                                    "B10":bad_debt,
                                    "B11":other_fixed_mgmt,
                                    "B12":mgmt_total,
                                    "B13":mgmt_total,
                                    "B14":depreciation_for_mgmt,
                                    "B15":bad_debt,
                                    "B16":cash_out_flows_due_to_mgmt
                                    })
FinalColumn2("F2",'sssbssssssbssss',AE)
wb9.save(os.path.join(excel_save_folder_path, "管理费用预算.xlsx"))






wb10 = Workbook()
liquid= wb10.active
liquid.title = "管理费用预算"
# 第1列表头（季度列）
header_cols = ["", "2025年第1季度", "2025年第2季度", "2025年第3季度", "2025年第4季度", "2025年余额或合计"]
for col, val in enumerate(header_cols, start=1):
    liquid.cell(row=1, column=col, value=val)

# A列行标题（完整行标签，和原图一一对应）
row_texts = [
    "一、期初现金余额",
    "二、现金流入",
    "1.销售活动现金流入",
    "三、可使用现金总量",
    "四、现金支出",
    "1.材料采购现金支出",
    "2.直接工资现金支出",
    "3.制造费用现金支出",
    "4.销售费用现金支出",
    "5.管理费用现金支出",
    "6.缴纳所得税现金支出",
    "7.股利分配现金支出",
    "购置生产用固定资产现金支",
    "9.股权投资现金支出",
    "10.支付长期借款利息",
    "五、现金结余或短缺",
    "六、短期融资计划",
    "1.短期借款",
    "2.归还短期借款本金",
    "3.支付短期借款利息",
    "七、期末现金余额"
]

for row, text in enumerate(row_texts, start=2):
    liquid.cell(row=row, column=1, value=text)


coPaylonginterest=(np.array(long)*APR/4).tolist()
cash_out_flows1=cash_out_flows_due_to_mat_purchase[:-2]+np.array(direct_labor_total_cost)+np.array(cash_out_flows_due_to_mfg)+np.array(cash_out_flows_due_to_selling)+cash_out_flows_due_to_mgmt
cash_out_flows2=coTax+np.array(coDividend)+coCapitalEx+np.array(coInvest)+coPaylonginterest
total_cash_out_flows=(cash_out_flows1+cash_out_flows2).tolist()

usable_cash=[0,0,0,0]
cash_balance=[0,0,0,0]
short_term_loan=[0,0,0,0]
return_principal=[0,0,0,0]
short_term_interest=[0,0,0,0]
final_cash_balance=[0,0,0,0]
short_fin_plan=[0,0,0,0]
initial_cash_balance=[0,0,0,0]

for i in range(0,4):
    if i == 0:
        initial_cash_balance[i]=begin_cash
      
    else:
        initial_cash_balance[i]=float(final_cash_balance[i-1])
        
    usable_cash[i]=float(initial_cash_balance[i])+float(cash_inflow[i])
    cash_balance[i]=float(usable_cash[i])-float(total_cash_out_flows[i])

    if cash_balance[i] < kill_line:
        short_term_loan[i]=10000*np.ceil((kill_line-cash_balance[i])/10000)
    else:
        short_term_loan[i]=0

    if cash_balance[i] > excess_line:
        return_principal[i]=max(cash_balance[i]-kill_line,short_term_loan[i])
    else:
        return_principal[i]=0
    short_term_loan[i] = short_term_loan[i]-return_principal[i]
    short_term_interest[i]=-Sinterest_rate*sum(short_term_loan)*3
    short_fin_plan[i]=np.array(short_term_loan[i])-np.array(return_principal[i])+np.array(short_term_interest[i])
    final_cash_balance[i]=cash_balance[i]+short_fin_plan[i]

write_row_start_at2(liquid,{"B2":initial_cash_balance,
                                    "B3":cash_inflow[:-2],
                                    "B4":cash_inflow[:-2],
                                    "B5":usable_cash,
                                    "B6":total_cash_out_flows,
                                    "B7":cash_out_flows_due_to_mat_purchase[:-2],
                                    "B8":direct_labor_total_cost,
                                    "B9":cash_out_flows_due_to_mfg,
                                    'B10':cash_out_flows_due_to_selling,
                                    "B11":cash_out_flows_due_to_mgmt,
                                    "B12":coTax,
                                    "B13":coDividend,
                                    "B14":coCapitalEx,
                                    "B15":coInvest,
                                    "B16":coPaylonginterest,
                                    "B17":cash_balance,
                                    "B18":short_fin_plan,
                                    "B19":short_term_loan,
                                    "B20":return_principal,
                                    "B21":short_term_interest,
                                    "B22":final_cash_balance}
                                    )
FinalColumn2("F2",'bssssssssssssssesssse',liquid)
liquid["F22"]="=E22"
wb10.save(os.path.join(excel_save_folder_path, "现金预算.xlsx"))



wb11 = Workbook()
incomeStatement= wb11.active
incomeStatement.title = "管理费用预算"

A=sum(Sales_revenue[:-2])
B=begin_inventory_value+sum(Sales_amount[:-2])*unit_inventory_cost-Initial_inventory[0]*unit_inventory_cost
B2=sum(Sales_amount)*unit_inventory_cost
C=sum(var_selling_expense)
D=sum(var_mgmt_cost)
E=sum(fix_mfg_total)
F=sum(fix_selling_total)
G=sum(fix_mgmt_cost)
H=sum(coPaylonginterest)-sum(short_term_interest)
P=A-B-C-D-E-F-G-H
income_tax=P*tax_rate if P>0 else 0
table_data = [
    (1, "", "2025年"),
    (2, "营业收入",A ),
    (3, "营业成本",B),
    (4, "变动销售费用", C),
    (5, "变动管理费用", D),
    (6, "边际贡献总额",A-B-C-D ),
    (7, "固定制造费用",E ),
    (8, "固定销售费用",F ),
    (9, "固定管理费用", G),
    (10, "财务费用", H),
    (11, "利润总额",P ),
    (12, "所得税", income_tax),
    (13, "净利润", P-income_tax)
]
for row_num, a_content, b_content in table_data:
    incomeStatement.cell(row=row_num, column=1, value=a_content)
    incomeStatement.cell(row=row_num, column=2, value=b_content)

wb11.save(os.path.join(excel_save_folder_path, "预算利润表.xlsx"))
# 创建工作簿与工作表
wb12 = Workbook()
balanceSheet = wb12.active
balanceSheet.title = "2025年资产负债表"

fixed_assets1=fixed_assets0+new-sum(depreciation_for_mfg_asset)-sum(depreciation_for_mgmt)-sum(depreciation_for_selling)
# 负债及所有者权益总计
total_liabilities_equity0 = 875350
# 硬编码表格全量数据（行号:A列项目,B列期初余额,C列期末余额）

L1=Payable_from_last_period[4]
L2=taxes_payable0-sum(coTax)+P*tax_rate
L3=sum(short_term_loan)
total_equity1=common_stock0+P-P*tax_rate-sum(coDividend)+retained_earnings0
current_assets1=final_cash_balance[3]+Sales_revenue[3]*credit_sales_ratio+num*unit_cost+NUM*UNIT_COST

table_data = [
    (1, "项目", "期初余额", "期末余额"),
    (2, "资产", "", ""),
    (3, "一、流动资产", current_assets0, current_assets1),
    (4, "1.现金及等价物", cash_cash_equivalents0, final_cash_balance[3]),
    (5, "2.应收账款", accounts_receivable0, Sales_revenue[3]*credit_sales_ratio),
    (6, "3.原材料存货", raw_material_inventory0, num*unit_cost),
    (7, "4.产成品存货", finished_goods_inventory0, NUM*UNIT_COST),
    (8, "二、非流动资产", non_current_assets0, land_use_rights0+fixed_assets1+long_term_equity_investment0+sum(coInvest)),
    (9, "1.土地使用权", land_use_rights0, land_use_rights0),
    (10, "2.固定资产", fixed_assets0,fixed_assets1 ),
    (11, "（1）生产用固定资产", production_fixed_assets0, production_fixed_assets0+new-sum(depreciation_for_mfg_asset)),
    (12, "（2）销售&管理用固定资产", sales_admin_fixed_assets0, sales_admin_fixed_assets0-sum(depreciation_for_mgmt)-sum(depreciation_for_selling)),
    (13, "3.长期股权投资", long_term_equity_investment0, long_term_equity_investment0+sum(coInvest)),
    (14, "三、资产合计", total_assets0,current_assets1+land_use_rights0+fixed_assets1+long_term_equity_investment0+sum(coInvest) ),
    (15, "", "", 0),
    (16, "负债与所有者权益", "", ""),
    (17, "一、流动负债", current_liabilities0, L1+L2+L3),
    (18, "1.应付账款", accounts_payable0, Payable_from_last_period[4]),
    (19, "2.应交税费", taxes_payable0, taxes_payable0-sum(coTax)+income_tax),
    (20, "3.短期借款", short_term_loan0, sum(short_term_loan)),
    (21, "二、非流动负债", non_current_liabilities0, long_term_loan0),
    (22, "1.长期借款", long_term_loan0, long_term_loan0),
    (23, "三、所有者权益", total_equity0, total_equity1),
    (24, "1.普通股", common_stock0, common_stock0),
    (25, "2.留存收益", retained_earnings0, retained_earnings0+P-income_tax-sum(coDividend)),
    (26, "四、负债与所有者权益合计", total_liabilities_equity0, L1+L2+L3+long_term_loan0+total_equity1)
]
# 循环写入单元格
for row_num, a_content, b_content, c_content in table_data:
    balanceSheet.cell(row=row_num, column=1, value=a_content)
    balanceSheet.cell(row=row_num, column=2, value=b_content)
    balanceSheet.cell(row=row_num, column=3, value=c_content)
# 自动调整列宽优化可读性
balanceSheet.column_dimensions['A'].width = 28
balanceSheet.column_dimensions['B'].width = 15
balanceSheet.column_dimensions['C'].width = 15
wb12.save(os.path.join(excel_save_folder_path, "预算资产负债表.xlsx"))

from pyecharts.charts import Pie, Timeline
from pyecharts import options as opts
from pyecharts.globals import ThemeType

# 期初资产明细（直接取你原变量，无嵌套列表下标）
data_qc = [
    ["现金及等价物", cash_cash_equivalents0],
    ["应收账款", accounts_receivable0],
    ["原材料存货", raw_material_inventory0],
    ["产成品存货", finished_goods_inventory0],
    ["土地使用权", land_use_rights0],
    ["生产用固定资产", production_fixed_assets0],
    ["销管固定资产", sales_admin_fixed_assets0],
    ["长期股权投资", long_term_equity_investment0]
]

# 期末资产明细（直接套用你原有计算公式，无硬编码数值）
data_qm = [
    ["现金及等价物", final_cash_balance[3]],
    ["应收账款", Sales_revenue[3]*credit_sales_ratio],
    ["原材料存货", num*unit_cost],
    ["产成品存货", NUM*UNIT_COST],
    ["土地使用权", land_use_rights0],
    ["生产用固定资产", production_fixed_assets0+new-sum(depreciation_for_mfg_asset)],
    ["销管固定资产", sales_admin_fixed_assets0-sum(depreciation_for_mgmt)-sum(depreciation_for_selling)],
    ["长期股权投资", long_term_equity_investment0+sum(coInvest)]
]

# 单函数生成饼图，仅微调标签，完整显示百分比
def gen_pie(data, title):
    return (
        Pie(init_opts=opts.InitOpts(theme=ThemeType.MACARONS))
        .add("资产结构", data, rosetype="area", radius=["30%", "75%"])
        .set_global_opts(
            title_opts=opts.TitleOpts(title=title),
            legend_opts=opts.LegendOpts(orient="vertical", pos_left="2%")
        )
        # 格式：名称 + 金额 + 百分比，调整标签位置避免重叠
        .set_series_opts(
            label_opts=opts.LabelOpts(
                formatter="{b}\n金额：{c}\n占比：{d}%",
                position="outside"
            )
        )
    )

# 时间轮播
tl = Timeline()
tl.add(gen_pie(data_qc, "期初资产结构"), "期初")
tl.add(gen_pie(data_qm, "期末资产结构"), "期末")
tl.render("test.html")
import webbrowser
webbrowser.open("test.html")