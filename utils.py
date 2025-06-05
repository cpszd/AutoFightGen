import colorsys
import os
import sys

import openpyxl
import pandas as pd

def resource_path(relative_path):
    if getattr(sys, 'frozen', False):  # 如果是 PyInstaller 打包环境
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def rgb_to_named_color(color):
    """将 RGB 颜色转换为红、橙、黄、绿、青、蓝、紫等主要颜色"""
    r, g, b = color
    h, s, v = colorsys.rgb_to_hsv(r / 255.0, g / 255.0, b / 255.0)
    # 计算色相（Hue）角度
    h = h * 180
    s = s * 255
    v = v * 255

    # 根据 HSV 值确定主颜色
    if 0 <= h <= 180 and s <= 43 and v <= 46:
        return "黑"
    elif 0 <= h <= 180 and s <= 30 and v >= 221:
        return "白"
    elif 0 <= h <= 180 and s <= 43 and 46 < v < 221:
        return "灰"
    elif (0 <= h <= 10 or 156 <= h <= 180) and s >= 43 and v >= 46:
        return "红"
    elif 11 <= h <= 23 and s >= 43 and v >= 46:
        return "橙"
    elif 24 <= h <= 34 and s >= 43 and v >= 46:
        return "黄"
    elif 35 <= h <= 82 and s >= 43 and v >= 46:
        return "绿"
    elif 83 <= h <= 124 and s >= 43 and v >= 46:
        return "蓝"
    elif 125 <= h <= 155 and s >= 43 and v >= 46:
        return "紫"
    else:
        return f"{h:.2f} {s:.2f} {v:.2f}"


def argb_to_rgb(argb):
    argb = int(argb, 16)
    # 提取 RGB 部分
    r = (argb >> 16) & 0xFF
    g = (argb >> 8) & 0xFF
    b = argb & 0xFF
    return r, g, b


def theme_to_named_color(theme):
    """将主题颜色转换为命名颜色"""
    theme_colors = {
        0: "白",
        1: "黑",
        2: "白",
        3: "蓝",
        4: "蓝",
        5: "红",
        6: "绿",
        7: "紫",
        8: "蓝",
        9: "橙",
    }
    return theme_colors.get(theme, "未知")


def read_excel_with_colors(file_path, use_header=False):
    """读取 Excel 表格中的文字和背景颜色"""
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    data = []
    for row in ws.iter_rows():
        if use_header:
            if row[0].row == 1:
                continue
        row_data = []
        for cell in row:
            text = cell.value  # 获取单元格文字
            if text is None:
                row_data.append("")
                continue
            fill = cell.fill  # 获取单元格填充样式
            if hasattr(fill, "start_color") and fill.start_color.index != "00000000":
                # 如果使用主题颜色
                if type(fill.start_color.theme) == int:
                    theme = fill.start_color.theme
                    color = theme_to_named_color(theme)  # 获取主题颜色
                elif fill.start_color.type == "rgb":
                    color = fill.start_color.rgb  # 获取 RGB 颜色
                    color = rgb_to_named_color(argb_to_rgb(color))  # 转换为命名颜色
                else:
                    color = "白"
            else:
                color = "白"  # 无背景颜色
            row_data.append(f"{color}{text}")  # 合并颜色和文字
        if any(cell != "" for cell in row_data):
            data.append(row_data)
    df = pd.DataFrame(data)
    return df
