import os
import pandas as pd
import re


# 匹配“数字+字符”的正则，比如 1A、2↑、5S
pattern = re.compile(r"^(\d+)([^\d\s])$")


# 处理每一行
def process_row(row):
    used_numbers = set()
    result_row = []

    # 先收集这一行里所有已用数字
    for cell in row:
        if isinstance(cell, str):
            match = pattern.match(cell)
            if match:
                used_numbers.add(int(match.group(1)))

    # 给缺数字的字符补上行内唯一编号
    def get_next_number():
        n = 1
        while n in used_numbers:
            n += 1
        used_numbers.add(n)
        return n

    for cell in row:
        if pd.isna(cell):
            result_row.append(cell)
            continue

        cell = str(cell)
        match = pattern.match(cell)
        if match:
            result_row.append(cell)  # 已带编号，保留
        elif len(cell) == 1 and (cell.isalpha() or cell in ["↑", "↓"]):
            num = get_next_number()
            result_row.append(f"{num}{cell}")
        else:
            result_row.append(cell)  # 其他保留

    return result_row


def main():
    # 读取 outputs 文件夹下的第一个 .xlsx 文件
    outputs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "outputs")
    xlsx_files = [f for f in os.listdir(outputs_dir) if f.endswith(".xlsx")]
    if not xlsx_files:
        raise FileNotFoundError("outputs 文件夹中未找到任何excel文件")
    xlsx_path = os.path.join(outputs_dir, xlsx_files[0])
    # 读取文件（换成你自己的 Excel 文件名）
    df = pd.read_excel(xlsx_path, header=None)
    # 应用到每一行
    new_df = df.apply(process_row, axis=1, result_type="expand")
    output_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "outputs",
        "actions.xlsx",
    )
    # 保存结果
    new_df.to_excel(output_path, index=False, header=False)


if __name__ == "__main__":
    main()
