import os
import pandas as pd
import re

from utils import resource_path


def process_row(row):
    used_numbers = set()
    result_row = []
    # 分割块的模式：字母、箭头、数字+字母等
    split_pattern = re.compile(r"\d*[^\d\s]")
    # 已编号的单元格，比如12A
    numbered_pattern = re.compile(r"^\d+[^\d\s]$")
    
    # 先收集所有用过的编号
    for cell in row:
        if isinstance(cell, str):
            parts = split_pattern.findall(cell)
            for part in parts:
                if numbered_pattern.match(part):
                    used_numbers.add(int(re.match(r"(\d+)", part).group(1)))

    # 获取下一个未用编号
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
        parts = split_pattern.findall(cell)
        processed_parts = []

        for part in parts:
            if numbered_pattern.match(part):
                processed_parts.append(part)
            elif len(part) == 1 and (part.isalpha() or part in ["↑", "↓"]):
                num = get_next_number()
                processed_parts.append(f"{num}{part}")
            else:
                processed_parts.append(part)

        result_row.append("".join(processed_parts))  # 合并结果

    return result_row



def main():
    # 读取 outputs 文件夹下的第一个 .xlsx 文件
    outputs_dir=resource_path("outputs")
    xlsx_files = [f for f in os.listdir(outputs_dir) if f.endswith(".xlsx")]
    if not xlsx_files:
        raise FileNotFoundError("outputs 文件夹中未找到任何excel文件")
    xlsx_path = os.path.join(outputs_dir, xlsx_files[0])
    # 读取文件（换成你自己的 Excel 文件名）
    df = pd.read_excel(xlsx_path, header=None)
    # 应用到每一行
    new_df = df.apply(process_row, axis=1, result_type="expand")
    output_path=resource_path("outputs/actions.xlsx")
    
    # 保存结果
    new_df.to_excel(output_path, index=False, header=False)
    print(f"添加顺序成功,处理后的 Excel 文件已保存到: {output_path}")


if __name__ == "__main__":
    main()
