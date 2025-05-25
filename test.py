import os
from utils import read_excel_with_colors


def main():
    # 读取 outputs 文件夹下的第一个 .xlsx 文件
    outputs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "outputs")
    xlsx_files = [f for f in os.listdir(outputs_dir) if f.endswith(".xlsx")]
    xlsx_path = os.path.join(outputs_dir, xlsx_files[0])
    # 合并excel中填充的颜色（转为对应颜色文本）与文本
    df = read_excel_with_colors(xlsx_path)
    # 保存为excel
    output_path = os.path.join(outputs_dir, "test.xlsx")
    df.to_excel(output_path, index=False, header=False)  # 不保存索引
    print(f"保存成功: {output_path}")


if __name__ == "__main__":
    main()
