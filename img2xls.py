# https://github.com/RapidAI/TableStructureRec
# 安装：pip install wired_table_rec lineless_table_rec table_cls
import os
import pandas as pd
from lineless_table_rec import LinelessTableRecognition
from lineless_table_rec.utils_table_recover import format_html
from table_cls import TableCls
from wired_table_rec import WiredTableRecognition
import yaml


def image_to_excel(img_path, output_path):
    lineless_engine = LinelessTableRecognition()
    wired_engine = WiredTableRecognition()
    # 默认小yolo模型(0.1s)，可切换为精度更高yolox(0.25s),更快的qanything(0.07s)模型
    table_cls = TableCls()  # TableCls(model_type="yolox"),TableCls(model_type="q")
    cls, elasp = table_cls(img_path)
    if cls == "wired":
        table_engine = wired_engine
    else:
        table_engine = lineless_engine
    html, elasp, polygons, logic_points, ocr_res = table_engine(img_path)

    complete_html = format_html(html)
    dfs = pd.read_html(complete_html)
    if dfs:
        df = dfs[0]
        df.to_excel(output_path, index=False, engine="openpyxl")
        print(f".xlsx 文件已保存到: {output_path}")
    else:
        print("未找到 HTML 表格，无法保存为 .xlsx 文件")


def main():
    config_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "configs", "config.yaml"
    )
    with open(config_path, "r", encoding="utf-8") as file:
        config = yaml.safe_load(file)

    img_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "images")
    img_files = [
        f for f in os.listdir(img_dir) if f.lower().endswith((".jpg", ".png", ".jpeg"))
    ]

    if not img_files:
        raise FileNotFoundError("outputs 文件夹中未找到任何 image 文件")

    img_path = os.path.join(img_dir, img_files[0])
    output_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "outputs",
        f"{config['excel_name']}.xlsx",
    )
    image_to_excel(img_path, output_path)


if __name__ == "__main__":
    main()
