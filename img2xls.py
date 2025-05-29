import os
import re
import base64
import yaml
import pandas as pd
from io import StringIO
from zhipuai import ZhipuAI
from utils import resource_path


def image_to_text(img_path,api_key=None, model="glm-4v-flash"):
    with open(img_path, 'rb') as f:
        img_base64 = base64.b64encode(f.read()).decode()

    client = ZhipuAI(api_key=api_key)
    response = client.chat.completions.create(
        model=model,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": img_base64}},
                {"type": "text", "text": "帮我把这个图片转为表格,输出html格式"}
            ]
        }]
    )
    return response.choices[0].message.content


def text_to_excel(content, output_path):
    match = re.search(r"<table.*?>.*?</table>", content, re.DOTALL)
    if not match:
        print("未能提取 HTML 表格。")
        return

    df = pd.read_html(StringIO(match.group(0)))[0]
    df = df.applymap(lambda x: str(x).replace(' ', '') if isinstance(x, str) else x)
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [col[-1] for col in df.columns]

    df.to_excel(output_path, index=False)
    print(f"Excel 保存成功: {output_path}")


def main():
    with open(resource_path("configs/config.yaml"), encoding="utf-8") as f:
        config = yaml.safe_load(f)

    img_dir = resource_path("images")
    img_files = [f for f in os.listdir(img_dir) if f.lower().endswith((".jpg", ".png", ".jpeg"))]

    if not img_files:
        raise FileNotFoundError("images 文件夹中未找到任何图片文件")

    img_path = os.path.join(img_dir, img_files[0])
    output_path = resource_path(os.path.join("outputs", f"{config['excel_name']}.xlsx"))

    content = image_to_text(img_path,config['api_key'], config['model'])
    text_to_excel(content, output_path)


if __name__ == "__main__":
    main()
