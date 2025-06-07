import os
import re
import base64
import yaml
import pandas as pd
from io import StringIO
from zhipuai import ZhipuAI
from utils import resource_path


def image_to_text(img_url,api_key=None, model="glm-4v-flash"):
    

    client = ZhipuAI(api_key=api_key)
    response = client.chat.completions.create(
        model=model,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": img_url}},
                {"type": "text", "text": "帮我把这个图片转为表格,输出Markdown格式"}
            ]
        }],
        
    )
    print(response.choices[0].message.content
)
    return response.choices[0].message.content


def markdown_table_to_excel(md_content, output_path):
    # 提取 Markdown 表格内容（以|开头的行）
    lines = [line.strip() for line in md_content.strip().split('\n') if line.strip().startswith('|')]
    if len(lines) < 3:
        print("Markdown表格格式异常或行数不足")
        return

    # 移除表头分隔行（例如|---|---|）
    lines = [line for line in lines if not re.match(r'^\|\s*[-:]+\s*(\|\s*[-:]+\s*)*\|$', line)]

    # 转成以逗号分隔的CSV格式文本，方便用pandas读取
    csv_str = '\n'.join([','.join([cell.strip() for cell in re.split(r'\s*\|\s*', line.strip('|'))]) for line in lines])
    df = pd.read_csv(StringIO(csv_str))

    # 去除字符串内空格
    df = df.applymap(lambda x: str(x).replace(' ', '') if isinstance(x, str) else x)

    df.to_excel(output_path, index=False)
    print(f"Excel 保存成功: {output_path}")


def main():
    with open(resource_path("configs/config.yaml"), encoding="utf-8") as f:
        config = yaml.safe_load(f)

    img_url=config['img_url']
    output_path = resource_path(os.path.join("outputs", f"{config['excel_name']}.xlsx"))
    content = image_to_text(img_url,config['api_key'], config['model'])
    markdown_table_to_excel(content, output_path)


if __name__ == "__main__":
    main()
