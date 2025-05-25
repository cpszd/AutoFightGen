import os
import re
import pandas as pd
import json
import yaml

from utils import read_excel_with_colors


COLUMNS = ["1", "2", "3", "4", "5"]


def read_json(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        return json.load(f)


# 处理并输出JSON
def process_and_output(df, config):
    template_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "configs", "fight_action.json"
    )
    action_templates = read_json(template_path)
    actionmap_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "configs", "action_map.json"
    )
    action_map = read_json(actionmap_path)
    operationmap_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "configs", "operation_map.json"
    )
    operation_map = read_json(operationmap_path)

    def get_action(action_code):
        position = action_code[0]  # 获取位置编号，例如 "1"
        action_type = action_code[1]  # 获取动作类型，例如 "普"
        key = f"{position}号位" + (
            "普攻"
            if action_type == "普"
            else (
                "上拉"
                if action_type == "大"
                else "下拉" if action_type == "下" else "O"
            )
        )
        return action_templates.get(key)

    def get_slide_operation(previous_color, target_color):

        color_list = config["color_list"]

        # 查找颜色在列表中的索引
        previous_index = color_list.index(previous_color)
        target_index = color_list.index(target_color)

        # 计算顺时针和逆时针的距离
        clockwise_distance = (target_index - previous_index) % len(color_list)
        counterclockwise_distance = (previous_index - target_index) % len(color_list)

        # 如果顺时针距离小于逆时针距离，选择右滑；否则选择左滑
        if clockwise_distance <= counterclockwise_distance:
            return ["右侧目标"] * clockwise_distance  # 右滑操作，重复输出距离次
        else:
            return ["左侧目标"] * counterclockwise_distance  # 左滑操作，重复输出距离次

    def set_operation_action(action_op, idx, i, result_config, current_action_key):
        if action_op != "未知":
            i += 1
            action_key = f"回合{idx}行动{i}"
            if action_op == "左侧目标":
                result_config[action_key] = {
                    "text_doc": "左侧目标",
                    "action": "Click",
                    "target": [154, 648, 1, 1],
                    "post_delay": 2000,
                    "duration": 800,
                }
            elif action_op == "右侧目标":
                result_config[action_key] = {
                    "text_doc": "右侧目标",
                    "action": "Click",
                    "target": [603, 413, 18, 21],
                    "post_delay": 2000,
                    "duration": 800,
                }
            elif action_op == "检测橙星":
                result_config[action_key] = {
                    "text_doc": "检测橙星",
                    "recognition": "ColorMatch",
                    "roi": [77, 167, 70, 70],
                    "method": 4,
                    "upper": [255, 255, 205],
                    "lower": [166, 140, 85],
                    "count": 1,
                    "order_by": "Score",
                    "connected": True,
                    "action": "Click",
                    "pre_delay": 2000,
                }
                if current_action_key:
                    result_config[current_action_key]["on_error"] = [
                        "抄作业点左上角重开"
                    ]
                    result_config[current_action_key]["timeout"] = 200
                else:
                    result_config[f"检测回合{idx}"]["on_error"] = ["抄作业点左上角重开"]
                    result_config[f"检测回合{idx}"]["timeout"] = 200

            if current_action_key:
                result_config[current_action_key]["next"] = [action_key]
            current_action_key = action_key
        else:
            print(f"未知的操作符: {direction}")

        return i, result_config, current_action_key

    # 解析行数据并生成行动顺序
    def parse_actions_for_row(row):
        action_order = {}

        for i, seq in enumerate(row):
            if isinstance(seq, str):
                seq = (
                    seq.replace("普攻", "普")
                    .replace("技能", "大")
                    .replace("防御", "防")
                )
                actions = re.findall(r"([\D]*)(\d)([\D])", seq)
                for action in actions:
                    operations = action[0]  # 操作部分
                    # 填补没有标颜色的操作符
                    if config["use_color"] == True:
                        if not operations or operations[0] not in config["color_list"]:
                            operations = actions[0][0] + operations
                    number = int(action[1])  # 数字部分
                    symbol = action[2]  # 文本部分
                    action_type = action_map.get(
                        symbol, "未知"
                    )  # 使用 '未知' 来处理未定义的符号

                    if action_type != "未知":
                        # 保存动作，COLUMNS[i]是列标
                        action_order[number] = {
                            "action": f"{operations}{COLUMNS[i]}{action_type}"  # 例如 '1大', '5普'
                        }
                    else:
                        print(f"未知的动作符号: {symbol}")
        return action_order

    result_config = {}
    max_round_num = len(df)

    pre_color = ""  # 用于存储上一个颜色
    for idx, row in enumerate(df.values, start=1):
        # 生成回合信息
        result_config[f"检测回合{idx}"] = {
            "recognition": "OCR",
            "expected": f"回合{idx}",
            "roi": [585, 28, 90, 65],
            "next": [
                f"回合{idx}行动1",
            ],
            "post_delay": config["round_post_delay"],
        }

        current_action_key = None

        # 解析当前行中的行动
        action_order = parse_actions_for_row(row)

        # 按照数字排序
        sorted_actions = sorted(action_order.items())

        total_actions = sum(len(action["action"]) - 1 for _, action in sorted_actions)

        i = 0
        # 生成并保存每个排序后的行动
        for _, action in sorted_actions:
            match = re.search(r"([\D]+)(?=\d)", action["action"])
            if match:
                directions = match.group(0)
                for direction in directions:
                    action_op = operation_map.get(direction, "未知")
                    if config["use_color"] == True:
                        if direction in config["color_list"]:
                            if pre_color == direction or pre_color == "":
                                pre_color = direction
                                continue
                            else:
                                slide_actions = get_slide_operation(
                                    pre_color, direction
                                )
                                pre_color = direction
                                for op in slide_actions:
                                    i, result_config, current_action_key = (
                                        set_operation_action(
                                            op,
                                            idx,
                                            i,
                                            result_config,
                                            current_action_key,
                                        )
                                    )
                        else:
                            i, result_config, current_action_key = set_operation_action(
                                action_op, idx, i, result_config, current_action_key
                            )
                    else:
                        i, result_config, current_action_key = set_operation_action(
                            action_op, idx, i, result_config, current_action_key
                        )

            i += 1

            action_config = get_action(action["action"][-2:])
            if action_config:
                action_key = f"回合{idx}行动{i}"
                result_config[action_key] = action_config.copy()
                result_config[action_key]["text_doc"] = action["action"][-2:]

                if current_action_key:
                    result_config[current_action_key]["next"] = [action_key]

                current_action_key = action_key

                # 如果是当前回合的最后一个动作，设置next指向下一回合
                if i == total_actions and int(idx) < max_round_num:
                    result_config[action_key]["next"] = [
                        "抄作业战斗胜利",
                        f"检测回合{int(idx)+1}",
                    ]

    # 添加重启信息
    add_restart_info(result_config, config)

    # 输出JSON并保存
    json_output = json.dumps(result_config, ensure_ascii=False, indent=4)
    outputpath = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "outputs",
        f"{config['json_name']}.json",
    )
    with open(outputpath, "w", encoding="utf-8") as f:
        f.write(json_output)
    print(f"JSON 文件已保存：{outputpath}")


# 添加重启信息
def add_restart_info(result_config, config):

    # 根据关卡类别设置重开后的导航节点
    if config["level_type"] == "主线":
        next_node = "抄作业找到关卡-主线"
    elif config["level_type"] == "洞窟":
        next_node = "抄作业进入关卡-洞窟"
    elif config["level_type"] == "活动有分级":
        next_node = "抄作业找到关卡-活动分级"
    elif config["level_type"] == "白鹄":
        next_node = "抄作业进入关卡-白鹄"
    else:
        next_node = "抄作业找到关卡-OCR"

    result_config["抄作业点左上角重开"] = {
        "recognition": "TemplateMatch",
        "template": "back.png",
        "green_mask": True,
        "threshold": 0.5,
        "roi": [6, 8, 123, 112],
        "action": "Click",
        "pre_delay": 2000,
        "post_delay": 2000,
        "next": ["抄作业确定左上角重开", next_node],
        "timeout": 20000,
    }

    # 根据关卡类别和识别名称设置对应的导航节点
    if config["level_type"] == "洞窟":
        if config["cave_type"] == "左":
            result_config["抄作业进入关卡-洞窟"] = {
                "text_doc": "左",
                "recognition": "OCR",
                "expected": "前往",
                "roi": [237, 810, 82, 89],
                "action": "Click",
                "target": [258, 833, 42, 39],
                "pre_delay": 1500,
                "next": ["抄作业战斗开始"],
                "timeout": 20000,
            }
        else:
            result_config["抄作业进入关卡-洞窟"] = {
                "text_doc": "右",
                "recognition": "OCR",
                "expected": "前往",
                "roi": [558, 804, 79, 89],
                "action": "Click",
                "target": [581, 832, 41, 41],
                "pre_delay": 1500,
                "next": ["抄作业战斗开始"],
                "timeout": 20000,
            }
    elif config["level_type"] == "活动有分级":
        result_config["抄作业找到关卡-活动分级"] = {
            "recognition": "OCR",
            "expected": config["level_recognition_name"],  # 使用传入的识别名称
            "roi": [0, 249, 720, 1030],
            "action": "Click",
            "pre_delay": 1500,
            "next": ["抄作业选择活动分级"],
            "timeout": 20000,
        }
        result_config["抄作业选择活动分级"] = {
            "recognition": "OCR",
            "expected": config["difficulty"],  # 使用传入的难度等级
            "roi": [37, 351, 647, 491],  # todo 需要根据活动调整
            "pre_delay": 1500,
            "action": "Click",
            "next": ["抄作业进入关卡"],
            "timeout": 20000,
        }
    elif (
        config["level_type"] != "主线" and config["level_type"] != "白鹄"
    ):  # 其他的情况
        result_config["抄作业找到关卡-OCR"] = {
            "recognition": "OCR",
            "expected": config["level_recognition_name"],  # 使用传入的识别名称
            "roi": [0, 249, 720, 1030],
            "action": "Click",
            "pre_delay": 2000,
            "next": ["抄作业战斗开始"],
            "timeout": 20000,
        }


# 主程序入口
if __name__ == "__main__":
    config_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "configs", "config.yaml"
    )
    with open(config_path, "r", encoding="utf-8") as file:
        config = yaml.safe_load(file)

    # 读取 outputs 文件夹下的第一个 .xlsx 文件
    outputs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "outputs")
    xlsx_files = [f for f in os.listdir(outputs_dir) if f.endswith(".xlsx")]

    if not xlsx_files:
        raise FileNotFoundError("outputs 文件夹中未找到任何excel文件")

    xlsx_path = os.path.join(outputs_dir, xlsx_files[0])
    # 读取excel
    if config["use_color"] == True and config["color_type"] == "fill":
        df = read_excel_with_colors(xlsx_path, config["use_header"])
    else:
        df = pd.read_excel(xlsx_path, header=0 if config["use_header"] else None)
    process_and_output(df, config)
