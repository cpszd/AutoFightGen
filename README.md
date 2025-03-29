# AutoFightGen

MaaYuan 作业生成：图片(?)表格转作业 json 文件，在C佬代码的基础上写的。

## 文件相关

**img2xls.py**：图片转excel表格

**auto_fight_gen.py**：表格转json作业

**config.yaml**：设置作业基本信息，包括关卡名，输出作业名等

**action_map.json**：设置回合动作符，如 "↓": "下" 指excel中↓对应下拉操作

**operation_map.json**：设置回合操作符，包括左侧目标，右侧目标，检测橙星

## 表格tips

- 可自定义符号，仅支持单字符

- 表格中操作需在动作前

## 图片转excel

将图片放入images文件夹下，运行**img2xls.py**，请将生成的.xlsx中的无关内容删掉。
