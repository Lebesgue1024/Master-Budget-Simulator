# Master‑Budget‑Simulator
管理会计全面预算编制仿真程序

## 📖项目简介
本项目是管理会计课程大作业，基于 Python + Tkinter 实现图形化参数录入，借助 openpyxl 自动批量生成全套管理会计预算Excel报表，同时使用 Pyecharts 生成期初、期末资产结构可视化饼图。
内置兴腾公司案例模板数据，支持一键填充案例参数，用户可直接在GUI界面修改各项预算参数，校验完成后自动导出全部预算表格。

## 📋生成报表清单
运行脚本后本地自动生成如下Excel文件：
- 销售预算.xlsx
- 生产预算.xlsx
- 原材料预算.xlsx
- 直接人工预算.xlsx
- 制造费用预算.xlsx
- 单位成本预算.xlsx
- 期末存货预算.xlsx
- 销售费用预算.xlsx
- 管理费用预算.xlsx
- 现金预算.xlsx
- 预算利润表.xlsx
- 预算资产负债表.xlsx

同时输出 `test.html`，包含期初、期末资产结构轮播玫瑰饼图。

## 🛠环境依赖
运行脚本前请安装所需第三方库
```bash
pip install numpy openpyxl pyecharts
