# Excel 参数依赖分析与计算可视化应用

这是一个基于Flask的Web应用，用于分析Excel文件中的参数依赖关系并进行可视化展示与计算。

## 功能特点

- 分析Excel文件中的参数及其依赖关系
- 可视化展示参数之间的依赖链
- 支持修改输入参数并实时计算结果
- 美化展示参数详细信息和计算公式
- 智能识别参数类型，分类显示输入参数、中间参数和输出参数
- 通过直接调用Excel进行计算，确保计算结果的准确性

## 安装与运行

### 环境要求
- Python 3.8+
- Flask
- openpyxl
- pandas
- xlwings (用于调用Excel进行计算)
- Microsoft Excel (需已安装)

### 安装依赖
```
pip install -r requirements.txt
```

### 运行应用
```
python run.py
```

运行后，在浏览器中访问：http://127.0.0.1:5000

## 使用方法

1. 在首页上传Excel文件（.xlsx或.xls格式）
2. 系统会分析Excel中的参数关系，并跳转到可视化页面
3. 在可视化页面中：
   - 左侧显示参数分类列表
   - 中间显示参数依赖关系图
   - 点击参数节点查看详细信息和计算公式
   - 修改输入参数值，系统会使用Excel重新计算结果

## Excel文件要求

应用假设Excel文件具有特定的结构：
- 第一列为参数名称
- 第二列为单位
- 第三列为数值或公式

## 技术栈

- 后端：Flask, Python
- 前端：JavaScript, D3.js, Bootstrap
- 数据处理：pandas, openpyxl
- 计算引擎：xlwings + Microsoft Excel
- 可视化：D3.js 力导向图
