# 缩略词提取工具

从 Word（`.docx`）正文和表格中提取“中文名称、英文全称、英文缩写”，去重后按缩写排序，并使用 Word 模板生成缩略词列表。适用于论文中的显式缩略词定义整理。

## 快速开始

需要 **Python 3.9+**。在项目目录中创建虚拟环境并安装依赖：

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Windows PowerShell 中使用 `python -m venv .venv` 创建环境，使用 `.venv\Scripts\Activate.ps1` 激活。

准备论文和模板后运行：

```bash
python extract_abbreviations.py "论文.docx" "缩略词格式.docx"
```

若当前目录已有 `缩略词格式.docx`，可省略模板参数：

```bash
python extract_abbreviations.py "论文.docx"
```

路径可以是相对路径或绝对路径；含空格时请加引号。仅支持 `.docx`，旧版 `.doc` 文件需先转换格式。

## 模板与输出

模板的**第一个表格**应为三列表格，仅保留以下表头；可在表格前添加标题段落：

| 英文缩写 | 英文全称 | 中文全称 |
| --- | --- | --- |

脚本在第一个表格末尾追加结果，不清除模板已有数据，因此不要将上次导出的列表用作模板。

默认输出到**当前工作目录**的 `缩略词列表.docx`，示例内容如下：

| 英文缩写 | 英文全称 | 中文全称 |
| --- | --- | --- |
| CNN | Convolutional Neural Networks | 卷积神经网络 |
| SAR | Synthetic Aperture Radar | 合成孔径雷达 |

- 再次运行会覆盖同名输出文件。
- 模板缺失、格式不符或 Word 保存失败时，会尝试写入 `缩略词列表.txt`。
- 未提取到结果时不生成新文件，也不会清除先前的输出。

## 支持的提取格式

| 类型 | 示例 |
| --- | --- |
| 标准定义 | `卷积神经网络（Convolutional Neural Networks, CNN）` |
| 混合中英文名称 | `旋转IoU（Rotated IoU, RIoU）` |
| 缩写在前 | `目标检测器（DETR, DEtection TRansformer）` |
| 小写开头的缩写 | `平均精度均值（mean Average Precision, mAP）` |
| 数字及连字符 | `三维网络（3D Convolutional Neural Network, 3D-CNN）` |
| Unicode 横线 | `精确率—召回率（Precision–Recall, PR）` |

支持中英文括号、逗号或分号，以及连续空白和不间断空格。按文档顺序读取正文和表格（含嵌套表格），同一段落内跨多个文本片段的定义也可识别。

### 去重规则

- 以“英文缩写 + 英文全称”为依据，忽略英文全称大小写、连续空白和横线样式差异。
- 保留缩写大小写区别；同一缩写对应不同英文全称时保留多条。
- 重复定义保留首次出现的中文名称，不自动纠正原文拼写。

### 识别边界

- 括号前须含中文，括号内须为全称和缩写两个字段；不根据孤立缩写推测全称。
- 缩写须至少含一个大写字母，默认至少两个字母，排除 Q/K/V 等单字母符号。
- 不支持嵌套括号定义或英文全称内含逗号的格式，也不读取图片、公式、页眉和页脚。
- 常见参考文献格式会被过滤，但不进行参考文献章节识别。
- 中文前缀清理采用启发式规则，复杂并列句和修饰语仍需人工复核。例如“高分辨率合成孔径雷达”会保留修饰语。

## Python 接口

```python
from extract_abbreviations import extract_abbreviations, save_to_word

# 返回列表，每条记录为 (英文缩写, 英文全称, 中文名称)。
rows = extract_abbreviations("论文.docx")
save_to_word(rows, "缩略词格式.docx", "自定义输出.docx")

# 如需提取 Q/K/V 等单字母符号：
rows_with_symbols = extract_abbreviations("论文.docx", include_single_letter=True)
```

命令行使用固定输出文件名；自定义输出路径或开启单字母识别时，使用 Python 接口。

## 开发与验证

```bash
python -m unittest -v
```

测试覆盖格式变体、误匹配过滤、中文前缀、嵌套表格、文档顺序、去重和 Word 导出。测试会创建临时文档，无需本地论文或模板。

| 文件 | 用途 |
| --- | --- |
| `extract_abbreviations.py` | 提取规则、Word 导出和命令行入口 |
| `test_extract_abbreviations.py` | 回归测试 |
| `requirements.txt` | 运行依赖 |
| [DESIGN.md](DESIGN.md) | 规则设计、限制和变更记录 |
| `.gitignore` | 本地文件及生成文件的忽略规则 |

`.gitignore` 默认忽略本地 Word 文档、生成结果、虚拟环境及缓存，但保留根目录 `缩略词格式.docx` 的版本管理入口。其他需要共享的 Word 样例应显式添加例外规则；本地论文无需提交。

## 作者与许可证

作者：Kequan。许可证：MIT。
