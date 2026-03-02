# 缩略词提取工具

从Word文档中自动提取格式为"中文（英文全称, 英文缩写）"的缩略词，并按照指定模板格式输出为Word文档。

## 功能特点

- 支持多种缩略词格式识别
  - 标准格式：`中文（English Full Name, EFN）`
  - 混合格式：`旋转IoU（Rotated IoU, RIoU）`
  - 纯中文格式：`卷积神经网络（Convolutional Neural Networks, CNN）`
- 自动去重，避免重复提取
- 同时检查正文段落和表格内容
- 按英文缩写字母顺序排序
- 输出为规范的Word文档格式

## 安装依赖

```bash
pip install -r requirements.txt
```

或手动安装：

```bash
pip install python-docx lxml
```

## 使用方法

### 基本用法

```bash
python extract_abbreviations.py <论文文档路径>
```

示例：

```bash
python extract_abbreviations.py 面向SAR图像舰船目标的多尺度轻量化检测研究.docx
```

### 指定模板

```bash
python extract_abbreviations.py <论文文档路径> <模板路径>
```

示例：

```bash
python extract_abbreviations.py 论文.docx 缩略词格式.docx
```

## 输出格式

脚本会生成 `缩略词列表.docx` 文件，包含一个三列表格：

| 英文缩写 | 英文全称 | 中文全称 |
|---------|---------|---------|
| CNN | Convolutional Neural Networks | 卷积神经网络 |
| SAR | Synthetic Aperture Radar | 合成孔径雷达 |

## 模板要求

模板文档（`缩略词格式.docx`）应包含：
- 标题段落（可选）
- 一个三列表格，表头为：英文缩写、英文全称、中文全称

## 示例输出

```
正在提取 论文.docx 中的缩略词...

共找到 53 个缩略词：

--------------------------------------------------------------------------------
序号    英文缩写           英文全称                                    中文全称
--------------------------------------------------------------------------------
1     CNN            Convolutional Neural Networks           卷积神经网络
2     SAR            Synthetic Aperture Radar                合成孔径雷达
...

结果已保存到：缩略词列表.docx
```

## 技术说明

- 使用正则表达式匹配多种缩略词格式
- 支持中英文括号（）和（）
- 支持中英文逗号，和,
- 自动清理提取结果中的多余标点和连接词
- 基于英文缩写和英文全称进行去重

## 系统要求

- Python 3.6+
- python-docx 1.2.0+
- lxml 3.1.0+

## 作者

Kequan

## 许可证

MIT License
