# Word 文档高亮文本提取工具

这是一个 Python 工具，用于从 Word 文档（.docx）中自动提取**黄色高亮部分**的文本，并按照规章条文整合成填空题。

## 功能

- 📄 读取 Word 文档（.docx 格式）
- 🟨 自动识别和提取黄色高亮文本
- 📋 按照"第XXX条"标记分组规章条文
- 📝 生成格式化的填空题输出文档

## 安装

### 环境要求
- Python 3.6 或更高版本

### 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 基本用法

```bash
python extract_highlights.py input.docx output.docx
```

### 参数说明

- `input.docx` - 输入的 Word 文档（必需）
- `output.docx` - 输出的 Word 文档（可选，默认为 `output.docx`）

### 示例

```bash
# 简单用法，输出到默认的 output.docx
python extract_highlights.py rules.docx

# 指定输出文件名
python extract_highlights.py rules.docx questions.docx
```

## 文件格式要求

### 输入文件格式

输入的 Word 文档应该包含以下结构：

```
第 247 条  列车应按本规程、列车编组计划和编挂条件的编组。

第 248 条  动车组为固定编组。单组动车组运用状态下不得解编...
```

其中，**黄色高亮的部分**就是你要提取的答案。

### 输出文件格式

程序会生成以下格式的 Word 文档：

```
第 247 条
题目：列车应按___、___和___的编组。
答案：本规程、列车编组计划、编挂条件

第 248 条
题目：动车组为___。单组动车组___...
答案：固定编组、运用状态下不得解编
```

## 工作原理

1. **解析文档** - 读取 Word 文档的所有段落
2. **识别条文** - 通过"第XXX条"模式识别每条规章的开始
3. **提取高亮** - 识别文本中的黄色高亮部分
4. **生成输出** - 创建新的 Word 文档，包含题目和答案

## 故障排除

### 问题：找不到高亮文本

**解决方案**：
- 确保你的 Word 文档中使用的是**黄色高亮**（不是其他颜色）
- 检查高亮是否正确应用到文本上

### 问题：条文识别不正确

**解决方案**：
- 确保条文号的格式为"第 XXX 条"或"第XXX条"
- 条文号应该在段落的开头

### 问题：导入错误

**解决方案**：
- 确保已安装所有依赖：`pip install -r requirements.txt`
- 确保使用的是 Python 3.6+

## 技术栈

- **python-docx** - Word 文档处理库

## 许可证

MIT

## 作者

Created with ❤️ by GitHub Copilot
