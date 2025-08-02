# AI Content to Word Converter

这是一个专门用于将AI生成的Markdown内容转换为Word文档的工具。

## 功能特点

1. 从剪贴板读取Markdown内容
2. 自动检测内容是否为Markdown格式
3. 自动转换LaTeX数学公式（将```math块转换为Word可识别的格式）
4. 生成带时间戳的Word文档
5. 将生成的文件路径复制到剪贴板

## 使用方法

### 简单使用

1. 复制AI生成的Markdown内容到剪贴板
2. 双击运行 `convert_clipboard_to_word.bat` 文件
3. 等待转换完成
4. 生成的Word文档路径会自动复制到剪贴板
5. 在文件资源管理器中粘贴路径即可打开生成的Word文档

### 命令行使用

也可以通过命令行运行Python脚本：

```bash
python clipboard_to_word.py
```

## LaTeX公式支持

本工具支持以下LaTeX公式格式：

1. 行内公式：`$...$` 或 `\(...\)`
2. 块级公式：`$$...$$` 或 `\[...\]`
3. 代码块公式：```math ...```

转换规则：

- ```math 环境中的单行公式会被转换为 `$...$` 格式
- ```math 环境中的多行公式会被转换为 `$$...$$` 格式
- `\(...\)` 格式的行内公式会被转换为 `$...$` 格式
- `\[...\]` 格式的块级公式会被转换为 `$$...$$` 格式

## 生成文件

生成的Word文档将保存在当前目录下，文件名格式为：
`markdown_conversion_YYYYMMDD_HHMMSS.docx`

其中：

- YYYYMMDD 是日期（年月日）
- HHMMSS 是时间（时分秒）

## 系统要求

1. Windows操作系统
2. Python 3.6或更高版本
3. Pandoc（已安装并在系统PATH中）

## 安装依赖

如果尚未安装必要的Python包，请运行：

```bash
pip install pyperclip
```

## 故障排除

### 剪贴板内容不是Markdown格式

如果出现此错误，请确保剪贴板中的内容是Markdown格式。

### Pandoc未找到

确保Pandoc已正确安装并添加到系统PATH环境变量中。

### 转换后的公式在Word中显示不正确

尝试以下方法：

1. 检查原始Markdown内容中的公式格式是否正确
2. 确保使用了正确的LaTeX语法
