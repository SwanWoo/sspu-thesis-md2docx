# 毕业论文 Markdown 转 Word 工具

将 Markdown 格式的毕业论文草稿转换为符合 SSPU 模板格式的 `.docx` 文件。

## 环境要求

- Python >= 3.10
- [uv](https://docs.astral.sh/uv/) — Python 包管理器
- XeLaTeX（TeX Live 2025+）— 公式渲染
- macOS / Linux

## 安装

```bash
uv sync
```

确保系统已安装 XeLaTeX：

```bash
brew install --cask mactex
xelatex --version
```

## 使用

```bash
uv run python thesis2docx.py your_paper.md
```

生成的文件：`毕业论文_生成.docx`。用 Word 打开后会自动更新目录页码。

## 项目结构

```text
thesis2docx.py          # 主转换脚本
final_paper.md          # 论文 Markdown 源文件
template/               # 大学 Word 模板（已解压，勿修改）
  word/
    document.xml        # 模板正文 XML
    styles.xml          # 样式定义
    media/              # 模板原始图片 + 运行时生成的公式/图片
pyproject.toml          # uv 项目配置
```

## 工作原理

脚本直接操作 `.docx` 的 OpenXML 结构（ZIP 内的 XML 文件）：

1. 解析 Markdown 为结构化块（标题、段落、表格、图片、公式、列表等）
2. 加载模板 `template/word/document.xml`，注入封面信息、替换摘要、生成目录
3. 将正文内容按样式映射转换为 Word 段落
4. LaTeX 公式通过 XeLaTeX + PyMuPDF 渲染为高清 PNG 嵌入
5. 网络图片自动下载并嵌入
6. 打包输出为新的 `.docx` 文件
