# 运行说明

## 使用

在 Codex 中说：“用艾宾浩斯式单词复习，上传这些单词，每天新增5个。”提供照片、文档或文本即可。识别复核由会话完成，排期和文件导出由本地程序完成。

命令行可直接使用已整理词表：

以下命令在 Skill 文件夹内执行。macOS/Linux 可用下面的快捷入口；Windows 使用 `py scripts/vocab.py` 替换 `bash scripts/run.sh`，或用已安装环境的 `python`。测试命令是 `py scripts/test_vocab.py`。独立环境先执行 `python -m pip install -r requirements.txt`。

```
bash scripts/run.sh build examples/词表示例.csv --daily 5 --out result
bash scripts/run.sh build examples/词表示例.csv --daily 10 --out result10
bash scripts/run.sh compare examples/词表示例.csv
bash scripts/run.sh extract 教材.pdf --out raw.txt
bash scripts/run.sh test
```

`build` 支持 JSON、带表头的 CSV/TSV、制表符文本和 XLSX。必需表头为 `english,chinese` 或 `英文,中文`；可选表头为 `ipa,pos` 或 `音标,词性`，顺序不限，无须凑足四列。JSON 格式：

```json
[
  {"english":"apple","chinese":"苹果","ipa":"/ˈæpəl/","pos":"n."},
  {"english":"get up","chinese":"起床","pos":"phr."},
  {"english":"pear","chinese":"梨"}
]
```

`extract` 支持照片（PNG/JPG/TIFF等）、PDF、DOCX、XLSX、EPUB、CSV/TSV/TXT/MD/JSON。HEIC/HEIF额外安装 `python -m pip install pillow-heif` 后可解码，也可先导出JPG。老式DOC/XLS先在办公软件另存为DOCX/XLSX。图片依赖本机Tesseract的eng和chi_sim语言包；没有OCR时由会话看图。PDF文字页直接抽取，扫描页逐页OCR，另需Poppler的pdftoppm在PATH中。EPUB按阅读顺序提取文本；固定版面/图片内容需提供对应页面或截图。不得绕过DRM。

OCR尤其容易读错音标，必须与原图核对。只有英文和中文必填，原文没有的音标/词性不自动补充，数据中省略该字段，Excel也不保留整组皆空的字段列。用户明确要求补音标且未指定口音时，才追问英式还是美式。

用户仅上传材料但未选方案时，先执行 `compare`，展示两种方案各自的组数、完整周期和实际每日峰值，再等待用户选5或10。不拿理论上限代替实际计算，不擅自默认。日期统一留空，不问开始日期；新版不再接受 `--start` 参数。

## 本地依赖与免费边界

优先使用 Codex 自带 Python、Node 和 artifact-tool。脚本自动查找运行时；也可设置 `VOCAB_PYTHON`、`VOCAB_NODE`、`VOCAB_ARTIFACT_MODULE`。独立运行时用 Python 3.10+ 安装 `requirements.txt`。Excel优先artifact-tool；不可用时使用开源XlsxWriter。本地生成过程不联网、不需要API密钥。照片识别可使用免费的本地Tesseract，也可由现有会话读取，宿主平台本身可能有费用。

可用 `VOCAB_XLSX_ENGINE=xlsxwriter` 显式选择独立导出引擎；默认 `auto`。PowerShell 设置方式为 `$env:VOCAB_XLSX_ENGINE = "xlsxwriter"`。这是开发和跨平台部署选项，家长无需选择。

PDF在macOS优先嵌入Arial Unicode字体，Windows尝试微软雅黑/宋体。其他系统可设置 `VOCAB_FONT` 指向支持中文、允许嵌入的TrueType字体；未找到时回退到PDF标准中文字体STSong-Light（阅读器需支持亚洲字体）。独立部署可安装适用的开源中文TTF并指定路径；并非所有OTF/TTC字体都能被ReportLab读取。Excel的字体随客户端字体环境而定。仓库不附带任何操作系统字体。

## 排版和统计

每组5次出现在首次日的0、2、6、14、29天后；理论每日上限是每天新增数×5，实际峰值按任务逐日计算。样例300词可分别达到25和50。没有任务的日期不会产生空白练习列，日号仍按真实日历连续计算。

PDF每列可用高度固定，按文字换行后的实际高度装入完整组；先用10pt中文字和22pt词行，必要时在保持可写空间的前提下用更紧凑的行距。常规5词方案峰值占2列，10词方案峰值占3列。列数由真实内容决定，不机械用词数阈值。空余列不回填未来日期，以保持时间顺序。

Excel采用6pt高的版面网格，合并单元格形成文字行；并非一个词占一个极薄网格行。每页126网格行，首行区标题后双区填充，手工分页保证不拆组。打印设置为A4纵向、100%缩放和明确打印区域；字号9pt。长词/长音标/长中文按字段宽度换行并增加行高。可以修改内容；修改到更长文字、增删组后应重新生成以重新分页。

版式按每组实际信息自适应：四字段用四列；只含音标或词性时用对应三列；仅中英文时用“中文｜英文｜中文｜英文”。双词对按先左后右的阅读顺序排列，10词5行，5词3行；尾组也不混入其他组。各组可以在同一页使用不同版式，缺失可选项不会被自动填充。内部Excel网格仅用于合并单元格排版，不是额外的数据字段列。

默认只生成两份文件，已有同名输出时拒绝覆盖，可明确加 `--overwrite`。可选 `--audit 路径.json` 导出排期和布局信息用于开发验证，家长不需要此文件。PDF始终打印日期空白线供手写。
