---
name: md-to-docx-cn-report
description: 把 Markdown 报告转成中文公文格式的 Word（.docx）——A4 版式、宋体正文/黑体标题、Table Grid 表格边框+表头底纹、页脚「第 X 页 共 Y 页」页码域、Word 导航窗格可用的标题层级。覆盖 python-docx 中文字体（w:eastAsia）设置、有序列表连号、列宽超页宽这两个必踩的坑。当用户说「生成 Word 版报告」「转成 docx」「md 转 word」「报告要交 Word 格式」「加页码」「公文格式排版」时使用。
agent_created: true
---

# Markdown → 中文公文格式 Word

把已经写好的 Markdown 报告（政府/执法/检查类，含大量表格）转成可直接打印、上报的 `.docx`。

## 依赖

```bash
C:\Users\admin\.workbuddy\binaries\python\envs\default\Scripts\pip.exe install python-docx
```
本机无 pandoc、无 LibreOffice、无 Word，**只能靠 python-docx 生成**，也没有办法本地渲染预览
→ 交付前用「读回 docx + 校验脚本」的方式自检（见下文）。

## 源文件是 HTML 时：先转 Markdown，别用正则切

若内容原本是 HTML，**必须先 HTML→MD 再跑本脚本**。这一段最容易踩坑，实测（2026-10-05）
把一个 50KB 的讲义 HTML 转 MD，连续失败 5 次才做对。**不要用正则切 HTML 边界**——
同层 `<div>` 之间无法可靠判断边界，会把提示框和卡片的文字互相吞掉（表现为
「字数变少 + 结构错位」而**不报错**，非常隐蔽）。

正确做法：用 `html.parser.HTMLParser` **建 DOM 树再递归渲染**，配合标签栈处理三类标签：

| 类别 | 标签 | 处理方式 | 踩坑后果 |
|---|---|---|---|
| 丢内容 | `style` `script` `title` | 内容整段丢弃，用计数器 | 计数器遇自闭合标签会撑爆，后面内容全被丢 |
| 只跳标签 | `html` `head` `body` | **内容必须保留**，只是不当容器 | 误放进"丢内容"类 → 整篇丢失，输出仅 1 字符 |
| 自闭合 | `meta` `link` | 直接忽略 | —— |

其他关键点：

1. **`handle_endtag` 回退靠 `parent` 指针**，所以 `Node.__slots__` 要留 `parent`，
   并在 `handle_starttag` 里就赋值（不能等解析完再补，那时回退已错）。
2. **容器 `div` 不能一律压成一行**。先判断有无块级子节点
   （`div/table/ul/ol/pre/h1-h4/p`）：有 → 递归子节点；无 → 才压成普通段落。
   否则外层 `wrap` 会把整篇正文挤成一行（表现为行数只有几十）。
3. **块内换行要显式保留**。`div.ln`（公式行）逐个 `split('\n')` 摊平，
   否则 149 行公式会挤成 1 行。
4. **`header`/`footer` 在渲染层跳过**（解析层保留，否则其闭合标签位置会影响 cur 回退）。
5. **带 class 的标题容器**（如 `div.blk-h` 里放 `<h3>板块名</h3>`）要**按子节点类型分发**：
   `h3` 转 `###` 前缀，纯文本另起一行。整块当纯文本会让所有板块标题消失。

转换后自查三指标（对照原 HTML 统计）：
- `.ln` 数应等于 HTML 里 `class="ln"` 的出现次数
- 输出行数应在数百量级（只有几十行 = 换行被吞）
- 标题数应等于 HTML 里 `<h1>`~`<h4>` 的数量

## 用法

```bash
python scripts/md2docx.py 报告1.md 报告2.md 报告3.md 报告4.md
#   → 输出到 <md所在目录>/Word版报告/报告N.docx

python scripts/md2docx.py 报告.md -o D:/某目录        # 指定输出目录
python scripts/md2docx.py 报告.md --flat             # 输出到 md 同目录
python scripts/md2docx.py 报告.md 输出.docx          # 指定完整输出文件
```

## 排版规格（已内置，不用改）

| 项 | 值 |
|---|---|
| 页面 | A4 21×29.7cm；页边距上下 2.54cm、左右 3.17cm（**可用宽度 14.66cm**） |
| 正文 | 宋体 12pt，行距 1.5，西文 Times New Roman |
| 一级标题 | 黑体 20pt 居中（`# `，报告主标题） |
| 二级标题 | 黑体 15pt（`## `，首个居中作副标题，其后左对齐） |
| 三级标题 | 黑体 13.5pt（`### `） |
| 表格 | Table Grid 边框，表头底纹 `DEEAF6` + 加粗，正文 10.5pt |
| 引用块 | 楷体 10.5pt 深灰 + 左侧灰竖线 |
| 图片 | `![题注](路径)` → 按页宽 14.4cm **居中**插入，题注楷体 10.5pt 居中；相对路径按 md 所在目录解析；文件缺失时插入红色 `［缺图：xxx］` 提示而不中断 |
| 页脚 | 居中「第 X 页  共 Y 页」（PAGE / NUMPAGES 域） |

## 图片

md 里直接写标准语法，题注会渲染成居中小字：

```markdown
![图1  2025年4月拆除承重墙类诉求「破坏形态」分布（全部 407 件）](chart_form_all.png)
```

**生成图表的建议**：优先 `pip install matplotlib -i https://pypi.tuna.tsinghua.edu.cn/simple`
（实测可装 3.11.x；会报 `SANDBOX PERMISSION DENIED: AppData\Local\matplotlib` 的缓存告警，
忽略即可，出图正常）。中文字体用
`fm.fontManager.addfont('C:/Windows/Fonts/msyh.ttc')` + `rcParams['font.family']='Microsoft YaHei'`。

### 图片"看不清"的真正原因（重要）

不是 dpi 不够，而是 **图像像素宽 ÷ Word 插入宽 的比值太大**。docx 按 96dpi 记录图片原始尺寸，
一张 1400px 宽的图相当于 37cm，被压到 14.4cm 时缩放 0.39 → 图内 10pt 字到纸面只剩 **4pt**。

**字号反算公式**（先算再画）：

```
纸面 pt = 图内 pt × 插入宽cm ÷ (figsize宽in × 2.54)
```

要让纸面 9~11pt（正文三号 16pt 的图注观感），插入 14.4cm、figsize 9.6in 时，
**图内字号必须 15~19pt**；刻度 14~15pt，轴标题 18pt，数值标注 16~19pt。

### 三条硬规矩

1. **图内不写标题**——图名放 Word 图题（`![图1 …](x.png)`），否则同一句话出现两遍。
   因此 `plt.title()` 一律不要调用。
2. **画布比例必须等于 Word 里的插入尺寸比例**，否则 Word 拉伸变形。做法：先按 figsize 出图，
   再用 PIL 居中补白到 `w_cm : h_cm`（只补白、不裁切），最后原样写回 `word/media/imageN`。
3. **替换既有 docx 的图**时，按 `word/media/` 里的原始成员名写回；**注意扩展名对应的真实格式**
   （原包里 `image2.jpeg` 是 JPEG，就必须存 JPEG，否则 Content-Type 不符，Word 可能报错）。

地图类图：北京 16 区区界取
`https://geo.datav.aliyun.com/areas_v3/bound/110000_full.json`（含 `centroid`），
`PatchCollection` 分级设色，`ax.set_aspect(1.305)` 校正 40°N 的经纬比。
**中心城区（东城/西城）面积极小，引线标注必然压盖**，改用"编号 + 右侧排序注记表"最稳。

词云：`wordcloud` + `jieba`，`prefer_horizontal=1.0`、`min_font_size` 按
`纸面下限pt ÷ (插入宽cm×28.35/像素宽)` 反算，`random_state` 多试几个挑版式；
关键词频次用**子串计数**（`txt.count(w)`）以便与正文数字严格对齐。

**图片必须在 md 同级目录**：脚本按 md 所在目录解析相对路径，**不认工作区外的绝对路径**。
生成图片时它们常落在临时目录（`C:\Users\admin\xxx\`），要先 `cp` 到 md 旁边再跑脚本，
否则会插入 `［缺图：xxx］` 红字提示。文件名含中文没问题，但**建议用中文可读名**
（如 `05-深底黑字不可读.png`）而不是 `s05.png`，交付给用户时省去解释。
图片会按页宽 14.4cm 居中，**截图务必用 1280×720 以上的分辨率**（对应 13.33×7.5in），
低分辨率截图放大后会糊。


## 两个必踩的坑（脚本已处理，改代码时别改回去）

### 坑 1：有序列表会跨章节连号

python-docx 默认模板里 `List Number` 样式链到固定 `numId=5`。
同一 numId 意味着**文档内所有有序列表共享一个计数器** ——
「七、处理建议」的 1~7 条之后，「八、核查说明」会接着显示 8、9、10，而不是重新从 1 开始。

**做法**：不用 Word 自动编号，直接把 Markdown 里的**字面序号**写成文本，
配悬挂缩进保证换行对齐：

```python
pf.left_indent = Cm(base + 0.62)
pf.first_line_indent = Cm(-0.62)   # 负值 = 悬挂
```

### 坑 2：列宽容易超过页面可用宽度

硬编码列宽（如 `[2.4, 4.0, 3.2, 4.6, 2.0]` 合计 16.2cm）会超过可用的 14.66cm，
Word 里表现为表格被截断或挤压。而且同一份文档里表格列数不一（2~7 列），一套宽度套不上。

**做法**：按每列**最大内容长度**按比例分配，中文按 2 个字符宽折算：

```python
def w(s): return sum(2 if ord(c) > 0x2E80 else 1 for c in s)
weights = [min(max(max(w(r[ci]) for r in rows), 4), 46) for ci in range(ncol)]
ws = [14.4 * x / sum(weights) for x in weights]   # 留 0.26cm 余量
# 再对 < 1.5cm 的列抬到 1.5cm，多出的宽度按权重从其他列扣回
```

### 其他要点

- **中文字体必须写 `w:eastAsia`**，只设 `run.font.name` 对中文无效：
  ```python
  rFonts.set(qn('w:eastAsia'), '宋体')
  ```
- **标题要用真 Heading 样式**（`style='Heading 1'`），否则 Word 导航窗格和目录都不认。
  但 Heading 样式自带主题蓝，必须显式覆盖颜色：
  ```python
  add_runs(p, txt, ea=HEI, size=20, bold=True, color=RGBColor(0, 0, 0))
  ```
- 模块级不要无条件重包 stdout；重包时加 **`write_through=True`**，否则没有进度输出。

## 交付前自检（本机无法渲染，必须做）

```python
from docx import Document
import re, zipfile
d = Document(f)
hs = [p.style.name for p in d.paragraphs if p.style.name.startswith('Heading')]
seq = [p.text for p in d.paragraphs if re.match(r'^\d+\. ', p.text.strip())]
assert len([s for s in seq if s.startswith('1. ')]) == 2        # 两个列表都从 1 重新开始
for t in d.tables:                                              # 没有超页宽的表
    assert sum(c.width.cm for c in t.rows[0].cells) <= 14.71
z = zipfile.ZipFile(f)
assert 'PAGE' in z.read('word/footer1.xml').decode('utf-8')      # 页码域在
assert 'NUMPAGES' in z.read('word/footer1.xml').decode('utf-8')  # 总页数域在
# 注意：package 里【有】word/numbering.xml —— 那是 python-docx 默认模板自带的部件
#（2026-10 实测 python-docx 1.2.0 生成的文件就在）。所以：
#   * 不要写 `assert z.read('word/numbering.xml')` 当作"列表编号正常"的证据；
#   * 也不要反过来断言"包内没有 numbering.xml"（旧版本这里就是这句，是错的）。
# "脚本不使用 Word 自动编号"指的是【列表段落没有挂 numPr】，与包内有无该部件无关。
```

**列表编号的正确验法是查"从 1 重开"**，不是查 numbering.xml：

```python
import re
seq = [p.text.strip() for p in d.paragraphs if re.match(r'^\d+\. ', p.text.strip())]
assert len([s for s in seq if s.startswith('1. ')]) == 文档里独立列表的个数   # 每个列表都从 1 开始
```

另外：**表格超过 14 列必然超页宽**（宽度按列数摊，每列至少要能放几个字）。
写 md 时遇到 10 列以上的横向数据表，先转成「两列并排」或「月份/数值」两列竖排，
否则 Word 里会挤成一团。本次一张 14 列的表算出 20.99cm，超出可用宽度 14.66cm 达 43%。

另外注意：**页脚读回的文本是「第  页  共  页」（域值空）**，这是正常的 ——
PAGE/NUMPAGES 由 Word 打开时才计算，不是生成错误。

### 坑 3：核对内容时只扫 paragraphs 会漏掉表格里的数据

报告类文档的关键数据（实测值、色号、坐标）往往**密集放在表格里**，
而 `d.paragraphs` 只包含段落级文本。2026-10 实测：核对 19 个关键数据点时，
只扫 paragraphs 报出 6 个「MISS」，实际它们全在 `d.tables` 里，属误报。

```python
all_txt = '\n'.join(p.text for p in d.paragraphs)
for t in d.tables:
    for r in t.rows:
        all_txt += '\n' + ' | '.join(c.text for c in r.cells)
# 再对 all_txt 做关键数据点断言
```

**判断某项检查是否「全部通过」时同理** —— 别把「我没扫到」当成「文档里没有」。
另外阈值类表述在文档里常写成 `≥3.0` 而非 `3.0:1`，直接搜 `3.0:1` 会误判为缺失。
