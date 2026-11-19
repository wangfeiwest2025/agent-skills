# WorkBuddy Agent Skills

一组面向**中文政务 / 执法检查 / 公文写作**场景的 Agent Skill，可直接放入 WorkBuddy（或任何兼容
`SKILL.md` 约定的 Agent）的 skills 目录使用。

每个 skill 都是「一份 SKILL.md（做什么 + 踩过什么坑）+ 可执行脚本」的组合，
脚本均为纯 Python 标准库或公开 PyPI 依赖，不依赖任何私有服务。

## 技能清单

| 技能 | 用途 | 依赖 |
|---|---|---|
| [`jzsc-market-behavior-check`](skills/jzsc-market-behavior-check/) | 建筑市场行为检查（非现场）：从「合同用工备案花名册」提取项目经理与安全员，调用北京市建设行业职业人员信息查询平台核验其执业资格与所属单位，按住建部建质规〔2025〕3 号判定安全员配备是否充足，输出分项目检查报告 | Python 3.10+（纯标准库） |
| [`md-to-docx-cn-report`](skills/md-to-docx-cn-report/) | 把 Markdown 报告转成中文公文格式的 Word（.docx）：A4 版式、宋体正文/黑体标题、表格边框与表头底纹、页脚页码域、Word 导航窗格可用的标题层级 | `python-docx` |

## 安装

把 `skills/` 下的目录整体复制到你的 Agent 技能目录即可：

```bash
# WorkBuddy（用户级，所有项目可用）
cp -r skills/* ~/.workbuddy/skills/

# Windows
xcopy /E /I skills\jzsc-market-behavior-check %USERPROFILE%\.workbuddy\skills\jzsc-market-behavior-check
```

技能目录内保持 `SKILL.md` 位于根、脚本位于 `scripts/` 即可被识别。

## 各技能简介

### jzsc-market-behavior-check

面向在建工程项目的**非现场执法检查**，核查两件事：

1. **项目经理是否配备、是否为本单位人员** —— 在「注册人员」库查建造师注册，
   在「安管人员」库查 B 证（项目负责人）。
2. **安全员是否配备充足、是否为本单位人员** —— 按建质规〔2025〕3 号第二十六、二十七条
   判定数量是否达标，并核验其安全生产考核合格证是否登记在用工单位名下。

判定采用**三分法**，可区分三种情形，避免把「无证」和「证挂他处」混为一谈：

| 情形 | 判定 |
|---|---|
| 本单位名下有记录 | ✓ 本单位人员且持证 |
| 本单位名下无，但全库同名者证书在**其他单位** | ✗ 证在他单位（疑似挂靠） |
| 本单位名下无，且全库无任何同名记录 | ✗ 无证 |

内含两个实测踩到的坑的处理方式：花名册常见为**网页导出的 HTML 伪装成 `.xls`**（pandas/xlrd 打不开），
以及查询平台是 **LIKE 模糊匹配**（必须「姓名 + 企业名称」双条件收紧）。

```bash
python skills/jzsc-market-behavior-check/scripts/market_check.py <花名册目录> out.json
```

### md-to-docx-cn-report

把已经写好的 Markdown 报告转成可直接打印上报的 Word 文档。

```bash
python skills/md-to-docx-cn-report/scripts/md2docx.py 报告1.md 报告2.md
# → 输出到 <md 所在目录>/Word版报告/
```

内置排版：A4、页边距 2.54/3.17cm、正文宋体 12pt 行距 1.5、标题黑体逐级递减、
表格 Table Grid + 表头底纹、页脚「第 X 页 共 Y 页」。

处理了两个 python-docx 的经典坑：

- **有序列表会跨章节连号**（默认样式链到固定 `numId`，全文共享计数器）→ 改用字面序号 + 悬挂缩进
- **列宽容易超过页面可用宽度**（14.66cm）→ 按各列内容长度按比例分配

## 环境说明

两个技能都在 Windows + 中文环境下实测可用，但代码本身不依赖 Windows：

- 输出中文时建议设置 `PYTHONIOENCODING=utf-8`，否则 GBK 控制台会 `UnicodeEncodeError`
- 长任务请用「单次调用内轮询」的方式跑，避免后台进程被父进程回收

## 数据与隐私

本仓库**不含任何真实个人信息**。技能中的示例均为脱敏后的通用占位
（如「张伟」「某总包单位」），不含真实姓名、身份证号、手机号或项目数据。

相关平台均为政府公开查询服务，请在各平台的服务条款与所在地法规允许的范围内使用。

## License

[MIT](LICENSE)
