# 试卷分析报告（exam-analysis-report）

一次考试，给每个学生出一份 A4 分析报告，另加一页班级卷面分析。

老师只需要填**一份 Excel**（两个 Sheet），其余交给脚本：收表校验 → 逐人点评 → 排版打印 → 自动质检。
不起服务、不联网、不装第三方库。

---

## 安装

### 方式一：直接上传 zip（最省事，推荐给同事）

1. 打开 WorkBuddy，左侧点「专家 · 技能 · 连接器」→ 进「技能」面板；
2. 右上角「添加技能」→「上传技能」，把 `exam-analysis-report-v2.4.1.zip` **原样拖进去**（不要先解压）；
3. 等系统校验完，确认弹窗点允许；
4. **重开一个对话**——技能列表是会话启动时加载的，当前对话认不到新装的包。

### 方式二：手动解压到技能目录

1. 解压这个包，得到 `exam-analysis-report/` 目录；
2. 整个目录放进 WorkBuddy 的 skills 目录：

| 系统 | 目录 |
|---|---|
| Windows | `%USERPROFILE%\.workbuddy\skills\`（即 `C:\Users\<你>\.workbuddy\skills\`） |
| macOS / Linux | `~/.workbuddy/skills/` |

3. 确认路径是 `.../skills/exam-analysis-report/SKILL.md`（**不要把 SKILL.md 直接摊在 skills 下**）；
4. 重开一个对话，说「做这次试卷分析」就会触发；也可以在技能列表里看到它。

> 项目级安装：放到 `<工作区>/.workbuddy/skills/` 下，随仓库提交给同事共享。

### 装完先自验

```bash
cd <skills 目录>/exam-analysis-report
python scripts/selftest.py     # 环境自检，五项全过
python scripts/demo.py         # 用 3 个虚构学生跑通全链路，能出 3 份 PDF 就算装好
```

`demo.py` 不写你任何真实数据，输出在系统临时目录，跑完会打印路径。

## 环境要求

| 项 | 要求 | 说明 |
|---|---|---|
| Python | 3.9 以上 | 只用标准库，不需要 pip install |
| 浏览器 | Microsoft Edge 或 Chrome | 用来把 HTML 打成 PDF；找不到时设环境变量 `EDGE_PATH` 指向可执行文件 |
| 网络 | 不需要 | 全程离线 |

先跑一次自检，五项全过再开始（加 `--pdf` 是六项，连打印链路一起验）：

```bash
python scripts/selftest.py          # 查依赖、素材、Excel 读写
python scripts/selftest.py --pdf    # 连打印链路一起验
```

> `python` 不在 PATH 上时，用 WorkBuddy 自带的解释器绝对路径，在
> `C:/Users/<你>/.workbuddy/binaries/python/versions/` 下找**任意版本**目录里的 `python.exe`
> （版本号每台机器不同，别照抄）。脚本会自己找到同目录的其它模块，不需要设 PYTHONPATH。

## 第一次用（老师视角）

1. 把 `assets/试卷分析模板.xlsx` 发给老师填——两个 Sheet：**试卷信息**（年级/班级/科目/满分 + 逐题 10 列）和**学生成绩**（学号/姓名/缺考 + 逐题得分 + 学生备注）；
2. 老师填完回传，可以顺手把原卷 PDF 一起给（可选）；
3. 收表：

   ```bash
   python scripts/ingest_excel.py --xlsx "老师给的表.xlsx" --out "期末-八年级物理" --paper "原卷.pdf"
   ```

   有错会指到单元格（如 `学生成绩!J4`），照提示改完再跑；
4. 出报告骨架 + 点评证据包：

   ```bash
   python scripts/build_reports.py --exam-dir "期末-八年级物理" --no-pdf
   ```

5. 读 `reports/点评证据包.json` 写 `panels.json`（逐人一段点评，规范见 `references/comment-prompt.md`）；
6. 正式出报告并质检：

   ```bash
   python scripts/build_reports.py --exam-dir "期末-八年级物理"
   python scripts/build_class_panel.py --exam-dir "期末-八年级物理"
   python scripts/verify.py --exam-dir "期末-八年级物理"
   ```

`verify.py` 必须全绿才算交付（14 项：分值、页数、字数、禁用词、数字来源、备注直抄等）。

## 机构 Logo 与主题色（可选）

本包**不含任何机构 Logo**：不给就顶部留空，报告用默认蓝紫主题；给了就换成你的。

- **Logo**：把校徽图片（png / jpg / svg）交给 AI 或自己传参——页眉显示，等比缩放；
- **主题色**：默认从 Logo 里自动取主色，报告的标题、分数、图表、表格线全部跟着变；
- **想自己定色**：指定任意 `#RRGGBB`；想回默认蓝就清掉配置。

```bash
# 收表时一起设
python scripts/ingest_excel.py --xlsx "老师给的表.xlsx" --out "期末-八年级物理" \
       --logo "校徽.png"                    # 或 --theme "#C8102E"

# 已经收过表，只改 Logo / 颜色（不用重新收表，重跑 build_reports 就生效）
python scripts/brand.py --exam-dir "期末-八年级物理" --logo "校徽.png"
python scripts/brand.py --exam-dir "期末-八年级物理" --theme "#1F6F5C"
python scripts/brand.py --exam-dir "期末-八年级物理" --show
python scripts/brand.py --exam-dir "期末-八年级物理" --clear
```

自动取色会按白纸印刷可读性把颜色压深（保证对比度），不会照搬 Logo 里的浅色；
Logo 建议用透明背景的横版 PNG。JPG 也能用（会先转一次 PNG 再取色），失败时脚本会提示改用 `--theme`。

## 包里有什么

```
exam-analysis-report/
├── SKILL.md                    给 AI 的执行说明（流程、口径、坑）
├── README.md                   本文件，给人看
├── LICENSE.md                  许可全文（PolyForm Noncommercial 1.0.0 + 教育使用补充授权）
├── assets/
│   ├── 试卷分析模板.xlsx          老师填的空白模板
│   ├── report.html             学生报告版式（@@TOKEN@@ 占位）
│   ├── class_panel.html        班级卷面分析版式
│   └── error-tags.json         8 类通用错因
├── scripts/
│   ├── selftest.py             环境自检
│   ├── demo.py                 示例数据一键跑通（装完自验）
│   ├── brand.py                机构 Logo 与主题色（自动取色）
│   ├── make_template.py        生成空白模板（题数可调）
│   ├── ingest_excel.py         收表 + 单元格级校验 + 原卷归档
│   ├── evidence.py             生成点评证据包（数字白名单）
│   ├── build_reports.py        统计 + 图表 + 排版 + 分页 + 打印 + 打包
│   ├── build_class_panel.py    班级卷面分析一页
│   ├── verify.py               14 项质检
│   └── xlsxlite.py             纯标准库 xlsx 读写
├── references/
│   ├── comment-prompt.md       点评写作规范
│   ├── teacher-guide.md        给老师的半页说明（可直接打印）
│   ├── paper-schema.md         paper.json 字段
│   └── troubleshooting.md      报错与排障
└── demo/
    ├── 示例-八年级物理.xlsx       3 个虚构学生 / 6 题的填好样例
    └── panels.json             对应的成品点评（可直接复用看效果）
```

## 常见问题

| 现象 | 处理 |
|---|---|
| 技能没被触发 | 确认目录层级是 `skills/exam-analysis-report/SKILL.md`，然后重开对话 |
| 提示找不到 Edge/Chrome | 装 Edge，或 `set EDGE_PATH=C:\...\msedge.exe` 后重跑 |
| 提示 PDF 正被其他程序占用 | 关掉阅读器/预览窗口；不关会打出旧报告 |
| 提示「题目分值合计 ≠ 卷面满分」 | 先核对漏题或抄错；确认卷子本身如此，就把卷面满分改成合计值 |
| Excel 是被改过的旧表 | 表头文字必须认得（题号/分值/模块/知识点/错因/页码/建议/备注），列序可变、列名不可变 |
| 想改题数 | `python scripts/make_template.py --questions 30 --students 60` |
| 报告顶部没有 Logo | 本包不带 Logo，属正常；给了 `--logo` 才会显示 |
| 自动取到的颜色不是想要的那一色 | 多色 Logo 按面积和饱和度取最显眼的一色；想精确指定就用 `--theme "#RRGGBB"` |
| 换了 Logo / 颜色，报告没变 | 配置存在 `paper.json`，要重跑一遍 `build_reports.py` 才生效 |

更多见 `references/troubleshooting.md`。

## 许可与素材

**许可**：PolyForm Noncommercial 1.0.0 + **教育使用补充授权**（全文见 `LICENSE.md`）。

- **可以**：教师在个人教学、学校教学、教培活动中免费使用（含批量出报告）；为教学目的修改、研究、分享；
  经作者授权的开源项目集成。
- **需要另行授权**：集成进闭源商业软件、做成收费 SaaS / 商业 API / 商业智能体服务，
  或用服务器托管、插件封装等方式绕开上述限制。
- 再分发时请随包保留 `LICENSE.md`。

**素材**：包里没有 Logo、没有真实师生姓名、没有真实考试数据；Logo 由使用者提供，用前确认授权；
主题色只写进你自己考试的 `paper.json`，不写进模板。`demo/` 里的学生全部虚构，仅用于装完自验。

## 版本

| 版本 | 变化 |
|---|---|
| 2.4.1 | 明确许可：PolyForm Noncommercial 1.0.0 + 教育使用补充授权，新增 `LICENSE.md` 并随包分发；`license` 字段与 README 同步 |
| 2.4.0 | 公开版：**移除机构 Logo**，包内不含任何机构标识；新增 `scripts/brand.py`——可上传机构 Logo（png/jpg/svg），自动从 Logo 取主色并套用到整份报告，也可 `--theme` 指定颜色（默认蓝紫 #3F3B92）；机构色接近失分红时图表自动分档；`selftest.py` 在 GBK 控制台不再崩 |
| 2.3.2 | 彻底优化个人点评“讲人话”：确立任课老师面对学生与家长的真诚交流口吻，以积极肯定闪光点和基本盘为基调，重构学情诊断与学习方法指导（拒绝法医病历体、拒绝刺耳语言、拒绝机械念实验操作规程）；提供系统化分层分档（A/B/C/D/E及个性化备注）标杆范文；引入子智能体干净上下文生成流程约束，减少上下文污染，提升全班点评质量稳定性与人情味 |
| 2.3.1 | 重写个人点评规范与示例：明确任课教师口吻、得分事实与个人错因的边界，禁止由整题得分虚构作答细节；证据包同步增加推断、训练量和文风自检约束，demo 点评改为自然、可执行的表述 |
| 2.3.0 | 补「上传技能」导入路径与装完自验；新增 `scripts/demo.py` 与 `demo/` 示例（3 个虚构学生，一键跑通全链路）；SKILL.md 补安装与依赖说明、删掉指向包外目录的引用、Python 解释器改为分步查找不再写死版本号；浏览器查找支持 macOS/Linux 与 PATH；`logo.png` 改为可缺省；等级口径可在 `paper.json` 用 `grades` 覆盖（文档补充） |
| 2.2.0 | 旧链路（网页标注 / 旧 xlsm 导入 / 旧话术 / 知识点树）整体移出包外，分享包只留主流程，包体 392KB/32 文件 → 120KB/19 文件；脱敏闸门扩到学科代号与班型；logo 改名通用名 |
| 2.1.1 | 隐私清理：移除旧链路文档里的真实班级对账表、示例里的真实师生姓名与校区；打包脚本加脱敏闸门（命中真实姓名/班名/机构即拒绝打包） |
| 2.1.0 | 报告去掉全部小字说明；页眉显示试卷名称、信息行补校区/考试日期；试卷名称/日期/校区/复习页码题号可留空；修 Excel 留空时把标签读成值的 bug；修 PDF 被占用时静默出旧报告的 bug；新增 `selftest.py` 与本 README |
| 2.0.0 | Excel 收表取代网页录入：两 Sheet 模板、单元格级校验、点评证据包与 14 项质检、新版两页报告与动态分页 |

检查版本：`SKILL.md` frontmatter 的 `version` 字段。
