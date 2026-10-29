# yueji-chuangyi · 短视频创意预览稿生成器

把"和 AI 聊完的零散创意" → **客户能看的预览稿 PDF**（9 模块结构）。

## 工作流位置

```
创意对话 → [yueji-chuangyi] → 客户预览稿 PDF → 客户确认 → [yueji-geshizhuanhua] → 制作版 xlsx
```

- 上游：你和 AI 的创意对话（草稿、想法、台词片段）
- 本 skill 输出：客户决策用的预览稿 PDF（HTML 是中间产物）
- 下游：客户确认后，用 `yueji-geshizhuanhua` 出拍摄用的 xlsx 分镜脚本

## 9 模块结构

| # | 模块 | 条件 |
|---|---|---|
| 1 | 项目卡片 | 必 |
| 2 | 创意一句话 | 必 |
| 3 | 形式说明 | 必 |
| 4 | 角色设定 | 必 |
| 5 | 参考视频 | 用户给了才写 |
| 6 | 叙事流程图 | 必 |
| 7 | 分镜简版 | 必 |
| 8 | 权益占位图 | 用户给了才写 |
| 9 | 待确认项 | 必 |

详细规范见 `references/format-spec.md`。

## 输出形式：PDF（看效果）+ Word（改文字）

| 文件 | 用途 |
|---|---|
| `preview.pdf` | **主交付**。客户在微信里可直接预览，跨平台一致，可打印归档 |
| `preview.docx` | **改稿用**。客户想在上面改文字时给这份——表格是真表格，标题是真标题，待确认项是 ☐ 清单 |

PDF 改不动字，客户说"我想改改文字"时就要出 Word。**PDF 是主交付，Word 不替代它。**

```bash
bash scripts/html_to_pdf.sh preview.html preview.pdf   # Chrome headless
python3 scripts/html_to_docx.py preview.html           # 零依赖，同目录出 .docx
```

> `html_to_docx.py` 是专用转换器，只用标准库，**不需要 pip install**。
> 别用 pandoc / html-for-docx 这类通用工具——它们解析不了预览稿的 flex 布局，会把表格整块丢掉。

## 安装

```bash
# 任意 agent 平台通用安装路径
git clone https://github.com/zds07/yueji-chuangyi.git \
  ~/.workbuddy/skills/yueji-chuangyi
# 或
git clone https://github.com/zds07/yueji-chuangyi.git \
  ~/.claude/skills/yueji-chuangyi
# 或
git clone https://github.com/zds07/yueji-chuangyi.git \
  ~/.codex/skills/yueji-chuangyi
```

**前置依赖**：Google Chrome（仅转 PDF 需要；转 Word 不需要任何依赖）

## 使用

把创意对话笔记给 Agent，让它按 SKILL.md 的流程生成预览稿。触发示例：

- "把刚才聊的创意做成客户预览稿"
- "客户还没拍板，先出一版预览稿看看"
- "出个给客户的提案"

PDF 转完直接给客户路径，**不部署成网页链接**。客户要改字时，把同名 .docx 一起发过去。

## 文件结构

```
yueji-chuangyi/
├── SKILL.md                     # skill 主入口：触发条件、9 模块、流程
├── README.md                    # 本文件
├── scripts/
│   ├── html_to_pdf.sh           # Chrome headless 转 PDF 脚本
│   └── html_to_docx.py          # HTML → 可编辑 Word（零依赖，只用标准库）
├── references/
│   └── format-spec.md           # 9 模块的详细规范
└── examples/
    ├── template.html            # 通用模板（带 [xxx·待客户提供] 占位）
    └── README.md                # 模板使用说明
```

## License

MIT — 自由使用、修改、商用。引用时附仓库地址即可。

## 关于「开源脱敏版」

本仓库是 **脱敏开源版**：只包含规则、结构、脚本与模板，**不包含任何客户项目数据**。

- ❌ 删掉了：真实客户脚本文件名、商户名 / 门店名、人物角色名、地名、地区活动名
- ✅ 保留了：9 模块预览稿结构、全部 CSS 样式、HTML→PDF 转档脚本、通用模板 `examples/template.html`
- 📌 `references/format-spec.md` 里的示例一律用「A 店 / B 店 / XX 产品」这类通用代号

**产出效果一致**：拿 `examples/template.html` 填自己的内容，转出来的 PDF 与原版完全相同。
想要更强的效果，把自己项目里的历史脚本丢进本机私有副本的 `references/` 目录即可——那是你自己的东西，不上公网。
