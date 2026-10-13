# SRC 漏洞提交报告规范模板

把 SRC / 众测漏洞挖掘结果，整理成标准化"四段式" Word(`.docx`) 报告模板的技能（skill）。

> 模板源自一批实战评测报告，章节结构、版式与内容取舍程度已逐项对齐。repo 内示例为占位数据（example.com），可直接替换。

## 报告模板（固定四段）

1. **一、漏洞危害** — `1、结论子标题`（二级标题）+ 现象段 + 危害段
2. **二、漏洞url位置：** — 每行 `系统名：URL`
3. **三、复现步骤** — `1、2、3…`
4. **四、修复建议** — `1、2、3…`

正文不设居中大标题，也**不设**「危害等级 / 漏洞类型」章节——等级与类型直接体现在报告标题里。

## 版式

A4 纵向；上下边距 1440 twips、左右 1797 twips；正文等线(DengXian) 10.5pt、1.5 倍行距；
一级标题宋体加粗 22pt（段前 280 / 段后 160）、二级标题宋体加粗 16pt（段前 200 / 段后 120）；
无封面、无目录、无表格、无页眉页脚页码。全篇保持精简（范本约 1300 字）。

## 快速开始

```bash
# 准备一份 JSON（参考 examples/sample.json）
python generate_report.py examples/sample.json -o "内部门户公网匿名暴露漏洞.docx"
```

或读 stdin：

```bash
cat examples/sample.json | python generate_report.py -o out.docx
```

## JSON 规格

```json
{
  "title": "报告标题（可选，仅写入文档属性，不渲染进正文）",
  "harm":  [ {"subtitle": "1、结论/危害子标题", "detail": "危害描述（字符串或多段数组）"} ],
  "urls":  [ {"label": "系统说明（可选）", "url": "https://..."} ],
  "steps": [ "1、复现步骤一", "2、复现步骤二" ],
  "fixes": [ "1、修复建议一", "2、修复建议二" ]
}
```

字段说明：

| 字段 | 必填 | 说明 |
|------|------|------|
| `title` | 否 | 仅写入文档属性；不渲染进正文 |
| `harm` | 是 | 危害列表，每项 `{subtitle, detail}`；`detail` 可为字符串或字符串数组 |
| `urls` | 是 | 位置列表，每项 `{label, url}`，label 可省略 |
| `steps` | 是 | 复现步骤，编号正文数组 |
| `fixes` | 是 | 修复建议，编号正文数组 |

## 非法字符

生成器对所有入库文本统一清洗：C0/C1 控制字符、零宽与方向控制符、BOM、替换符。
正文中也不要写 `#`、半角 `<` `>`（会破坏提交平台校验），需要时改用全角或改写。

## 作为 WorkBuddy 技能安装

把本仓库内容放到 `~/.workbuddy/skills/src-vuln-report-docx/` 即可被 WorkBuddy 识别。

## 依赖

**无**。仅用 Python 3 标准库（`zipfile` / `json` / `re` / `argparse` / `datetime`），无需 `pip install`。
