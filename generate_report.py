#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SRC 漏洞报告 .docx 生成器（零依赖，仅用标准库）

读取一份 JSON 规格，按标准 Word 漏洞报告模板产出 .docx，固定四段式：

    一、漏洞危害
        1、<危害/结论子标题>          （Heading2）
        <现象描述段>                  （正文）
        <危害说明段>                  （正文）
    二、漏洞url位置：
        <系统名>：<url>               （Heading2）
    三、复现步骤
        1、...  2、...                （正文，编号已写在文本里）
    四、修复建议
        1、...  2、...                （正文）

版式与标准报告范本逐项一致：
  A4 纵向、上下页边距 1440 twips、左右 1797 twips；
  正文 等线(DengXian) 10.5pt、行距 1.5 倍(line=360 auto)、段后 120；
  一级标题 加粗 22pt、段前 280 / 段后 160；二级标题 加粗 16pt、段前 200 / 段后 120；
  无封面、无目录、无表格、无页眉页脚页码。

JSON 结构：
{
  "title": "报告标题（可选，仅写入文档属性，不渲染进正文）",
  "harm":  [ {"subtitle": "1、危害子项/结论标题", "detail": "危害描述正文"}, ... ],
  "urls":  [ {"label": "系统或接口说明（可选）", "url": "https://..."}, ... ],
  "steps": [ "1、复现步骤一", "2、复现步骤二" ],
  "fixes": [ "1、修复建议一", "2、修复建议二" ]
}

说明：
  - 正文首段即「一、漏洞危害」，不生成居中大标题（与范本一致）；`title` 只写入 docProps。
  - 「危害等级 / 漏洞类型」不进正文（范本无该章节），如需提级请在标题文字里体现。
  - 所有文本入库前统一清洗非法字符（控制字符 / 零宽字符 / BOM / 替换符）。

用法：
  python generate_report.py report.json -o 输出.docx
  cat report.json | python generate_report.py -o 输出.docx
"""
import json
import sys
import re
import zipfile
import datetime

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# 非法字符清洗：
#   C0 控制符（去掉 \t\n\r 之外的全部）/ C1 控制符 / 零宽与方向控制符 /
#   BOM 与替换符 / 非字符码位
_ILLEGAL = re.compile(
    "[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f-\u009f"
    "\u200b-\u200f\u202a-\u202e\u2060-\u206f\ufeff\ufffc\ufffd\ufffe\uffff]"
)


def sanitize(s):
    """剔除非法/不可见字符，避免办公软件打开报错或提交平台校验失败。"""
    if not isinstance(s, str):
        s = str(s)
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = _ILLEGAL.sub("", s)
    s = re.sub(r"[ \t]*\n[ \t]*", " ", s)
    s = re.sub(r"[ \t]{2,}", " ", s)
    return s.strip()


def esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;"))


def run(text):
    return '<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % esc(sanitize(text))


def para(text="", style=None, body=False):
    """style：Heading1 / Heading2；body=True 时显式写 1.5 倍行距（对齐范本正文段）。"""
    ppr = ""
    if style:
        ppr = '<w:pPr><w:pStyle w:val="%s"/></w:pPr>' % style
    elif body:
        ppr = '<w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>'
    return "<w:p>%s%s</w:p>" % (ppr, run(text))


def as_list(v):
    if isinstance(v, list):
        return v
    if v:
        return [v]
    return []


def build_body(data):
    parts = []
    # 一、漏洞危害
    parts.append(para("一、漏洞危害", style="Heading1"))
    for item in data.get("harm", []):
        sub = item.get("subtitle", "")
        if sub:
            parts.append(para(sub, style="Heading2"))
        for d in as_list(item.get("detail", "")):
            if d:
                parts.append(para(d, body=True))
    # 二、漏洞url位置：
    parts.append(para("二、漏洞url位置：", style="Heading1"))
    for u in data.get("urls", []):
        label = sanitize(u.get("label", ""))
        url = sanitize(u.get("url", ""))
        line = ("%s：%s" % (label, url)) if label else url
        if line:
            parts.append(para(line, style="Heading2"))
    # 三、复现步骤
    parts.append(para("三、复现步骤", style="Heading1"))
    for s in data.get("steps", []):
        if s:
            parts.append(para(s, body=True))
    # 四、修复建议
    parts.append(para("四、修复建议", style="Heading1"))
    for f in data.get("fixes", []):
        if f:
            parts.append(para(f, body=True))
    parts.append(
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1797" w:bottom="1440" w:left="1797" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>'
    )
    return "".join(parts)


STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="%s">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="DengXian" w:eastAsia="DengXian" w:hAnsi="DengXian" w:cs="Times New Roman"/>
      <w:sz w:val="21"/><w:szCs w:val="21"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>
    <w:pPr><w:spacing w:line="306" w:lineRule="auto" w:before="0" w:after="120"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="DengXian" w:eastAsia="DengXian" w:hAnsi="DengXian"/><w:color w:val="000000"/><w:sz w:val="21"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="280" w:after="160"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="SimSun" w:hAnsi="Calibri"/><w:b/><w:bCs/><w:color w:val="000000"/><w:sz w:val="44"/><w:szCs w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="120"/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="SimSun" w:hAnsi="Calibri"/><w:b/><w:bCs/><w:color w:val="000000"/><w:sz w:val="32"/><w:szCs w:val="26"/></w:rPr>
  </w:style>
</w:styles>''' % W

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''


def build_doc(data):
    body = build_body(data)
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<w:document xmlns:w="%s"><w:body>%s</w:body></w:document>') % (W, body)
    now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    title = sanitize(data.get("title", ""))
    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>%s</dc:title><dc:creator>SRC Report Generator</dc:creator>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
            '</cp:coreProperties>') % (esc(title), now, now)
    app = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
           '<Application>src-vuln-report-docx</Application></Properties>')
    return {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": RELS,
        "word/document.xml": document,
        "word/styles.xml": STYLES,
        "word/_rels/document.xml.rels": DOC_RELS,
        "docProps/core.xml": core,
        "docProps/app.xml": app,
    }


def main():
    import argparse
    ap = argparse.ArgumentParser(description="SRC 漏洞报告 .docx 生成器（四段式）")
    ap.add_argument("json", nargs="?", help="JSON 规格文件路径（缺省读 stdin）")
    ap.add_argument("-o", "--output", required=True, help="输出 .docx 路径")
    args = ap.parse_args()
    if args.json:
        with open(args.json, "r", encoding="utf-8") as f:
            data = json.load(f)
    else:
        data = json.load(sys.stdin)
    parts = build_doc(data)
    with zipfile.ZipFile(args.output, "w", zipfile.ZIP_DEFLATED) as z:
        for name, content in parts.items():
            z.writestr(name, content)
    print("written: %s" % args.output)


if __name__ == "__main__":
    main()
