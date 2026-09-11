"""检查改稿结果：新内容写进去了吗？旧内容/模板占位符清干净了吗？

简历改数字时最容易出错 —— 明面上改了，别处还留着旧值；
从模板起稿时，模板自带的 your@email.com 之类占位符漏换，
HR 一看就直接淘汰，比任何技术问题都致命。

用法:
    # 只检查模板占位符
    python check_residue.py <简历.docx>

    # 同时检查新旧关键词
    python check_residue.py <简历.pdf> '{"must":["新关键词"],"must_not":["旧关键词"]}'

环境要求:
    pip install pymupdf     # 读 PDF
    pandoc                  # 读 docx 时需要
"""
import json
import os
import re
import subprocess
import sys

# assets/template.docx 里自带的占位符 —— 交付前必须一个不剩
TEMPLATE_PLACEHOLDERS = [
    "你的姓名",
    "目标方向（2027届）",
    "your@email.com",
    "your-name/your-repo",
    "起止时间",
    "【标签】",
    "【类别】",
    "【做了什么",
    "【技能清单",
    "【竞赛名称",
    "【奖学金",
    "【职能方向",
]


def load_text(path):
    """取文档全文。PDF 用 PyMuPDF，docx 走 pandoc。"""
    ext = os.path.splitext(path)[1].lower()

    if ext == ".pdf":
        import fitz
        doc = fitz.open(path)
        return "\n".join(page.get_text() for page in doc)

    if ext in (".docx", ".doc"):
        # 直接从 OOXML 取所有 <w:t> 文本节点，不走 pandoc。
        # 原因：pandoc 的 plain 输出会把超链接字段的显示文本整个丢掉，
        # 而邮箱、GitHub 链接恰恰就在超链接里 —— 用 pandoc 会漏检，
        # 报出"clean"这种最危险的假阴性。
        import zipfile
        from xml.etree import ElementTree as ET

        ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        chunks = []
        with zipfile.ZipFile(path) as z:
            for part in z.namelist():
                base = os.path.basename(part)
                if not base.startswith(("document", "footer", "header")):
                    continue
                if not part.endswith(".xml"):
                    continue
                try:
                    root = ET.fromstring(z.read(part))
                except ET.ParseError:
                    continue
                # <w:t> 是可见文本；<w:instrText> 是字段指令，
                # 超链接的 URL 就藏在里面（HYPERLINK "mailto:..."），只取 w:t 会漏掉。
                for tag in ("t", "instrText"):
                    for node in root.iter(ns + tag):
                        if node.text:
                            chunks.append(node.text)
        return "\n".join(chunks)

    with open(path, encoding="utf-8") as fh:
        return fh.read()


def check_metadata(path):
    """检查文档属性（docProps）。

    这块最隐蔽 —— 正文里什么都看不出来，但打开文件属性就能看到
    作者名、原标题、甚至模板来源。从别人的模板起稿时，
    这些字段会原样带过来，等于在简历里夹带了一张便签。
    """
    if not path.lower().endswith((".docx", ".doc")):
        return []

    import zipfile
    from xml.etree import ElementTree as ET

    findings = []
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()

            if "docProps/core.xml" in names:
                root = ET.fromstring(z.read("docProps/core.xml"))
                for tag in ("creator", "lastModifiedBy", "title", "subject",
                            "description", "keywords"):
                    for node in root.iter():
                        if node.tag.endswith("}" + tag) and node.text and node.text.strip():
                            findings.append("%s: %s" % (tag, node.text.strip()))

            if "docProps/custom.xml" in names:
                custom = z.read("docProps/custom.xml").decode("utf-8", "ignore")
                if "userId" in custom or "KSOTemplate" in custom:
                    findings.append("custom.xml 含设备标识（WPS userId 等）")

            if any("thumbnail" in n for n in names):
                findings.append("含缩略图 thumbnail（可能是原文档的渲染图）")
    except Exception as exc:  # 元数据读不出来不该阻断主流程
        findings.append("元数据读取失败: %r" % (exc,))

    return findings


def check_style_links(path):
    """检查段落样式引用是否都能在 styles.xml 里找到定义。

    这是个特别隐蔽的故障：<w:pStyle w:val="X"> 里的 X 必须和
    styles.xml 里的某个 w:styleId 完全一致。一旦对不上，Word 不会报错，
    而是静默回退到默认样式 —— 文件正常打开、字数页数都对、渲染出来
    看着也像份简历，但配色、字号、标题下划线全部丢失，变成一张白板。

    出错方式通常是：把数字 ID（模板用 164-168）误当成不规范的写法，
    "顺手改成"样式名（Resume Bullet → bullet）。
    """
    if not path.lower().endswith((".docx", ".doc")):
        return []

    import zipfile

    try:
        with zipfile.ZipFile(path) as z:
            doc = z.read("word/document.xml").decode("utf-8", "ignore")
            styles = z.read("word/styles.xml").decode("utf-8", "ignore")
    except Exception:
        return []

    used = set(re.findall(r'<w:pStyle w:val="([^"]+)"', doc))
    defined = set(re.findall(r'<w:style [^>]*w:styleId="([^"]+)"', styles))
    return sorted(used - defined)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    path = os.path.abspath(sys.argv[1])
    if not os.path.exists(path):
        sys.exit("文件不存在: %s" % path)

    spec = json.loads(sys.argv[2]) if len(sys.argv) > 2 else {}
    text = load_text(path)
    # 去掉所有空白再匹配 —— PDF 提取会在任意位置断行，
    # 直接匹配长关键词会大量误报。
    flat = re.sub(r"\s+", "", text)

    ok = True

    # 模板占位符：无论用户传不传 spec 都要检查
    leftover = [p for p in TEMPLATE_PLACEHOLDERS if re.sub(r"\s+", "", p) in flat]
    if leftover:
        print("模板占位符（必须先换掉，HR 看到直接淘汰）:")
        for p in leftover:
            print("  残留! ", p)
        ok = False
    else:
        print("模板占位符:  clean")

    # 文档属性：HR 右键「属性」就能看到，正文里查不出来。
    # 这里只提示不改判定 —— 作者名是本人是正常的，
    # 但如果是从别人模板起的稿，会带上原作者的痕迹。
    meta = check_metadata(path)
    if meta:
        print("文档属性（打开文件属性可见，投递前确认是否需要清）:")
        for m in meta:
            print("  注意! ", m)
    else:
        print("文档属性:    clean")

    # 样式断链：最隐蔽的排版故障，必须查
    broken = check_style_links(path)
    if broken:
        print("样式引用断链（Word 会静默回退默认样式，配色字号全丢）:")
        for b in broken:
            print("  断链! ", b)
        ok = False
    else:
        print("样式引用:    clean")

    for kw in spec.get("must", []):
        hit = re.sub(r"\s+", "", kw) in flat
        print(("OK    " if hit else "缺失! "), kw)
        ok = ok and hit

    for kw in spec.get("must_not", []):
        hit = re.sub(r"\s+", "", kw) in flat
        print(("残留! " if hit else "clean "), kw)
        ok = ok and not hit

    print()
    print("结果:", "全部通过" if ok else "有问题，需要修正")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
