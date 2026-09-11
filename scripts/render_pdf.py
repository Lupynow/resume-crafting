"""把 PDF 渲染成 PNG，用来目视检查排版。

排版问题（文字溢出、行距挤压、内容断裂）从文本提取里看不出来，
必须渲染成图看一眼。改完简历后这一步不能省。

用法:
    python render_pdf.py <file.pdf> [输出目录] [dpi]

环境要求:
    pip install pymupdf
"""
import os
import sys


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    src = os.path.abspath(sys.argv[1])
    if not os.path.exists(src):
        sys.exit("文件不存在: %s" % src)

    outdir = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.dirname(src)
    dpi = int(sys.argv[3]) if len(sys.argv) > 3 else 180
    os.makedirs(outdir, exist_ok=True)

    import fitz  # PyMuPDF

    doc = fitz.open(src)
    print("页数:", len(doc))

    for i, page in enumerate(doc, 1):
        out = os.path.join(outdir, "page%d.png" % i)
        page.get_pixmap(dpi=dpi).save(out)
        print("已保存:", out)

    # 报告首页底部剩余空间 —— 接近 0 说明随时会溢到第二页，
    # 还剩一大截说明可以适当放松行距或补内容。
    page = doc[0]
    blocks = page.get_text("blocks")
    if blocks:
        bottom = max(b[3] for b in blocks)
        remain = page.rect.height - bottom
        print("首页内容底部: %.1f pt / 页高 %.1f pt" % (bottom, page.rect.height))
        print("底部剩余: %.1f pt (约 %.1f 行)" % (remain, remain / 9.5))


if __name__ == "__main__":
    main()
