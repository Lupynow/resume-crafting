"""统计 docx 页数，可选导出 PDF。

Windows 上走 Word / WPS 的 COM 接口 —— 这是页数最准的办法，
纯 Python 库无法准确还原 Word 的分页结果。

用法:
    python page_count.py <简历.docx> [导出.pdf]

环境要求:
    pip install pywin32
    未装 Word 时会自动回退到 WPS（KWps.Application）。
"""
import os
import sys


def open_document(path):
    """依次尝试 Word / WPS，返回第一个能打开该文档的引擎。

    用 DispatchEx 而不是 Dispatch —— 它强制创建新实例，
    避免命中残留的、已经卡死的 COM 进程。这一点在跑过几轮之后很关键，
    否则会看到 ComputeStatistics 之类的莫名其妙的报错。
    """
    import win32com.client as win32

    last_err = None
    for progid in ("Word.Application", "KWps.Application", "Wps.Application"):
        try:
            app = win32.DispatchEx(progid)
        except Exception as exc:
            last_err = exc
            continue

        for attr, value in (("Visible", False), ("DisplayAlerts", 0)):
            try:
                setattr(app, attr, value)
            except Exception:
                pass  # WPS 不一定支持全部属性，忽略即可

        try:
            doc = app.Documents.Open(path, ReadOnly=True)
            return app, doc, progid
        except Exception as exc:
            last_err = exc
            try:
                app.Quit()
            except Exception:
                pass

    raise RuntimeError("无法打开文档，Word 和 WPS 都失败了: %r" % (last_err,))


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    src = os.path.abspath(sys.argv[1])
    dst = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else None

    if not os.path.exists(src):
        sys.exit("文件不存在: %s" % src)

    app, doc, progid = open_document(src)
    try:
        print("引擎:", progid)
        print("PAGES:", doc.ComputeStatistics(2))  # wdStatisticPages
        print("WORDS:", doc.ComputeStatistics(0))  # wdStatisticWords
        if dst:
            doc.SaveAs(dst, FileFormat=17)  # wdFormatPDF
            print("PDF:", dst)
    finally:
        try:
            doc.Close(False)
        finally:
            try:
                app.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    main()
