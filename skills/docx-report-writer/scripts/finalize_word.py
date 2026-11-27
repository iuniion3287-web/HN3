#!/usr/bin/env python3
"""Microsoft Word(Windows)로 DOCX를 열어 목차·페이지 번호 필드를 갱신하고 저장한 뒤 PDF로 내보낸다.

사용법:
    python finalize_word.py 보고서.docx [--pdf 보고서.pdf]

- Word가 없으면 건너뛴다(종료 코드 3). 이 경우 DOCX는 Word에서 처음 열 때
  "필드 업데이트" 안내가 나오고, [예]를 누르면 목차가 채워진다.
- 성공하면 settings.xml의 updateFields 플래그를 제거해 사용자가 열 때 확인 창이 뜨지 않게 한다.
"""
import argparse
import os
import shutil
import sys
import tempfile
import zipfile


def strip_update_fields(path):
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/settings.xml":
                import re
                data = re.sub(rb"<w:updateFields[^>]*/>", b"", data)
            if not item.filename.endswith("/"):
                zout.writestr(item, data)
    os.replace(tmp, path)


def finalize(docx, pdf=None):
    try:
        import win32com.client  # type: ignore
    except ImportError:
        print("[SKIP] pywin32가 없어 Word 자동화를 건너뜁니다.")
        return 3
    docx = os.path.abspath(docx)
    pdf = os.path.abspath(pdf or os.path.splitext(docx)[0] + ".pdf")
    try:
        word = win32com.client.DispatchEx("Word.Application")
    except Exception as e:  # Word 미설치
        print(f"[SKIP] Microsoft Word를 실행할 수 없습니다: {e}")
        return 3
    word.Visible = False
    word.DisplayAlerts = 0
    doc = None
    try:
        doc = word.Documents.Open(docx, ConfirmConversions=False, ReadOnly=False, AddToRecentFiles=False)
        for _ in range(2):  # TOC 갱신 후 쪽 번호가 바뀔 수 있어 두 번 갱신
            for toc in doc.TablesOfContents:
                toc.Update()
            doc.Fields.Update()
            for sec in doc.Sections:
                for hf in list(sec.Headers) + list(sec.Footers):
                    hf.Range.Fields.Update()
        doc.Save()
        doc.ExportAsFixedFormat(pdf, 17)  # 17 = wdExportFormatPDF
        pages = doc.ComputeStatistics(2)  # 2 = wdStatisticPages
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()
    strip_update_fields(docx)
    print(f"[OK] 필드 갱신 및 저장: {docx}")
    print(f"[OK] PDF 내보내기: {pdf} ({pages}쪽)")
    return 0


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx")
    ap.add_argument("--pdf")
    a = ap.parse_args()
    sys.exit(finalize(a.docx, a.pdf))


if __name__ == "__main__":
    main()
