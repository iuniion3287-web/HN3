#!/usr/bin/env python3
"""PDF를 페이지별 PNG와 한 장짜리 미리보기(contact sheet)로 렌더링한다. (PyMuPDF 필요)

사용법:
    python render_preview.py 보고서.pdf -o preview/ [--dpi 110]
"""
import argparse
import os


def render(pdf, out_dir, dpi=110, cols=4):
    import fitz  # PyMuPDF
    os.makedirs(out_dir, exist_ok=True)
    doc = fitz.open(pdf)
    paths = []
    for i, page in enumerate(doc, 1):
        pix = page.get_pixmap(dpi=dpi)
        p = os.path.join(out_dir, f"page-{i:02d}.png")
        pix.save(p)
        paths.append(p)
    # contact sheet
    thumb_w = 300
    first = fitz.Pixmap(paths[0])
    thumb_h = int(first.height * thumb_w / first.width)
    rows = (len(paths) + cols - 1) // cols
    gap = 12
    sheet = fitz.open()
    W, H = cols * thumb_w + (cols + 1) * gap, rows * thumb_h + (rows + 1) * gap
    pg = sheet.new_page(width=W, height=H)
    pg.draw_rect(pg.rect, color=None, fill=(0.93, 0.94, 0.96))
    for i, p in enumerate(paths):
        r, c = divmod(i, cols)
        x, y = gap + c * (thumb_w + gap), gap + r * (thumb_h + gap)
        rect = fitz.Rect(x, y, x + thumb_w, y + thumb_h)
        pg.insert_image(rect, filename=p)
        pg.draw_rect(rect, color=(0.75, 0.78, 0.82), width=0.8)
    sheet_path = os.path.join(out_dir, "overview.png")
    pg.get_pixmap(dpi=96).save(sheet_path)
    print(f"{len(paths)} pages -> {out_dir} (overview: {sheet_path})")
    return paths, sheet_path


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdf")
    ap.add_argument("-o", "--out", default="preview")
    ap.add_argument("--dpi", type=int, default=110)
    a = ap.parse_args()
    render(a.pdf, a.out, a.dpi)
