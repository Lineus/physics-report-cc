# -*- coding: utf-8 -*-
"""从讲义/PPT 导出的 PDF 里取图：位图拿原图，矢量图按 bbox 渲染。

依赖：PyMuPDF(fitz) + Pillow。用法见 docx-report-build/SKILL.md Step 2。
"""
import os

import fitz
from PIL import Image


# ---------------------------------------------------------------- 探测
def page_stats(pdf_path):
    """每页的位图数 / 矢量数，用来判断目标图是位图还是矢量。"""
    doc = fitz.open(pdf_path)
    for i, p in enumerate(doc):
        print(i + 1, "imgs", len(p.get_image_info()), "drawings", len(p.get_drawings()))


# ---------------------------------------------------------------- 位图
def extract_embedded(doc, page_no, out_prefix, bg_ratio=0.95, icon_ratio=0.05):
    """取该页内嵌位图原图（干净、无水印）。返回落盘路径列表。

    过滤：整页背景图（覆盖 >= bg_ratio 页宽高）、小图标（两边 < icon_ratio 页尺寸）。
    实测一页 PPT 常有 10+ 张位图，其中只有 1-2 张是正文图。
    """
    p = doc[page_no - 1]
    R = p.rect
    out = []
    for info in p.get_image_info(xrefs=True):
        b = fitz.Rect(info["bbox"])
        if (b.width >= R.width * bg_ratio and b.height >= R.height * bg_ratio):
            continue
        if b.width < R.width * icon_ratio and b.height < R.height * icon_ratio:
            continue
        im = doc.extract_image(info["xref"])
        path = f"{out_prefix}_p{page_no}_x{info['xref']}.{im['ext']}"
        with open(path, "wb") as f:
            f.write(im["image"])
        out.append(path)
        print(path, im["width"], im["height"], "bbox", tuple(round(v) for v in b))
    return out


# ---------------------------------------------------------------- 矢量
def union_bbox(page, y0, y1):
    """y0<图<y1 区间内矢量线段与位图的并集 bbox（向外留 6pt 边距）。

    区间用图注文字定位：page.search_for("图4-1-8") 拿到 y，图在它上方。
    """
    xs, ys, xe, ye = [], [], [], []
    boxes = [d["rect"] for d in page.get_drawings()]
    boxes += [fitz.Rect(i["bbox"]) for i in page.get_image_info()]
    for b in boxes:
        if y0 < b.y0 and b.y1 < y1 and b.width > 1 and b.height > 1:
            xs.append(b.x0); ys.append(b.y0); xe.append(b.x1); ye.append(b.y1)
    if not xs:
        raise SystemExit("union_bbox: 区间内没找到图形，检查 y0/y1")
    return fitz.Rect(min(xs) - 6, min(ys) - 6, max(xe) + 6, max(ye) + 6)


def render_clip(doc, page_no, rect, out, dpi=300):
    """矢量图按 bbox 渲染成位图。"""
    doc[page_no - 1].get_pixmap(dpi=dpi, clip=rect).save(out)
    print(out, "clip", tuple(round(v) for v in rect))
    return out


# ---------------------------------------------------------------- 后处理
def to_white(img):
    """带透明通道的图铺白底成 RGB（PPT 内嵌 jpeg 多带 smask）。"""
    if img.mode in ("RGBA", "LA"):
        bg = Image.new("RGB", img.size, "white")
        bg.paste(img, mask=img.split()[-1])
        return bg
    return img.convert("RGB")


def trim_caption(img, frac=0.05):
    """粗略裁掉底部图注条（渲染时图注会落进 bbox）；有 caption 坐标就用下面的精确版。"""
    w, h = img.size
    return img.crop((0, 0, w, int(h * (1 - frac))))


def crop_by_caption(img_path, clip, caption_top_y, out=None, margin_pt=5):
    """按图注位置精确裁掉源图自带的“图N”（会和报告自编图号打架）。

    clip = 渲染这张图时用的 fitz.Rect，caption_top_y = page.search_for("图N") 拿到的 y0。
    裁完必须 view_image 复检，别把正图裁掉。2026-09-28 实测：8 张图里 6 张带内嵌图号。
    """
    im = Image.open(img_path)
    w, h = im.size
    cut = round((caption_top_y - clip.y0 - margin_pt) / clip.height * h)
    cut = max(1, min(h - 1, cut))
    out = out or img_path
    im.crop((0, 0, w, cut)).save(out)
    print(out, (w, h), "->", (w, cut))
    return out


def compose(paths, target_h=520, gap=40, margin=14):
    """多张图统一缩到同高、并排拼一张白底图（电路图 + 波形图常用）。"""
    ims = []
    for path in paths:
        im = to_white(Image.open(path))
        w, h = im.size
        ims.append(im.resize((max(1, round(w * target_h / h)), target_h), Image.LANCZOS))
    canvas = Image.new("RGB", (sum(i.size[0] for i in ims) + gap + 2 * margin,
                              target_h + 2 * margin), "white")
    x = margin
    for im in ims:
        canvas.paste(im, (x, margin))
        x += im.size[0] + gap
    return canvas


def save(img, path):
    img.save(path)
    print(path, img.size)
    return path


if __name__ == "__main__":
    import sys
    page_stats(sys.argv[1])
