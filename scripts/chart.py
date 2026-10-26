#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""reportlab 矢量出图：线 + 逐点误差棒 + 坐标轴 + 图例。
matplotlib 在本机因 numpy ABI 不匹配不可用，统一用这个。

用法：
    from chart import chart
    chart('fig4', [dict(kind='line', pts=theory, color=NAVY, label='Theory'),
                   dict(kind='dot',  pts=pts,    label='Measured')],
          xr=(-6, 6), yr=(0.65, 1.05), xticks=[-6,-4,-2,0,2,4,6],
          yticks=[0.7,0.8,0.9,1.0], ylab='B / mT')
    # pts: line -> [(x, y)];  dot -> [(x, y, u)]，u 为 0/None 时不画误差棒
"""
import fitz
from reportlab.graphics.shapes import Drawing, String, Line, Circle, PolyLine
from reportlab.graphics import renderPDF
from reportlab.lib import colors

NAVY = colors.HexColor('#244A70')
GREY = colors.HexColor('#C9C9C9')
RED = colors.HexColor('#A6362A')
GREEN = colors.HexColor('#4F7A4A')


def chart(fname, series, xr, yr, xticks, yticks, ylab='B / mT',
          xlab='Axial distance x / cm', legend=True, W=470, H=280, dpi=320):
    L, Rr, T, Bm = 62, 430, 226, 62
    dr = Drawing(W, H)
    X = lambda v: L + (v - xr[0]) / (xr[1] - xr[0]) * (Rr - L)
    Y = lambda v: Bm + (v - yr[0]) / (yr[1] - yr[0]) * (T - Bm)

    for t in xticks:
        dr.add(Line(X(t), T, X(t), Bm, strokeColor=GREY, strokeWidth=0.45))
    for t in yticks:
        dr.add(Line(L, Y(t), Rr, Y(t), strokeColor=GREY, strokeWidth=0.45))
    dr.add(Line(L, T, L, Bm, strokeColor=colors.black, strokeWidth=1.1))
    dr.add(Line(L, Bm, Rr, Bm, strokeColor=colors.black, strokeWidth=1.1))

    for t in xticks:
        dr.add(Line(X(t), Bm, X(t), Bm - 4, strokeColor=colors.black, strokeWidth=1.0))
        dr.add(String(X(t), Bm - 16, '%g' % t, fontName='Times-Roman',
                      fontSize=9.5, textAnchor='middle'))
    for t in yticks:
        dr.add(Line(L, Y(t), L - 4, Y(t), strokeColor=colors.black, strokeWidth=1.0))
        lab = ('%.3f' % t).rstrip('0').rstrip('.')
        dr.add(String(L - 7, Y(t) - 3, lab, fontName='Times-Roman',
                      fontSize=9.5, textAnchor='end'))

    for s in series:
        if s['kind'] == 'line':
            dr.add(PolyLine([(X(a), Y(b)) for a, b in s['pts']],
                            strokeColor=s['color'], strokeWidth=s.get('w', 1.5),
                            strokeDashArray=s.get('dash')))
        else:
            for pt in s['pts']:
                xv, bv = pt[0], pt[1]
                uv = pt[2] if len(pt) > 2 else 0
                cx, cy = X(xv), Y(bv)
                if uv:
                    e = abs(Y(bv + uv) - Y(bv))
                    dr.add(Line(cx, cy - e, cx, cy + e, strokeColor=colors.black, strokeWidth=0.7))
                    dr.add(Line(cx - 2.4, cy + e, cx + 2.4, cy + e, strokeColor=colors.black, strokeWidth=0.7))
                    dr.add(Line(cx - 2.4, cy - e, cx + 2.4, cy - e, strokeColor=colors.black, strokeWidth=0.7))
                dr.add(Circle(cx, cy, 2.2, fillColor=colors.black, strokeColor=colors.black))

    dr.add(String((L + Rr) / 2, 30, xlab, fontName='Times-Roman',
                  fontSize=10.5, textAnchor='middle'))
    dr.add(String(10, T + 30, ylab, fontName='Times-Roman', fontSize=10.5))

    if legend:
        x0 = W - 150
        for i, s in enumerate(series):
            yy = T + 26 - i * 13
            if s['kind'] == 'line':
                dr.add(Line(x0, yy, x0 + 22, yy, strokeColor=s['color'],
                            strokeWidth=s.get('w', 1.5), strokeDashArray=s.get('dash')))
            else:
                dr.add(Circle(x0 + 11, yy, 2.2, fillColor=colors.black, strokeColor=colors.black))
            dr.add(String(x0 + 28, yy - 4, s['label'], fontName='Times-Roman', fontSize=9.0))

    renderPDF.drawToFile(dr, fname + '.pdf')
    d = fitz.open(fname + '.pdf')
    d[0].get_pixmap(dpi=dpi).save(fname + '.png')
    d.close()
    print('wrote', fname + '.png')
