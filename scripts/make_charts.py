#!/usr/bin/env python3
# Genera los 5 visuales de la auditoria con el sistema visual VelogigX (Montserrat, sin riel).
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib.patches import FancyBboxPatch, Wedge, Circle
import numpy as np
from PIL import Image, ImageDraw, ImageFont
import os

import os as _os
_HERE = _os.path.dirname(_os.path.abspath(__file__))
FD = _os.environ.get("VLGX_FONTS", _os.path.join(_HERE, "..", "assets", "fonts"))
for f in ["Regular","Medium","SemiBold","Bold","ExtraBold"]:
    fm.fontManager.addfont(f"{FD}/Montserrat-{f}.ttf")
def F(weight, size):
    name = "Montserrat" if weight=="Regular" else f"Montserrat {weight}"
    return fm.FontProperties(family=name, size=size)

# Paleta VelogigX
ACCENT="#1a5c2e"; ACCENT_L="#effced"; GREEN="#2E7D32"; GREEN_L="#E8F5E9"
RED="#C62828"; RED_L="#fff1f1"; ORANGE="#E65100"; ORANGE_L="#FFF3E0"
GREEN="#2E7D32"; GREEN_L="#eaf6ec"
SLATE="#5b6ب".replace("ب","b")+"82" if False else "#5b6b82"
GRAY="#999999"; GRAY_D="#666666"; NAVY="#2b3a52"; TRACK="#ededed"
OUT="/home/claude/charts"
os.makedirs(OUT, exist_ok=True)

# ---------- Optimizacion de PNGs (v5): charts livianos SIN perdida visible ----------
# Reescala cada visual a 2x su ancho de display y lo pasa a paleta de 256 colores
# (son graficos de color plano -> el ojo no nota diferencia). Reduce el peso del .docx
# ~50% (p.ej. donut 156KB -> 50KB; total imagenes 308KB -> ~132KB). Corre automaticamente
# al terminar el script (atexit), despues de generar todos los visuales. NO toca logo.png.
import atexit, glob as _glob
def _optimize_charts(maxw=1280):
    for _p in _glob.glob(f"{OUT}/*.png"):
        if os.path.basename(_p) == "logo.png":
            continue
        try:
            _im = Image.open(_p).convert("RGBA")
            _bg = Image.new("RGBA", _im.size, (255,255,255,255))   # aplanar sobre blanco (paginas del doc son blancas)
            _im = Image.alpha_composite(_bg, _im).convert("RGB")
            if _im.size[0] > maxw:
                _h = round(_im.size[1] * maxw / _im.size[0])
                _im = _im.resize((maxw, _h), Image.LANCZOS)
            _im = _im.quantize(colors=256, method=Image.MEDIANCUT, dither=Image.NONE)
            _im.save(_p, optimize=True)
        except Exception as _e:
            print(f"[optimize] omitido {_p}: {_e}")
atexit.register(_optimize_charts)

def rrect(ax, x, y, w, h, color, r=0.06, ec="none", lw=0):
    p = FancyBboxPatch((x,y), w, h, boxstyle=f"round,pad=0,rounding_size={r}",
                       fc=color, ec=ec, lw=lw, mutation_aspect=1)
    ax.add_patch(p); return p

# ---------- 1. DONUT: composicion del trafico ----------
def donut(commercial_pct, center_label, title, body_lines, fname):
    fig = plt.figure(figsize=(13,5), dpi=200)
    fig.patch.set_alpha(0)
    # donut left
    axd = fig.add_axes([0.02,0.05,0.34,0.9]); axd.set_aspect("equal"); axd.axis("off")
    axd.set_xlim(-1.3,1.3); axd.set_ylim(-1.3,1.3)
    R, width = 1.05, 0.30
    # track
    axd.add_patch(Wedge((0,0), R, 0, 360, width=width, fc=TRACK, ec="none"))
    # commercial slice (desde arriba, sentido horario)
    sweep = max(commercial_pct/100*360, 4)
    axd.add_patch(Wedge((0,0), R, 90-sweep, 90, width=width, fc=ACCENT, ec="none"))
    axd.text(0,0.12, center_label, ha="center", va="center", color=ACCENT, fontproperties=F("ExtraBold",46))
    axd.text(0,-0.34, "es comercial", ha="center", va="center", color=GRAY, fontproperties=F("Medium",17))
    # right text
    axt = fig.add_axes([0.40,0.05,0.58,0.9]); axt.axis("off"); axt.set_xlim(0,1); axt.set_ylim(0,1)
    axt.text(0,0.92, title, ha="left", va="top", color=ACCENT, fontproperties=F("Bold",25))
    y=0.66
    for ln in body_lines:
        axt.text(0,y, ln, ha="left", va="top", color=GRAY_D, fontproperties=F("Regular",16)); y-=0.115
    # legend
    ly=0.20
    axt.plot([0.018],[ly], marker="o", markersize=13, color=TRACK, transform=axt.transAxes, clip_on=False)
    axt.text(0.055,ly, "Informativo / marca", ha="left", va="center", color=GRAY_D, fontproperties=F("Medium",15))
    axt.plot([0.018],[ly-0.13], marker="o", markersize=13, color=ACCENT, transform=axt.transAxes, clip_on=False)
    axt.text(0.055,ly-0.13, "Comercial del negocio", ha="left", va="center", color=GRAY_D, fontproperties=F("Medium",15))
    fig.savefig(f"{OUT}/{fname}", transparent=True, bbox_inches="tight", pad_inches=0.15); plt.close(fig)

# ---------- 2. KPI CARDS (PIL) ----------
def kpi_cards(cards, fname):
    W,H = 1600, 360; pad=24; gap=24
    img = Image.new("RGBA",(W,H),(255,255,255,0)); d=ImageDraw.Draw(img)
    fnum=ImageFont.truetype(f"{FD}/Montserrat-Bold.ttf",80)
    flab=ImageFont.truetype(f"{FD}/Montserrat-Medium.ttf",30)
    n=len(cards); cw=(W-2*pad-(n-1)*gap)//n
    def hx(c): return tuple(int(c[i:i+2],16) for i in (1,3,5))
    for i,(num,lab,good) in enumerate(cards):
        x0=pad+i*(cw+gap); x1=x0+cw
        bg = hx(GREEN_L) if good else hx(RED_L)
        fg = hx(ACCENT) if good else hx(RED)
        d.rounded_rectangle([x0,40,x1,H-20], radius=28, fill=bg+(255,))
        bb=d.textbbox((0,0),num,font=fnum); tw=bb[2]-bb[0]
        d.text((x0+(cw-tw)//2-bb[0], 132), num, font=fnum, fill=fg+(255,))
        # label (puede ser 2 lineas)
        words=lab.split(); lines=[lab]
        if d.textlength(lab,font=flab) > cw-30:
            mid=len(words)//2; lines=[" ".join(words[:mid])," ".join(words[mid:])]
        ly=250
        for ln in lines:
            lw=d.textlength(ln,font=flab); d.text((x0+(cw-lw)//2, ly), ln, font=flab, fill=hx(GRAY_D)+(255,)); ly+=38
    img.save(f"{OUT}/{fname}")

# ---------- 3. BARRAS HORIZONTALES (competidores) ----------
def hbars(items, fname, bar_color=NAVY, val_color="#5b6b82"):
    vals=[i[1] for i in items]; n=len(items)
    fig,ax=plt.subplots(figsize=(13,n*1.02+0.4),dpi=200); fig.patch.set_alpha(0); ax.axis("off")
    maxv=max(vals); ax.set_xlim(0,maxv*1.16); ax.set_ylim(-0.55,(n-1)+0.78)
    bh=0.30
    for i,(lab,v) in enumerate(items):
        y=n-1-i
        ax.text(0, y+0.30, lab, ha="left", va="bottom", color=NAVY, fontproperties=F("Bold",18))
        rrect(ax, 0, y-bh/2, v, bh, bar_color, r=bh/2)
        ax.text(v+maxv*0.012, y, f"{v:,}", ha="left", va="center", color=val_color, fontproperties=F("Bold",18))
    fig.savefig(f"{OUT}/{fname}", transparent=True, bbox_inches="tight", pad_inches=0.12); plt.close(fig)

# ---------- 4. COMPARATIVA 2 BARRAS (eficiencia / valor) ----------
def compare2(rows, fname, prefix="$"):
    # rows: [(label, value, color)]  primer = cliente (gris), segundo = referente (verde)
    fig,ax=plt.subplots(figsize=(13,3.2),dpi=200); fig.patch.set_alpha(0); ax.axis("off")
    maxv=max(r[1] for r in rows); ax.set_xlim(0,maxv*1.16); ax.set_ylim(-0.3,len(rows)-0.2)
    bh=0.40
    for i,(lab,v,col) in enumerate(rows):
        y=len(rows)-1-i
        ax.text(0, y+bh*0.75, lab, ha="left", va="bottom", color=NAVY if col!=GRAY else GRAY_D, fontproperties=F("SemiBold",18))
        w=max(v/maxv*maxv, maxv*0.012)
        rrect(ax, 0, y-bh/2, max(v, maxv*0.006), bh, col, r=bh/2)
        ax.text(v+maxv*0.015, y, f"{prefix}{v:,.0f}", ha="left", va="center", color=col, fontproperties=F("ExtraBold",22))
    fig.savefig(f"{OUT}/{fname}", transparent=True, bbox_inches="tight", pad_inches=0.1); plt.close(fig)

# ---------- 5. CORE WEB VITALS (barras con meta) ----------
def cwv(metrics, fname):
    # metrics: [(label, value_s, target_s, color)]
    fig,ax=plt.subplots(figsize=(13,4.2),dpi=200); fig.patch.set_alpha(0); ax.axis("off")
    maxv=max(m[1] for m in metrics)*1.12; n=len(metrics)
    ax.set_xlim(-maxv*0.16, maxv*1.05); ax.set_ylim(-0.5,n-0.3); bh=0.42
    for i,(lab,val,tgt,col) in enumerate(metrics):
        y=n-1-i
        ax.text(-maxv*0.02, y, lab, ha="right", va="center", color="#333333", fontproperties=F("SemiBold",20))
        rrect(ax, 0, y-bh/2, val, bh, col, r=bh/2)
        ax.text(val+maxv*0.012, y, f"{val:.1f}s", ha="left", va="center", color=col, fontproperties=F("ExtraBold",21))
        # marca meta (tick gris)
        ax.plot([tgt,tgt],[y-bh*0.85,y+bh*0.85], color="#7a7a7a", lw=3, solid_capstyle="round")
    fig.savefig(f"{OUT}/{fname}", transparent=True, bbox_inches="tight", pad_inches=0.12); plt.close(fig)

# ---------- 6. VISIBILIDAD COMERCIAL (barras agrupadas Top5/Top10) ----------
def visibility(groups, fname, total_kw=10):
    n=len(groups)
    fig,ax=plt.subplots(figsize=(13,4.6),dpi=200); fig.patch.set_alpha(0); ax.axis("off")
    ax.set_xlim(-0.6,n-0.4); ax.set_ylim(0,total_kw*1.18)
    bw=0.30
    for i,(sch,t5,t10) in enumerate(groups):
        rrect(ax, i-bw-0.02, 0, bw, max(t5,0.05), ACCENT, r=0.0)
        rrect(ax, i+0.02,     0, bw, max(t10,0.05), "#b9d8c1", r=0.0)
        ax.text(i-bw/2-0.02, t5+0.18, str(t5), ha="center", va="bottom", color=ACCENT, fontproperties=F("ExtraBold",19))
        ax.text(i+bw/2+0.02, t10+0.18, str(t10), ha="center", va="bottom", color="#5e8c6a", fontproperties=F("ExtraBold",19))
        col = ACCENT if i==0 else NAVY
        ax.text(i, -total_kw*0.06, sch, ha="center", va="top", color=col, fontproperties=F("Bold",18))
    ax.add_patch(FancyBboxPatch((-0.55,total_kw*1.06),0.16,total_kw*0.05,boxstyle="round,pad=0,rounding_size=0",fc=ACCENT,ec="none"))
    ax.text(-0.34,total_kw*1.085,"en Top 5", ha="left", va="center", color=GRAY_D, fontproperties=F("Medium",15))
    ax.add_patch(FancyBboxPatch((1.0,total_kw*1.06),0.16,total_kw*0.05,boxstyle="round,pad=0,rounding_size=0",fc="#b9d8c1",ec="none"))
    ax.text(1.21,total_kw*1.085,"en Top 10", ha="left", va="center", color=GRAY_D, fontproperties=F("Medium",15))
    fig.savefig(f"{OUT}/{fname}", transparent=True, bbox_inches="tight", pad_inches=0.15); plt.close(fig)

# ================= DATOS ESAN =================
donut(15, "~15%", "De dónde viene el tráfico orgánico",
      ["De las ~173,844 visitas/mes, la mayor parte",
       "llega por artículos informativos de Conexión",
       "ESAN (tipo de cambio, recibos por honorarios,",
       "marcas) \u2014 no por quien evalúa una maestría."],
      "donut.png")

kpi_cards([("173 844","visitas/mes orgánicas",True),
           ("6.0 s","tiempo de carga (LCP)",False),
           ("24 MB","peso de la home",False),
           ("~37 000","búsquedas/mes del mercado",True)],
          "kpis.png")

visibility([("ESAN",5,7),("UPC",4,7),("Centrum",4,4),("Pacífico",1,1)], "visibilidad.png", total_kw=10)

cwv([("LCP",6.0,2.5,RED),("FCP",2.6,1.8,ORANGE),
     ("Speed Index",4.7,3.4,ORANGE),("TTI",6.6,3.8,RED)], "cwv.png")

print("OK charts:", os.listdir(OUT))
