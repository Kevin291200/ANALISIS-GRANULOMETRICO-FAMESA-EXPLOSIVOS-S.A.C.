"""Gráficas del informe, con la paleta FAMESA: negro, plomo y rojo."""
import io
from bisect import bisect_right

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

import core
from core import IN_MM, PCTS, SIZES

NEGRO, PLOMO, PLOMO_CL, ROJO, ROJO_OSC = "#111111", "#6F7378", "#C9CCD0", "#E10613", "#8E0A12"
PALETA = [ROJO, NEGRO, PLOMO, ROJO_OSC, "#F0827D", "#3A3D42", "#A9ADB2", "#C45A5F"]
ESTILOS = ["-", "--", "-.", ":"]
plt.rcParams.update({"axes.edgecolor": "#444", "axes.labelcolor": "#222", "xtick.color": "#333",
                     "ytick.color": "#333", "font.size": 9})


def _png(fig):
    b = io.BytesIO()
    fig.savefig(b, format="png", dpi=150, bbox_inches="tight", pad_inches=0.3)
    plt.close(fig)
    return b.getvalue()


SERIE_125 = [m * 10 ** e for e in range(-2, 3) for m in (1, 2, 5)]      # 0.01 ... 500
SERIE_MM = [1, 2, 5, 10, 20, 50, 100, 200, 500, 1000, 2000]


def _eje_tamano(ax, lo, hi):
    """Eje X logarítmico con marcas uniformes 1-2-5 (in) abajo y 1-2-5 (mm) arriba, sin etiquetas pisadas."""
    from matplotlib.ticker import FixedLocator, FuncFormatter, LogLocator, NullFormatter, NullLocator
    lo_s = max([v for v in SERIE_125 if v <= lo] or [SERIE_125[0]])
    hi_s = min([v for v in SERIE_125 if v >= hi] or [SERIE_125[-1]])
    if hi_s <= lo_s:
        hi_s = lo_s * 10
    ax.set_xscale("log")
    ax.set_xlim(lo_s, hi_s)
    ax.xaxis.set_major_locator(FixedLocator([v for v in SERIE_125 if lo_s <= v <= hi_s]))
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:g}"))
    ax.xaxis.set_minor_locator(NullLocator())
    ax.grid(True, which="major", alpha=.30)
    ax.set_xlabel("Tamaño (in)", labelpad=6)
    sec = ax.secondary_xaxis("top", functions=(lambda x: x * IN_MM, lambda x: x / IN_MM))
    tk = [v for v in SERIE_MM if lo_s * IN_MM <= v <= hi_s * IN_MM]
    if len(tk) >= 2:
        sec.xaxis.set_major_locator(FixedLocator(tk))
        sec.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:g}"))
        sec.xaxis.set_minor_locator(NullLocator())
    sec.set_xlabel("Tamaño (mm)", labelpad=6)
    return lo_s, hi_s


def _leyenda_abajo(ax, ncol=2):
    """Leyenda fuera del área de la gráfica (no tapa curvas ni puntos)."""
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.17), ncol=ncol, frameon=False, fontsize=8.5,
              handlelength=2.6, columnspacing=1.6)


def _estilo(i):
    return PALETA[i % len(PALETA)], ESTILOS[(i // len(PALETA)) % len(ESTILOS)]


def fig_curva_tajo(t, mina=None):
    res, meds, obj = t["res"], t["meds"], t["obj"]
    pts = [(res[f"P{p}"], p) for p in PCTS if res[f"P{p}"] is not None]
    if not pts:
        return None
    fig, ax = plt.subplots(figsize=(8, 5.0))
    for i, m in enumerate(meds):
        xy = [(m["pcts"].get(str(p)), p) for p in PCTS if m["pcts"].get(str(p)) is not None]
        if xy:
            ax.plot(*zip(*xy), color=PLOMO_CL, lw=1.2, label="Muestras individuales" if i == 0 else None)
    xs, ys = zip(*pts)
    ax.plot(xs, ys, color=NEGRO, lw=2.8, marker="o", ms=3.5, label="Curva promedio del tajo")
    if obj:
        ax.axvline(obj, color=ROJO, ls="--", lw=1.6, label=f"P80 objetivo voladura = {obj:.2f} in")
    if mina:
        ax.axvline(mina, color=NEGRO, ls="-.", lw=1.6, label=f"P80 objetivo mina = {mina:.2f} in")
    todos = [x for m in meds for x in m["pcts"].values() if x] + list(xs) + [v for v in (obj, mina) if v]
    lo_s, hi_s = _eje_tamano(ax, min(todos) * .85, max(todos) * 1.15)
    ax.set_ylim(0, 100)
    ax.set_yticks(range(0, 101, 20))
    ax.set_ylabel("% pasante acumulado")
    for p in (20, 50, 80):
        x = res[f"P{p}"]
        if x is None:
            continue
        ax.vlines(x, 0, p, ls=":", color=ROJO, lw=1)
        ax.plot(x, p, "o", color=ROJO, ms=7, zorder=5)
        ax.annotate(f"P{p} = {x:.2f} in ({x * IN_MM:.0f} mm)", (x, p), textcoords="offset points",
                    xytext=(9, -15), ha="left", fontsize=8.5, fontweight="bold", color=ROJO, zorder=6,
                    bbox=dict(boxstyle="round,pad=0.25", fc="white", ec=ROJO, lw=.6, alpha=.95))
    _leyenda_abajo(ax, 2)
    ax.set_title(f"Curva granulométrica · {t['label']}", fontweight="bold", color=NEGRO, pad=40)
    return _png(fig)


def fig_hist(t):
    res = t["res"]
    vals = [res["sizes"][f"{s:g}"] for s in SIZES]
    if any(v is None for v in vals):
        return None
    ret = [vals[0]] + [max(vals[i] - vals[i - 1], 0) for i in range(1, len(vals))] + [max(100 - vals[-1], 0)]
    labels = [f"<{SIZES[0]:g}"] + [f"{a:g}–{b:g}" for a, b in zip(SIZES[:-1], SIZES[1:])] + [f">{SIZES[-1]:g}"]
    idx = bisect_right([0] + SIZES, res["P80"]) - 1 if res["P80"] else -1
    fig, ax = plt.subplots(figsize=(8, 3.6))
    bars = ax.bar(labels, ret, color=[ROJO if i == idx else PLOMO for i in range(len(ret))], edgecolor="white")
    for b, v in zip(bars, ret):
        if v >= 1:
            ax.text(b.get_x() + b.get_width() / 2, v + .6, f"{v:.1f}", ha="center", fontsize=7.5)
    ax.set_xlabel("Rango de tamaño (in)  ·  en rojo, el rango donde cae el P80")
    ax.set_ylabel("% en peso retenido")
    ax.grid(axis="y", alpha=.25)
    plt.setp(ax.get_xticklabels(), rotation=45, ha="right", fontsize=8)
    ax.set_title(f"Distribución por rangos de tamaño · {t['label']}", fontweight="bold", color=NEGRO)
    return _png(fig)


def fig_forma_cov(t):
    meds = t["meds"]
    n = len(meds)
    etiquetas = [f"M{i + 1}" for i in range(n)]
    esf = [m["sphericity"] for m in meds]
    cov = [m["coverage"] for m in meds]
    fig, axs = plt.subplots(1, 2, figsize=(8, 3.3))
    for ax, vals, titulo, refs in ((axs[0], esf, "Esfericidad por muestra (%)", (80, 60)),
                                   (axs[1], cov, "Coverage por muestra (%)", (85, 70))):
        if all(v is None for v in vals):
            ax.text(.5, .5, "Sin datos en el archivo", ha="center", va="center", transform=ax.transAxes)
            ax.set_axis_off()
            continue
        v2 = [v or 0 for v in vals]
        bars = ax.bar(etiquetas, v2, color=ROJO if titulo.startswith("Esf") else NEGRO, width=.6)
        for b, v in zip(bars, vals):
            if v is not None:
                ax.text(b.get_x() + b.get_width() / 2, v + 1.5, f"{v:.0f}", ha="center", fontsize=8,
                        fontweight="bold")
        ax.axhline(refs[0], color=PLOMO, ls="--", lw=1)
        ax.axhline(refs[1], color=PLOMO, ls=":", lw=1)
        ax.text(n - .45, refs[0] + 1, f"{refs[0]} %", fontsize=7, color=PLOMO, ha="right")
        ax.text(n - .45, refs[1] + 1, f"{refs[1]} %", fontsize=7, color=PLOMO, ha="right")
        ax.set_ylim(0, 112)
        ax.set_title(titulo, fontweight="bold", fontsize=9.5)
        ax.grid(axis="y", alpha=.25)
    fig.suptitle(f"Esfericidad y coverage · {t['label']}", fontweight="bold", color=NEGRO, y=1.03)
    return _png(fig)


def fig_curvas_generales(tajos, mina=None):
    fig, ax = plt.subplots(figsize=(8, 5.0))
    todos = [mina] if mina else []
    for i, t in enumerate(tajos):
        pts = [(t["res"][f"P{p}"], p) for p in PCTS if t["res"][f"P{p}"] is not None]
        if not pts:
            continue
        c, ls = _estilo(i)
        ax.plot(*zip(*pts), color=c, ls=ls, lw=2, label=t["label"])
        todos += [x for x, _ in pts]
    if mina:
        ax.axvline(mina, color=NEGRO, ls="-.", lw=1.4, label=f"P80 objetivo mina = {mina:.2f} in")
    if not todos:
        plt.close(fig)
        return None
    _eje_tamano(ax, min(todos) * .85, max(todos) * 1.15)
    ax.axhline(80, color=PLOMO, ls=":", lw=1)
    ax.text(ax.get_xlim()[1], 81.5, "P80 ", fontsize=8.5, color=PLOMO, ha="right", fontweight="bold")
    ax.set_ylim(0, 100)
    ax.set_yticks(range(0, 101, 20))
    ax.set_ylabel("% pasante acumulado")
    _leyenda_abajo(ax, 2 if len(tajos) > 3 else 1)
    ax.set_title("Curvas granulométricas promedio por tajo", fontweight="bold", pad=40)
    return _png(fig)


def fig_ajuste(tajos, tipo):
    """tipo = 'rr' (Rosin-Rammler) o 'sw' (Swebrec): puntos medidos y curva del modelo."""
    fig, ax = plt.subplots(figsize=(8, 5.0))
    series, xs_all = [], []
    for i, t in enumerate(tajos):
        par = t["par"]
        pts = par["pts"]
        if not pts:
            continue
        xs = np.array([x for x, _ in pts])
        lo, hi = xs.min() * .5, xs.max() * 1.3
        xg = np.geomspace(lo, hi, 240)
        if tipo == "rr" and par["n"] and par["xc"]:
            yg, etq = core.curva_rr(xg, par["n"], par["xc"]), f"{t['label']}  (n = {par['n']:.2f})"
        elif tipo == "sw" and par["b"] and par["xmax"] and par["x50"] and par["xmax"] > par["x50"]:
            yg, etq = core.curva_sw(xg, par["b"], par["xmax"], par["x50"]), f"{t['label']}  (b = {par['b']:.2f})"
        else:
            continue
        series.append((i, pts, xg, yg, etq))
        xs_all += [xs.min(), xs.max()]
    if not series:
        plt.close(fig)
        return None
    _eje_tamano(ax, min(xs_all) * .6, max(xs_all) * 1.2)
    lo_ax, hi_ax = ax.get_xlim()
    for i, pts, xg, yg, etq in series:
        c, ls = _estilo(i)
        m = (xg >= lo_ax) & (xg <= hi_ax)
        ax.plot(xg[m], yg[m], color=c, ls=ls, lw=1.9, label=etq, zorder=3)
        ax.scatter([x for x, _ in pts], [p for _, p in pts], color=c, s=26, zorder=4, edgecolor="white",
                   linewidth=.6)
    ax.set_ylim(0, 100)
    ax.set_yticks(range(0, 101, 20))
    ax.set_ylabel("% pasante acumulado")
    _leyenda_abajo(ax, 2 if len(series) > 3 else 1)
    nombre = "Rosin-Rammler" if tipo == "rr" else "Swebrec"
    ax.set_title(f"Ajuste de la distribución {nombre}", fontweight="bold", pad=40)
    ax.text(.99, .03, "Puntos = datos medidos  ·  Línea = modelo", transform=ax.transAxes, ha="right",
            fontsize=8, color=PLOMO, bbox=dict(boxstyle="round,pad=0.25", fc="white", ec="none", alpha=.9))
    return _png(fig)


def fig_p80_tajos(tajos, mina=None):
    datos = [t for t in tajos if t["res"]["P80"] is not None]
    if not datos:
        return None
    fig, ax = plt.subplots(figsize=(8, 3.9))
    etq = [t["label"].replace(" · ", "\n") for t in datos]
    cols = []
    for t in datos:
        meta = t["meta"]
        cols.append(PLOMO if not meta else (NEGRO if t["res"]["P80"] <= meta else ROJO))
    ax.bar(etq, [t["res"]["P80"] for t in datos], color=cols, width=.6)
    for x, t in enumerate(datos):
        ax.text(x, t["res"]["P80"] + .05, f"{t['res']['P80']:.2f}", ha="center", fontsize=8, fontweight="bold")
        if t["obj"]:
            ax.hlines(t["obj"], x - .38, x + .38, color=ROJO, lw=2.4)
    if mina:
        ax.axhline(mina, color=NEGRO, ls="-.", lw=1.4)
    from matplotlib.lines import Line2D
    from matplotlib.patches import Patch
    h = [Patch(color=NEGRO, label="Cumple"), Patch(color=ROJO, label="No cumple"),
         Patch(color=PLOMO, label="Sin objetivo")]
    if any(t["obj"] for t in datos):
        h.append(Line2D([0], [0], color=ROJO, lw=2.4, label="Objetivo de la voladura"))
    if mina:
        h.append(Line2D([0], [0], color=NEGRO, ls="-.", label=f"Objetivo de mina ({mina:.2f} in)"))
    ax.legend(handles=h, fontsize=8, ncol=3, loc="upper center", bbox_to_anchor=(0.5, -0.22 if len(datos) > 6 else -0.16),
              frameon=False)
    ax.set_ylabel("P80 (in)")
    ax.grid(axis="y", alpha=.25)
    plt.setp(ax.get_xticklabels(), fontsize=7.5, rotation=(0 if len(datos) <= 6 else 35),
             ha=("center" if len(datos) <= 6 else "right"))
    ax.set_title("P80 por tajo frente al objetivo", fontweight="bold", pad=12)
    return _png(fig)
