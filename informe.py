"""Informe semanal de fragmentación: contenido en bloques (la salida PDF/Word está en formato.py)."""
import hashlib
import io
import json
import os
import re
import shutil
import tempfile
import threading
from datetime import datetime
from xml.sax.saxutils import escape

import numpy as np

import assets
import core
import graficos
from core import IN_MM, PCTS, fdate, fmt

ROJO, NEGRO, PLOMO = "#E10613", "#111111", "#6F7378"

DEF_TITULO = "INFORME SEMANAL\nFRAGMENTACIÓN\nBAJO CONTROL"
DEF_SUBTITULO = "Informe semanal de análisis granulométrico · Taladros largos"
DEF_OBJETIVO = (
    "Evaluar la fragmentación de la roca volada en los tajos de taladros largos del periodo, "
    "mediante análisis granulométrico por procesamiento de imágenes con WipFrag 4, y verificar "
    "el cumplimiento del P80 objetivo, a fin de retroalimentar el diseño de voladura y sostener "
    "el desempeño de las etapas posteriores de carguío, acarreo y conminución.")
DEF_METODOLOGIA = "\n".join([
    "Captura de imágenes del material fragmentado en la pila o frente de carguío, con un objeto de "
    "escala de dimensiones conocidas en cada fotografía.",
    "Procesamiento de cada imagen en WipFrag 4: calibración de escala, delineación de fragmentos y "
    "corrección de contornos cuando es necesario.",
    "Obtención de la curva granulométrica acumulada de cada muestra, de los tamaños característicos "
    "(Xmax, X50, Xc) y de los porcentajes pasantes por cada malla (1.5 in a 20 in).",
    "Cálculo de los percentiles P-05 a P-99 y ajuste de las distribuciones de Rosin-Rammler "
    "(coeficiente de uniformidad n) y de Swebrec (parámetro de forma b).",
    "Consolidación por tajo: promedio de los percentiles de las muestras de cada voladura.",
    "Evaluación del P80 frente al objetivo de mina y al objetivo de cada voladura. Criterio de "
    "cumplimiento: P80 menor o igual al objetivo.",
    "Control de calidad de cada muestra mediante la esfericidad y el coverage de la imagen.",
])


# ------------------------------------------------------------------ caché de gráficos
_LOCK = threading.RLock()   # matplotlib/pyplot no es seguro entre hilos: un gráfico a la vez


def _fig(nombre, fn, objeto, *args):
    """Cada gráfico se dibuja una sola vez por sesión; si los datos del tajo no cambian, se reutiliza."""
    ts = objeto if isinstance(objeto, list) else [objeto]
    clave = (nombre, tuple((t["h"], t["obj"]) for t in ts), args)
    cache = core.estado().setdefault("_figcache", {})
    if clave not in cache:
        with _LOCK:
            if len(cache) > 300:
                cache.clear()
            cache[clave] = fn(objeto, *args)
    return cache[clave]


def png_de(bloque):
    """Devuelve el PNG de un bloque 'img'; los gráficos se dibujan al pedirlos (y quedan en caché)."""
    if "png" in bloque:
        return bloque["png"]
    nombre, fn, objeto, args = bloque["lazy"]
    return _fig(nombre, fn, objeto, *args)


# ------------------------------------------------------------------ preparación
def preparar(ids, mina=None):
    tajos = []
    for vid in ids:
        v, meds = core.cargar_voladura(int(vid))
        if not meds:
            continue
        res = core.resumen(meds)
        obj = v["p80_objetivo"]
        sin_id = [{k: x for k, x in m.items() if k not in ("id", "voladura_id")} for m in meds]
        h = hashlib.md5(json.dumps([v["labor"], v["fecha"], v["p80_objetivo"], sin_id], sort_keys=True,
                                   default=str).encode()).hexdigest()
        tajos.append({"v": v, "meds": meds, "res": res, "par": core.parametros(meds, res), "h": h,
                      "obj": obj, "mina": mina, "meta": obj or mina,
                      "label": f"{v['labor']} · {fdate(v['fecha'])}"})
    tajos.sort(key=lambda t: (t["v"]["fecha"], t["v"]["labor"]))
    return tajos


# ------------------------------------------------------------------ textos
def _f(x, d=2):
    return fmt(x, d)


def _mm(x):
    return "-" if x is None else f"{x * IN_MM:.0f}"


def _est(p80, meta):
    """(texto, 'ok'|'bad') para celdas de tabla; '-' si no hay objetivo."""
    if p80 is None or not meta:
        return "-"
    ok = p80 <= meta
    return (f"{'Cumple' if ok else 'No cumple'} ({(p80 - meta) / meta * 100:+.1f} %)", "ok" if ok else "bad")


def _est_txt(p80, meta):
    e = _est(p80, meta)
    return e[0] if isinstance(e, tuple) else "sin objetivo definido"


def _u_n(n):
    if n is None:
        return "-"
    if n < 1.0:
        return "Baja: distribución amplia, con mezcla marcada de finos y gruesos."
    if n < 1.5:
        return "Media: dispersión moderada de tamaños alrededor del Xc."
    return "Alta: distribución concentrada alrededor del tamaño característico."


def _u_n_corta(n):
    return "-" if n is None else ("baja" if n < 1.0 else "media" if n < 1.5 else "alta")


def _u_b(b, bprom):
    if b is None:
        return "-"
    if bprom is None or abs(b - bprom) / bprom < 0.05:
        return "Similar al promedio del periodo: curva con ondulación típica."
    if b > bprom:
        return "Superior al promedio: curva más empinada alrededor del X50 (distribución más concentrada)."
    return "Inferior al promedio: curva más tendida, con mayor dispersión de tamaños."


def _u_esf(v):
    if v is None:
        return "Sin dato de esfericidad."
    if v >= 80:
        return f"Esfericidad alta ({v:.0f} %): fragmentos equidimensionales, con baja presencia de lajas."
    if v >= 60:
        return (f"Esfericidad media ({v:.0f} %): predominio de formas moderadamente equidimensionales, "
                "con presencia de piezas alargadas.")
    return (f"Esfericidad baja ({v:.0f} %): presencia importante de fragmentos alargados o lajosos, "
            "que pueden sesgar la lectura de tamaño.")


def _u_cov(v):
    if v is None:
        return "Sin dato de coverage."
    if v >= 85:
        return f"Coverage adecuado ({v:.0f} %): imagen bien cubierta por fragmentos delineados; medición representativa."
    if v >= 70:
        return f"Coverage aceptable ({v:.0f} %): cobertura suficiente; se recomienda interpretar con ligera cautela."
    return f"Coverage bajo ({v:.0f} %): cobertura limitada de la imagen; interpretar el resultado con cautela."


def _lectura_p(p, x, t):
    mm, s = _mm(x), _f(x)
    if p == 5:
        return f"El 5 % del material es menor a {mm} mm. Fracción ultrafina generada en el entorno inmediato de la carga."
    if p == 10:
        return f"El 10 % es menor a {mm} mm. Indicador del contenido de finos de la voladura."
    if p == 20:
        return f"El 20 % es menor a {mm} mm. Límite superior de la fracción fina."
    if p == 25:
        return f"Primer cuartil: el 25 % del material es menor a {mm} mm."
    if p == 50:
        return f"Tamaño mediano: la mitad del material pasa por una abertura de {mm} mm."
    if p == 75:
        return f"Tercer cuartil: el 75 % pasa por {mm} mm; el 25 % restante conforma la fracción gruesa."
    if p == 80:
        mina, obj = t["mina"], t["obj"]
        partes = []
        if mina:
            partes.append(f"frente al objetivo de mina ({_f(mina)} in): [[{_est_txt(x, mina)}]]")
        if obj:
            partes.append(f"frente al objetivo de la voladura ({_f(obj)} in): [[{_est_txt(x, obj)}]]")
        return ("**Tamaño de control.** " + "; ".join(partes) + "." if partes
                else "**Tamaño de control.** No hay objetivo definido para evaluar el cumplimiento.")
    if p == 90:
        return f"El 10 % del material es mayor a {mm} mm. Inicio de la cola gruesa."
    if p == 95:
        return f"El 5 % es mayor a {mm} mm. Fracción que condiciona el acarreo y la alimentación al chancado."
    return f"El 1 % es mayor a {mm} mm. Tamaño máximo representativo; referencia de sobretamaño."


# ------------------------------------------------------------------ contenido
def construir(tajos, meta):
    mina = meta.get("mina")
    B = []

    def add(t, **k):
        B.append({"t": t, **k})

    N = sum(t["res"]["n_img"] for t in tajos)
    M = len(tajos)
    labores = list(dict.fromkeys(t["v"]["labor"] for t in tajos))
    cuerpos = list(dict.fromkeys(t["v"]["cuerpo"] for t in tajos if t["v"]["cuerpo"]))
    periodo = meta.get("periodo") or "-"
    portada = [("Cliente", meta.get("cliente") or "-"), ("Periodo", periodo),
               ("Labores", ", ".join(labores) or "-"), ("Cuerpos", ", ".join(cuerpos) or "-"),
               ("Análisis granulométricos realizados", str(N)),
               ("Elaborado por", meta.get("elaborado_por") or "-"),
               ("Fecha de emisión", meta.get("emision") or datetime.now().strftime("%d/%m/%Y"))]
    if not tajos:
        return {"portada": portada, "blocks": []}

    p80s = [t["res"]["P80"] for t in tajos if t["res"]["P80"] is not None]
    w80 = (sum(t["res"]["P80"] * t["res"]["n_img"] for t in tajos if t["res"]["P80"] is not None)
           / max(sum(t["res"]["n_img"] for t in tajos if t["res"]["P80"] is not None), 1)) if p80s else None

    # 1 -----------------------------------------------------------------------
    add("h1", text="1. Objetivo")
    for par in (meta.get("objetivo_txt") or DEF_OBJETIVO).split("\n\n"):
        if par.strip():
            add("p", text=par.strip())
    add("p", text="Objetivos específicos:")
    add("bullets", items=[
        "Determinar la distribución granulométrica de cada tajo, desde el P-05 hasta el P-99.",
        "Verificar el cumplimiento del P80 frente al objetivo de mina y al objetivo de cada voladura.",
        "Caracterizar la uniformidad (Rosin-Rammler) y la forma de la curva (Swebrec).",
        "Validar la calidad de cada muestra mediante la esfericidad y el coverage de la imagen.",
        "Emitir conclusiones y recomendaciones para la operación."])

    # 2 -----------------------------------------------------------------------
    add("h1", text="2. Datos generales")
    add("kv", title="DATOS GENERALES", rows=[
        ("Cliente", meta.get("cliente") or "-"), ("Periodo", periodo),
        ("Tipo de voladura", "Taladros largos"), ("Tajos evaluados", str(M)),
        ("Labores", ", ".join(labores)), ("Cuerpos", ", ".join(cuerpos) or "-"),
        ("Análisis granulométricos realizados", str(N)), ("Software de análisis", "WipFrag 4"),
        ("P80 objetivo de mina", f"{_f(mina)} in ({_mm(mina)} mm)" if mina else "No definido"),
        ("Elaborado por", meta.get("elaborado_por") or "-")])
    add("table", head=["Tajo (labor)", "Fecha", "Turno", "Cuerpo", "Tipo de roca", "N° de muestras"],
        rows=[[t["v"]["labor"], fdate(t["v"]["fecha"]), t["v"]["turno"] or "-", t["v"]["cuerpo"] or "-",
               t["v"]["tipo_roca"] or "-", str(t["res"]["n_img"])] for t in tajos],
        widths=[3.0, 2.4, 4.2, 2.8, 2.8, 1.8], align="lcccc c".replace(" ", ""),
        note=f"Total: {N} análisis granulométricos en {M} tajos.")

    # 3 -----------------------------------------------------------------------
    add("h1", text="3. Parámetros de cálculo")
    add("kv", title="PARÁMETROS DE ANÁLISIS", rows=[
        ("Software", "WipFrag 4"),
        ("Unidad de tamaño", "Pulgadas (1 in = 25.4 mm)"),
        ("Mallas reportadas", "1.5, 2, 3, 4, 5, 6, 8, 10, 12, 14, 16, 18 y 20 in"),
        ("Percentiles reportados", "P-05, P-10, P-20, P-25, P-50, P-75, P-80, P-90, P-95 y P-99"),
        ("P80 objetivo de mina", f"{_f(mina)} in ({_mm(mina)} mm)" if mina else "No definido"),
        ("Criterio de cumplimiento", "P80 menor o igual al objetivo"),
        ("Consolidación por tajo", "Promedio de las muestras de cada voladura")])
    add("table", head=["Tajo", "Burden (m)", "Espac. (m)", "Ø (mm)", "Volumen (m3)", "Tonelaje (t)",
                       "Anfo / Emulsión (kg)", "Total expl. (kg)", "F. potencia (kg/t)"],
        rows=[[t["label"], _f(t["v"]["burden"]), _f(t["v"]["espaciamiento"]), _f(t["v"]["diametro"], 0),
               _f(t["v"]["volumen"], 0), _f(t["v"]["tonelaje"], 0),
               f"{_f(t['v']['anfo'], 0)} / {_f(t['v']['emulsion'], 0)}",
               _f(t["v"]["total_explosivo"], 0), _f(t["v"]["factor_potencia"], 3)] for t in tajos],
        widths=[3.2, 1.5, 1.5, 1.2, 1.8, 1.8, 2.3, 1.9, 1.8], align="l" + "c" * 8,
        note="Parámetros de diseño de la voladura registrados en la base de datos del cliente.")

    # 4 -----------------------------------------------------------------------
    add("h1", text="4. Metodología")
    add("bullets", items=[x.strip() for x in (meta.get("metodologia_txt") or DEF_METODOLOGIA).split("\n")
                          if x.strip()])

    # 5 -----------------------------------------------------------------------
    add("h1", text="5. Software de análisis granulométrico")
    add("p", text="**WipFrag 4** es un software de análisis granulométrico por imágenes digitales: mide el "
                  "tamaño de los fragmentos a partir de fotografías del material, sin necesidad de tamizado "
                  "físico. Para cada imagen entrega la curva de distribución acumulada, los tamaños "
                  "característicos y los parámetros de ajuste de las distribuciones de Rosin-Rammler y "
                  "Swebrec, que se exportan a una hoja de cálculo para su consolidación.")
    add("p", text="Salidas utilizadas en este informe:")
    add("bullets", items=["Tamaños característicos: Xmax, X50 y Xc.",
                          "Porcentaje pasante por malla, de 1.5 a 20 in.",
                          "Percentiles P-05 a P-99.",
                          "Coeficiente de uniformidad de Rosin-Rammler (n) y parámetro de forma de Swebrec (b).",
                          "Esfericidad y coverage de cada imagen."])

    # 6 -----------------------------------------------------------------------
    add("h1", text="6. Curva granulométrica e interpretación")
    add("p", text="La curva granulométrica acumulada representa el porcentaje en peso del material que pasa por "
                  "cada abertura. Se lee de izquierda a derecha: una curva desplazada hacia la izquierda indica "
                  "material más fino y hacia la derecha, material más grueso. La pendiente refleja la "
                  "uniformidad: cuanto más empinada, más concentrado es el rango de tamaños. El **P80**, tamaño "
                  "por el que pasa el 80 % del material, es el indicador de control de la fragmentación.")
    add("img", lazy=('gen', graficos.fig_curvas_generales, tajos, (mina,)), width=16,
        caption="Curvas granulométricas promedio por tajo.")
    orden = sorted([t for t in tajos if t["res"]["P80"] is not None], key=lambda t: t["res"]["P80"])
    if len(orden) > 1:
        a, z = orden[0], orden[-1]
        add("p", text=f"La curva más fina corresponde al tajo **{a['label']}** (P80 = {_f(a['res']['P80'])} in, "
                      f"{_mm(a['res']['P80'])} mm) y la más gruesa al tajo **{z['label']}** "
                      f"(P80 = {_f(z['res']['P80'])} in, {_mm(z['res']['P80'])} mm), una diferencia de "
                      f"{(z['res']['P80'] / a['res']['P80'] - 1) * 100:.0f} % entre ambos.")
    add("table", head=["Tajo", "N°"] + [f"P-{p:02d}" for p in PCTS],
        rows=[[t["label"], str(t["res"]["n_img"])] + [_f(t["res"][f"P{p}"]) for p in PCTS] for t in tajos],
        widths=[3.4, 0.9] + [1.27] * 10, align="lc" + "c" * 10,
        note="Tamaños en pulgadas; promedio de las muestras de cada tajo.")

    # 7 -----------------------------------------------------------------------
    add("h1", text="7. Coeficiente de uniformidad de Rosin-Rammler")
    add("p", text="La distribución de Rosin-Rammler describe la curva granulométrica con dos parámetros: el "
                  "tamaño característico **Xc** (abertura por la que pasa el 63.2 % del material) y el "
                  "coeficiente de uniformidad **n**. Un n mayor indica una distribución más uniforme; un n menor, "
                  "una mezcla más amplia de finos y gruesos.")
    add("eq", text="P(x) = 1 - exp[ -(x / Xc)^n ]")
    fu = {t["par"]["n_f"] for t in tajos}
    if fu == {"WipFrag"}:
        com = "En este informe, el n de cada tajo es el promedio de sus muestras, tal como lo entrega WipFrag 4."
    elif fu == {"Ajuste"}:
        com = ("El archivo no incluyó el valor de n, por lo que se calculó por regresión lineal sobre los "
               "percentiles P-05 a P-99 de cada tajo.")
    else:
        com = ("El n de cada tajo es el promedio de sus muestras según WipFrag 4; donde el archivo no lo "
               "incluyó (columna Fuente = Ajuste) se calculó por regresión sobre los percentiles P-05 a P-99.")
    add("p", text="**Cómo se obtuvo.** WipFrag 4 ajusta la distribución a la curva de cada imagen. Al linealizar la "
                  "curva en ejes ln(x) frente a ln[-ln(1-P)], se obtiene una recta cuya pendiente es n y cuyo "
                  "punto al 63.2 % define Xc. " + com)
    add("img", lazy=('rr', graficos.fig_ajuste, tajos, ('rr',)), width=16,
        caption="Ajuste de Rosin-Rammler: puntos medidos y curva del modelo por tajo.")
    add("table", head=["Tajo", "Xc (in)", "Xc (mm)", "n", "Lectura de la uniformidad", "Fuente"],
        rows=[[t["label"], _f(t["par"]["xc"]), _mm(t["par"]["xc"]), _f(t["par"]["n"]),
               _u_n(t["par"]["n"]), t["par"]["n_f"]] for t in tajos],
        widths=[3.4, 1.5, 1.5, 1.2, 7.6, 1.8], align="lcccl c".replace(" ", ""),
        note="Referencia de lectura de n: menor a 1.0 baja; de 1.0 a 1.5 media; mayor a 1.5 alta.")

    # 8 -----------------------------------------------------------------------
    add("h1", text="8. Parámetro de forma (ondulación) de la distribución Swebrec")
    add("p", text="La función Swebrec describe la curva granulométrica a partir del tamaño máximo **Xmax**, el "
                  "tamaño mediano **X50** y el parámetro de forma **b**, que controla la ondulación de la curva. "
                  "A diferencia de Rosin-Rammler, representa mejor la cola de finos y el tamaño máximo. Un b mayor "
                  "produce una curva más empinada alrededor del X50 (distribución más concentrada); un b menor, "
                  "una curva más tendida.")
    add("eq", text="P(x) = 1 / { 1 + [ ln(Xmax / x) / ln(Xmax / X50) ]^b }")
    fb = {t["par"]["b_f"] for t in tajos}
    est = any(t["par"]["xmax_f"] == "estimado" for t in tajos)
    if fb == {"WipFrag"}:
        com = "En este informe, el b de cada tajo es el promedio de sus muestras, tal como lo entrega WipFrag 4."
    else:
        com = ("Donde el archivo no incluyó b (columna Fuente = Ajuste), se calculó por regresión sobre los "
               "percentiles P-05 a P-99, tomando X50 y Xmax del tajo" +
               ("; cuando faltó Xmax se estimó como 1.05 veces el P-99." if est else "."))
    add("p", text="**Cómo se obtuvo.** WipFrag 4 ajusta la función a la curva de cada imagen. Al linealizarla, "
                  "ln[ln(Xmax/x) / ln(Xmax/X50)] frente a ln(1/P - 1) forma una recta cuya pendiente es b. " + com)
    add("img", lazy=('sw', graficos.fig_ajuste, tajos, ('sw',)), width=16,
        caption="Ajuste de Swebrec: puntos medidos y curva del modelo por tajo.")
    bs = [t["par"]["b"] for t in tajos if t["par"]["b"]]
    bprom = float(np.mean(bs)) if bs else None
    add("table", head=["Tajo", "Xmax (in)", "X50 (in)", "b", "Lectura de la ondulación", "Fuente"],
        rows=[[t["label"], _f(t["par"]["xmax"]), _f(t["par"]["x50"]), _f(t["par"]["b"]),
               _u_b(t["par"]["b"], bprom), t["par"]["b_f"]] for t in tajos],
        widths=[3.4, 1.6, 1.5, 1.1, 7.6, 1.8], align="lcccl c".replace(" ", ""),
        note="La lectura de b es relativa al promedio de los tajos del periodo.")

    # 9 -----------------------------------------------------------------------
    add("pb")
    add("h1", text="9. Análisis por tajo")
    add("p", text="Para cada tajo se presenta la curva granulométrica con los percentiles destacados, la "
                  "distribución por rangos de tamaño, la lectura de los valores de P-05 a P-99 con énfasis en el "
                  "cumplimiento del P80, y el control de calidad de cada muestra (esfericidad y coverage).")
    for i, t in enumerate(tajos, 1):
        v, r = t["v"], t["res"]
        if i > 1:
            add("pb")
        add("h2", text=f"9.{i} Tajo {t['label']}")
        add("p", text=f"Tajo **{v['labor']}**, fecha **{fdate(v['fecha'])}**, cuerpo **{v['cuerpo'] or '-'}**, "
                      f"tipo de roca **{v['tipo_roca'] or '-'}**. Se analizaron **{r['n_img']}** muestras.")
        add("h3", text="Curva y distribución granulométrica")
        add("img", lazy=('curva', graficos.fig_curva_tajo, t, (mina,)), width=14,
            caption=f"Curva granulométrica del tajo {t['label']}.")
        add("img", lazy=('hist', graficos.fig_hist, t, ()), width=13,
            caption=f"Distribución por rangos de tamaño del tajo {t['label']}.")
        rows = []
        for k, m in enumerate(t["meds"], 1):
            rows.append([f"M{k}"] + [(_f(m["pcts"].get(str(p))) if p != 80 else
                                      ((_f(m["pcts"].get("80")), _est(m["pcts"].get("80"), t["meta"])[1])
                                       if t["meta"] and m["pcts"].get("80") is not None
                                       else _f(m["pcts"].get("80")))) for p in PCTS])
        rows.append(["**Prom.**"] + [f"**{_f(r[f'P{p}'])}**" for p in PCTS])
        add("table", head=["Muestra"] + [f"P-{p:02d}" for p in PCTS], rows=rows,
            widths=[1.9] + [1.51] * 10, align="c" * 11,
            note="Tamaños en pulgadas. En la columna P-80 se marca si cada muestra cumple "
                 "(objetivo de la voladura, o de mina si no existe).")
        add("h3", text="Interpretación de los percentiles P-05 a P-99")
        add("table", head=["Percentil", "in", "mm", "Lectura"],
            rows=[[f"P-{p:02d}", _f(r[f"P{p}"]), _mm(r[f"P{p}"]), _lectura_p(p, r[f"P{p}"], t)] for p in PCTS],
            widths=[1.7, 1.3, 1.3, 12.7], align="cccl")
        p80, p50 = r["P80"], r["P50"]
        txt = (f"El tajo presenta un P50 de **{_f(p50)} in ({_mm(p50)} mm)** y un P80 de "
               f"**{_f(p80)} in ({_mm(p80)} mm)**. ")
        if mina:
            txt += f"Frente al objetivo de mina ({_f(mina)} in): [[{_est_txt(p80, mina)}]]. "
        if t["obj"]:
            txt += f"Frente al objetivo de la voladura ({_f(t['obj'])} in): [[{_est_txt(p80, t['obj'])}]]. "
        p80m = [m["pcts"].get("80") for m in t["meds"] if m["pcts"].get("80") is not None]
        if len(p80m) > 1:
            cv = np.std(p80m) / np.mean(p80m) * 100
            txt += (f"El P80 entre muestras varía de {_f(min(p80m))} a {_f(max(p80m))} in "
                    f"(coeficiente de variación de {cv:.1f} %), "
                    f"{'lo que indica una fragmentación homogénea en el tajo.' if cv < 10 else 'lo que sugiere variabilidad de la fragmentación dentro del tajo.'}")
        add("p", text=txt)
        add("h3", text="Esfericidad y coverage por muestra")
        add("img", lazy=('fc', graficos.fig_forma_cov, t, ()), width=15,
            caption=f"Esfericidad y coverage por muestra del tajo {t['label']}.")
        add("table", head=["Muestra", "Esfericidad (%)", "Coverage (%)", "Interpretación"],
            rows=[[f"M{k}", _f(m["sphericity"], 0), _f(m["coverage"], 0),
                   _u_esf(m["sphericity"]) + " " + _u_cov(m["coverage"])] for k, m in enumerate(t["meds"], 1)],
            widths=[1.6, 2.2, 2.0, 11.2], align="ccc" + "l",
            note="Criterio referencial: esfericidad alta desde 80 %, media de 60 a 80 %; coverage adecuado desde "
                 "85 %, aceptable de 70 a 85 %.")

    # 10 ----------------------------------------------------------------------
    add("pb")
    add("h1", text="10. Resumen de resultados")
    con = [t for t in tajos if t["meta"] and t["res"]["P80"] is not None]
    ok = [t for t in con if t["res"]["P80"] <= t["meta"]]
    ko = [t for t in con if t["res"]["P80"] > t["meta"]]
    add("img", lazy=('p80', graficos.fig_p80_tajos, tajos, (mina,)), width=15.5, caption="P80 por tajo frente al objetivo.")
    add("table", head=["Tajo", "Cuerpo", "N°", "P50 (in)", "P80 (in)", "P80 (mm)", "Objetivo (in)", "Estado", "n", "b"],
        rows=[[t["label"], t["v"]["cuerpo"] or "-", str(t["res"]["n_img"]), _f(t["res"]["P50"]),
               _f(t["res"]["P80"]), _mm(t["res"]["P80"]), _f(t["meta"]), _est(t["res"]["P80"], t["meta"]),
               _f(t["par"]["n"]), _f(t["par"]["b"])] for t in tajos],
        widths=[3.0, 2.5, 0.8, 1.4, 1.4, 1.3, 1.5, 2.9, 1.1, 1.1], align="lcc" + "c" * 7,
        note="Objetivo: de la voladura si existe; de lo contrario, el de mina.")
    puntos = [f"Se realizaron **{N} análisis granulométricos** en **{M} tajos**"
              f"{' de ' + str(len(cuerpos)) + ' cuerpo(s) (' + ', '.join(cuerpos) + ')' if cuerpos else ''}, "
              "mediante WipFrag 4."]
    if con:
        s = (f"**Cumplimiento del P80:** {len(ok)} de {len(con)} tajos ({len(ok) / len(con) * 100:.0f} %) "
             "se ubican dentro del objetivo.")
        if ko:
            s += " No cumplen: " + "; ".join(f"{t['label']} ({(t['res']['P80'] - t['meta']) / t['meta'] * 100:+.1f} %)"
                                              for t in ko) + "."
        puntos.append(s)
    else:
        puntos.append("No se definió un P80 objetivo; los resultados constituyen la línea base del periodo.")
    if p80s:
        puntos.append(f"El P80 promedio ponderado por muestras fue **{_f(w80)} in ({_mm(w80)} mm)**, con un rango "
                      f"de {_f(min(p80s))} a {_f(max(p80s))} in entre tajos.")
    ns = [t["par"]["n"] for t in tajos if t["par"]["n"]]
    if ns:
        puntos.append(f"El coeficiente de uniformidad n promedio fue **{np.mean(ns):.2f}** "
                      f"(uniformidad {_u_n_corta(float(np.mean(ns)))})"
                      + (f" y el parámetro b de Swebrec promedio, **{bprom:.2f}**." if bprom else "."))
    esf = [t["res"]["sphericity"] for t in tajos if t["res"]["sphericity"] is not None]
    cov = [t["res"]["coverage"] for t in tajos if t["res"]["coverage"] is not None]
    if esf or cov:
        puntos.append("Control de calidad de las imágenes: "
                      + (f"esfericidad promedio de {np.mean(esf):.0f} %" if esf else "")
                      + (" y " if esf and cov else "")
                      + (f"coverage promedio de {np.mean(cov):.0f} %" if cov else "") + ".")
    if len(orden) > 1:
        puntos.append(f"Mejor fragmentación: **{orden[0]['label']}** (P80 = {_f(orden[0]['res']['P80'])} in). "
                      f"Mayor P80: **{orden[-1]['label']}** ({_f(orden[-1]['res']['P80'])} in).")
    add("bullets", items=puntos)

    # 11 ----------------------------------------------------------------------
    add("h1", text="11. Conclusiones")
    c = []
    if con:
        if not ko:
            c.append("La fragmentación del periodo **cumple** el P80 objetivo en todos los tajos evaluados.")
        elif not ok:
            c.append("La fragmentación del periodo **no cumple** el P80 objetivo en ninguno de los tajos evaluados.")
        else:
            c.append(f"La fragmentación del periodo cumple el P80 objetivo en {len(ok)} de {len(con)} tajos; "
                     f"los {len(ko)} restantes lo superan.")
    if len(orden) > 1:
        c.append(f"El tajo {orden[0]['label']} presenta la fragmentación más fina (P80 = "
                 f"{_f(orden[0]['res']['P80'])} in) y el tajo {orden[-1]['label']} la más gruesa "
                 f"(P80 = {_f(orden[-1]['res']['P80'])} in).")
    if ns:
        c.append(f"La uniformidad de la distribución es {_u_n_corta(float(np.mean(ns)))} "
                 f"(n promedio = {np.mean(ns):.2f}).")
    fp_ok = [t["v"]["factor_potencia"] for t in ok if t["v"]["factor_potencia"]]
    fp_ko = [t["v"]["factor_potencia"] for t in ko if t["v"]["factor_potencia"]]
    if fp_ok and fp_ko:
        c.append(f"Los tajos que no cumplen registran un factor de potencia promedio de {np.mean(fp_ko):.2f} kg/t, "
                 f"frente a {np.mean(fp_ok):.2f} kg/t en los que cumplen.")
    if cov:
        c.append(f"El coverage promedio de {np.mean(cov):.0f} % "
                 + ("respalda la representatividad de las mediciones." if np.mean(cov) >= 85
                    else "sugiere interpretar con cautela los resultados de las imágenes de menor cobertura."))
    add("bullets", items=c or ["Los resultados del periodo constituyen la línea base de fragmentación."])

    # 12 ----------------------------------------------------------------------
    add("h1", text="12. Recomendaciones")
    r = []
    if ko:
        lista = ", ".join(t["label"] for t in ko)
        r.append(f"Revisar el diseño de malla y el factor de potencia de los tajos que superan el objetivo ({lista}), "
                 "priorizando la reducción de burden y espaciamiento o el incremento del factor de potencia, "
                 "y validar el ajuste con una voladura de prueba controlada.")
    if ns and np.mean(ns) < 1.5:
        r.append("Evaluar la secuencia y los tiempos de iniciación, así como la distribución de la carga en el "
                 "taladro, para reducir la dispersión de tamaños y elevar la uniformidad de la fragmentación.")
    if ok:
        r.append("Estandarizar y replicar los parámetros de diseño y carguío de los tajos que cumplen "
                 f"({', '.join(t['label'] for t in ok)}) en las labores de condiciones geomecánicas similares.")
    r.append("Compartir estos resultados con las áreas de acarreo y planta, a fin de validar el P80 objetivo "
             "frente al rendimiento real de la conminución.")
    r.append("Mantener el control semanal de fragmentación con el mismo protocolo, para consolidar la línea base "
             "y evaluar la tendencia de cada tajo.")
    add("bullets", items=r)
    add("cierre")
    return {"portada": portada, "blocks": B}


# ------------------------------------------------------------------ marcado
def _mk(t):
    t = escape(str(t))
    t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
    return re.sub(r"\[\[(.+?)\]\]", rf'<font color="{ROJO}"><b>\1</b></font>', t)


def _cell(c):
    return c if isinstance(c, tuple) else (c, None)
