"""Datos: base de datos SQLite, lectura del Excel de WipFrag 4 y ajuste de distribuciones."""
import json
import re
import sqlite3
import threading
import unicodedata
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

BASE = Path(__file__).parent
_TL = threading.local()   # sesión del visitante que está ejecutando en este hilo
IN_MM = 25.4
SIZES = [1.5, 2, 3, 4, 5, 6, 8, 10, 12, 14, 16, 18, 20]   # columnas "1.5 in" ... "20 in"
PCTS = [5, 10, 20, 25, 50, 75, 80, 90, 95, 99]            # columnas "P-05" ... "P-99"

# campo en base de datos -> encabezado en el Excel de WipFrag
CAMPOS = {
    "periodo": "Periodo", "turno": "Turno", "guardia": "Guardia",
    "burden": "Burden", "espaciamiento": "Espaciamiento", "diametro": "Diámetro",
    "vol_planificado": "Volumen Planificado(m3)", "volumen": "Volumen (m3)",
    "tonelaje": "Tonelaje (Ton)", "factor_potencia": "Factor de Potencia (kg/Ton)",
    "anfo": "Anfo (kg)", "emulsion": "Emulsión (kg)",
    "total_explosivo": "Total de Explosivo (kg)",
    "cuerpo": "Cuerpo", "tipo_roca": "Tipo de Roca",
}
NUM = {"burden", "espaciamiento", "diametro", "vol_planificado", "volumen", "tonelaje",
       "factor_potencia", "anfo", "emulsion", "total_explosivo", "p80_objetivo"}
EDITABLES = ["turno", "guardia", "tipo_roca", "cuerpo", "burden", "espaciamiento", "diametro",
             "vol_planificado", "volumen", "tonelaje", "anfo", "emulsion", "total_explosivo",
             "factor_potencia", "p80_objetivo"]


# ----------------------------------------------------------------- utilidades
def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]", "", s)


def num(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    if isinstance(v, str):
        s = v.strip().replace(",", ".").rstrip("%").strip()
        try:
            return float(s)
        except ValueError:
            return None
    return float(v)


def pct(v):
    """Porcentaje en escala 0-100 (acepta '56.30%', 56.3 o 0.563)."""
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    if isinstance(v, str):
        return num(v)
    x = float(v)
    return x * 100 if x <= 1 else x


def fmt(x, d=2):
    return "-" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:,.{d}f}"


def fdate(s):
    try:
        return datetime.strptime(s, "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        return s or "-"


def parse_fecha(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    d = pd.to_datetime(v, dayfirst=True, errors="coerce")
    return None if pd.isna(d) else d.strftime("%Y-%m-%d")


# ------------------------------------------------------------- base de datos
def vincular(estado):
    """Asocia este hilo con el estado de UNA sesión (st.session_state). Nada se guarda en disco:
    la base de datos vive en memoria y desaparece al cerrar o recargar la página."""
    _TL.estado = estado


def estado():
    return _TL.estado


def conn():
    e = _TL.estado
    c = e.get("_db")
    if c is None:
        c = sqlite3.connect(":memory:", check_same_thread=False)
        c.row_factory = sqlite3.Row
        c.execute("PRAGMA foreign_keys=ON")
        e["_db"] = c
    return c


def init_db():
    with conn() as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS voladuras(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            labor TEXT NOT NULL, fecha TEXT NOT NULL,
            periodo TEXT, turno TEXT, guardia TEXT,
            burden REAL, espaciamiento REAL, diametro REAL,
            vol_planificado REAL, volumen REAL, tonelaje REAL, factor_potencia REAL,
            anfo REAL, emulsion REAL, total_explosivo REAL,
            cuerpo TEXT, tipo_roca TEXT, p80_objetivo REAL,
            UNIQUE(labor, fecha));
        CREATE TABLE IF NOT EXISTS mediciones(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            voladura_id INTEGER NOT NULL REFERENCES voladuras(id) ON DELETE CASCADE,
            sphericity REAL, coverage REAL, n_rr REAL, xc REAL, swebrec_b REAL, xmax REAL, x50 REAL,
            sizes TEXT, pcts TEXT);
        CREATE TABLE IF NOT EXISTS config(clave TEXT PRIMARY KEY, valor REAL);
        CREATE TABLE IF NOT EXISTS ajustes(clave TEXT PRIMARY KEY, valor TEXT);
        """)
        cols = [r["name"] for r in c.execute("PRAGMA table_info(mediciones)")]
        if "coverage" not in cols:                     # bases creadas con versiones anteriores
            c.execute("ALTER TABLE mediciones ADD COLUMN coverage REAL")


def datos_version():
    """Huella de los tajos y muestras: cambia solo si cambian los datos (no los id internos ni los textos de portada)."""
    import hashlib
    h = hashlib.md5()
    with conn() as c:
        for q in ("SELECT labor,fecha,burden,espaciamiento,diametro,volumen,tonelaje,factor_potencia,anfo,emulsion,"
                  "total_explosivo,cuerpo,tipo_roca,p80_objetivo,turno,guardia FROM voladuras ORDER BY labor,fecha",
                  "SELECT v.labor,v.fecha,m.sphericity,m.coverage,m.n_rr,m.xc,m.swebrec_b,m.xmax,m.x50,m.sizes,m.pcts "
                  "FROM mediciones m JOIN voladuras v ON v.id=m.voladura_id ORDER BY v.labor,v.fecha,m.id"):
            for r in c.execute(q):
                h.update(repr(tuple(r)).encode())
    return h.hexdigest()


def get_config(clave):
    with conn() as c:
        r = c.execute("SELECT valor FROM config WHERE clave=?", (clave,)).fetchone()
    return r["valor"] if r else None


def set_config(clave, valor):
    with conn() as c:
        c.execute("INSERT INTO config(clave,valor) VALUES(?,?) "
                  "ON CONFLICT(clave) DO UPDATE SET valor=excluded.valor", (clave, valor))


def get_ajuste(clave, default=""):
    with conn() as c:
        r = c.execute("SELECT valor FROM ajustes WHERE clave=?", (clave,)).fetchone()
    return r["valor"] if r and r["valor"] is not None else default


def set_ajuste(clave, valor):
    with conn() as c:
        c.execute("INSERT INTO ajustes(clave,valor) VALUES(?,?) "
                  "ON CONFLICT(clave) DO UPDATE SET valor=excluded.valor", (clave, valor))


def leer_archivo(f):
    name = f.name.lower()
    if name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(f)
    else:
        df = pd.read_csv(f, sep=None, engine="python")
    df.columns = [str(c).strip() for c in df.columns]
    return df.dropna(how="all")


def importar(df):
    """Cada fila = una imagen/muestra; Labor+Fecha = una voladura (tajo)."""
    m = {norm(c): c for c in df.columns}

    def g(r, name):
        k = m.get(norm(name))
        if k is None:
            return None
        v = r[k]
        return None if (not isinstance(v, str) and pd.isna(v)) else v

    vistas, n_med = set(), 0
    with conn() as c:
        for _, r in df.iterrows():
            labor = str(g(r, "Labor") or "").strip()
            fecha = parse_fecha(g(r, "Fecha"))
            if not labor or not fecha:
                continue
            c.execute("INSERT OR IGNORE INTO voladuras(labor,fecha) VALUES(?,?)", (labor, fecha))
            vid = c.execute("SELECT id FROM voladuras WHERE labor=? AND fecha=?",
                            (labor, fecha)).fetchone()["id"]
            for k, col in CAMPOS.items():      # solo pisa lo que venga con dato
                v = g(r, col)
                v = num(v) if k in NUM else (None if v is None else str(v).strip())
                if v is not None:
                    c.execute(f"UPDATE voladuras SET {k}=? WHERE id=?", (v, vid))
            if vid not in vistas:               # reimportar = reemplaza las muestras
                c.execute("DELETE FROM mediciones WHERE voladura_id=?", (vid,))
                vistas.add(vid)
            sizes = {f"{s:g}": pct(g(r, f"{s:g} in")) for s in SIZES}
            pcts = {str(p): num(g(r, f"P-{p:02d}")) for p in PCTS}
            c.execute(
                "INSERT INTO mediciones(voladura_id,sphericity,coverage,n_rr,xc,swebrec_b,xmax,x50,sizes,pcts)"
                " VALUES(?,?,?,?,?,?,?,?,?,?)",
                (vid, pct(g(r, "Sphericity")), pct(g(r, "Coverage")),
                 num(g(r, "Coeficiente de Uniformidad Rosin-Rammler (n)")),
                 num(g(r, "Xc")), num(g(r, "Distribución Swebrec (b)")),
                 num(g(r, "Xmax")), num(g(r, "X50")),
                 json.dumps(sizes), json.dumps(pcts)))
            n_med += 1
    return len(vistas), n_med


def voladuras_df():
    cols = ["id", "labor", "fecha"] + EDITABLES
    with conn() as c:
        df = pd.read_sql_query(f"SELECT {','.join(cols)} FROM voladuras ORDER BY fecha DESC, labor", c)
    for k in NUM:
        if k in df.columns:
            df[k] = pd.to_numeric(df[k], errors="coerce")
    return df


def _factor(vals):
    tot = vals["total_explosivo"] or ((vals["anfo"] or 0) + (vals["emulsion"] or 0)) or None
    if tot and not vals["total_explosivo"]:
        vals["total_explosivo"] = tot
    if vals["factor_potencia"] is None and tot and vals["tonelaje"]:
        vals["factor_potencia"] = round(tot / vals["tonelaje"], 3)
    return vals


def guardar_voladuras(df):
    with conn() as c:
        for _, r in df.iterrows():
            vals = {k: (None if pd.isna(r[k]) else (float(r[k]) if k in NUM else str(r[k])))
                    for k in EDITABLES}
            vals = _factor(vals)
            c.execute(f"UPDATE voladuras SET {','.join(k + '=?' for k in vals)} WHERE id=?",
                      list(vals.values()) + [int(r["id"])])


def actualizar_voladura(vid, campos):
    """Guarda campos sueltos de una voladura y recalcula el factor de potencia si se puede."""
    with conn() as c:
        for k, val in campos.items():
            if k in EDITABLES:
                c.execute(f"UPDATE voladuras SET {k}=? WHERE id=?", (val, vid))
        r = dict(c.execute("SELECT tonelaje,total_explosivo,anfo,emulsion,factor_potencia FROM voladuras "
                           "WHERE id=?", (vid,)).fetchone())
        tot = r["total_explosivo"] or ((r["anfo"] or 0) + (r["emulsion"] or 0)) or None
        if tot and r["tonelaje"]:
            c.execute("UPDATE voladuras SET factor_potencia=? WHERE id=?", (round(tot / r["tonelaje"], 3), vid))


def cargar_voladura(vid):
    with conn() as c:
        v = dict(c.execute("SELECT * FROM voladuras WHERE id=?", (vid,)).fetchone())
        meds = [dict(x) for x in c.execute(
            "SELECT * FROM mediciones WHERE voladura_id=? ORDER BY id", (vid,))]
    for md in meds:
        md["sizes"], md["pcts"] = json.loads(md["sizes"]), json.loads(md["pcts"])
    return v, meds


# ------------------------------------------------------------ resultados
def _prom(vals):
    vals = [x for x in vals if x is not None]
    return float(np.mean(vals)) if vals else None


def resumen(meds):
    r = {f"P{p}": _prom([m["pcts"].get(str(p)) for m in meds]) for p in PCTS}
    r["sphericity"] = _prom([m["sphericity"] for m in meds])
    r["coverage"] = _prom([m["coverage"] for m in meds])
    r["n_img"] = len(meds)
    r["sizes"] = {f"{s:g}": _prom([m["sizes"].get(f"{s:g}") for m in meds]) for s in SIZES}
    return r


def curva_rr(x, n, xc):
    x = np.asarray(x, float)
    return 100 * (1 - np.exp(-(x / xc) ** n))


def curva_sw(x, b, xmax, x50):
    x = np.asarray(x, float)
    out = np.full_like(x, 100.0)
    ok = x < xmax
    q = np.log(xmax / x[ok]) / np.log(xmax / x50)
    out[ok] = 100 / (1 + q ** b)
    return out


def ajuste_rr(pts):
    """Regresión sobre ln(x) vs ln(-ln(1-P)). n = pendiente; Xc = tamaño al 63.2 %."""
    pts = [(x, p) for x, p in pts if x and x > 0 and 0 < p < 100]
    if len(pts) < 3:
        return None
    X = np.log([x for x, _ in pts])
    Y = np.log(-np.log(1 - np.array([p for _, p in pts]) / 100))
    n, b0 = np.polyfit(X, Y, 1)
    r2 = 1 - np.sum((Y - (n * X + b0)) ** 2) / max(np.sum((Y - Y.mean()) ** 2), 1e-12)
    return {"n": float(n), "xc": float(np.exp(-b0 / n)), "r2": float(r2)}


def ajuste_swebrec(pts, xmax, x50):
    """Regresión por el origen sobre ln[ln(Xmax/x)/ln(Xmax/X50)] vs ln(1/P-1). b = pendiente."""
    t, y = [], []
    for x, p in pts:
        if x and 0 < x < xmax and 0 < p < 100 and abs(np.log(xmax / x) / np.log(xmax / x50) - 1) > 1e-6:
            t.append(np.log(np.log(xmax / x) / np.log(xmax / x50)))
            y.append(np.log(100 / p - 1))
    if len(t) < 3:
        return None
    t, y = np.array(t), np.array(y)
    b = float(np.sum(t * y) / np.sum(t * t))
    r2 = 1 - np.sum((y - b * t) ** 2) / max(np.sum((y - y.mean()) ** 2), 1e-12)
    return {"b": b, "r2": float(r2)}


def parametros(meds, res):
    """n, Xc, b, Xmax, X50 de un tajo: promedio de lo que trae WipFrag; si falta, se calcula."""
    pts = [(res[f"P{p}"], float(p)) for p in PCTS if res[f"P{p}"]]
    rep = {k: _prom([m[k] for m in meds]) for k in ("n_rr", "xc", "swebrec_b", "xmax", "x50")}
    rr = ajuste_rr(pts)
    x50 = rep["x50"] or res["P50"]
    xmax, xmax_f = rep["xmax"], "WipFrag"
    if not xmax and res["P99"]:
        xmax, xmax_f = res["P99"] * 1.05, "estimado"
    sw = ajuste_swebrec(pts, xmax, x50) if (xmax and x50 and xmax > x50) else None
    return {
        "n": rep["n_rr"] or (rr["n"] if rr else None), "n_f": "WipFrag" if rep["n_rr"] else "Ajuste",
        "xc": rep["xc"] or (rr["xc"] if rr else None), "xc_f": "WipFrag" if rep["xc"] else "Ajuste",
        "b": rep["swebrec_b"] or (sw["b"] if sw else None), "b_f": "WipFrag" if rep["swebrec_b"] else "Ajuste",
        "xmax": xmax, "xmax_f": xmax_f, "x50": x50, "pts": pts,
    }
