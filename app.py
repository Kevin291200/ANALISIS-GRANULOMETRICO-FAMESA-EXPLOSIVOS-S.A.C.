import re
from datetime import datetime

import pandas as pd
import streamlit as st

import core
import formato
import informe

st.set_page_config(page_title="Informe de Fragmentación · FAMESA", page_icon="📊", layout="wide")
core.vincular(st.session_state)   # datos solo en la memoria de esta sesión
core.init_db()
st.markdown("<h2 style='margin-bottom:0'>📊 Informe de Fragmentación <span style='color:#E10613'>· FAMESA</span></h2>",
            unsafe_allow_html=True)
st.caption("Sube tu Excel de WipFrag, completa los datos del informe y revisa todo en pantalla antes de descargar. "
           "🔒 Tus datos no se guardan: viven solo en esta sesión y se borran al cerrar o recargar la página.")


def ajuste(label, clave, default="", area=False, key=None):
    """Campo de texto que se recuerda entre sesiones."""
    f = st.text_area if area else st.text_input
    kw = {"height": 150} if area else {}
    val = f(label, value=core.get_ajuste(clave, default), key=key or clave, **kw)
    if val != core.get_ajuste(clave, default):
        core.set_ajuste(clave, val)
    return val


# ------------------------------------------------------------ barra lateral
with st.sidebar:
    st.header("1. Archivo WipFrag")
    f = st.file_uploader("Excel completo", type=["xlsx", "xls", "csv", "txt", "tsv"])
    if f is not None:
        clave = f"{f.name}-{f.size}"
        if st.session_state.get("archivo") != clave:       # importa una sola vez por archivo
            n_v, n_m = core.importar(core.leer_archivo(f))
            st.session_state["archivo"] = clave
            st.session_state["msg"] = f"{n_v} tajo(s) y {n_m} muestra(s) cargados."
    if st.session_state.get("msg"):
        st.success(st.session_state["msg"])
    if st.session_state.get("archivo") and st.button("🗑 Empezar de cero"):
        try:
            st.session_state["_db"].close()
        except Exception:
            pass
        st.session_state.clear()
        st.rerun()

vol = core.voladuras_df()
if vol.empty:
    st.info("Sube el Excel de WipFrag en la barra lateral para armar el informe.")
    st.stop()

with st.sidebar:
    st.header("2. Portada y pie de página")
    titulo = ajuste("Título de portada (hasta 3 líneas)", "titulo3", informe.DEF_TITULO, area=True)
    cliente = ajuste("Cliente / compañía", "cliente", "NOMBRE DE LA COMPAÑÍA")
    unidad = ajuste("Unidad minera", "unidad", "UNIDAD MINERA")
    periodo = ajuste("Periodo / semana", "periodo", "SEMANA " + datetime.now().strftime("%V, %Y"))
    dirigido = ajuste("Dirigido a", "dirigido", "Ing. NOMBRE APELLIDO (Cargo – CLIENTE)")
    copia = ajuste("Con copia a (una persona por línea, máx. 2)", "copia",
                   "Ing. NOMBRE APELLIDO (Cargo – FAMESA EXPLOSIVOS SAC)", area=True)
    pie = ajuste("Pie de página (lo que sigue a «INFORME TÉCNICO»)", "pie", "CLIENTE - UNDERGROUND")
    elaborado = ajuste("Elaborado por", "elaborado_por", "")
    firma = st.checkbox("Incluir tarjeta de firma al final", value=True, key="firma") if formato.FIRMA.exists() else False
    mina = st.number_input("P80 objetivo de MINA (in)", min_value=0.0, step=0.1, format="%.2f",
                           value=float(core.get_config("p80_objetivo_mina") or 0.0), key="mina") or None
    if mina != core.get_config("p80_objetivo_mina"):
        core.set_config("p80_objetivo_mina", mina)
    with st.expander("Textos editables (objetivo y metodología)"):
        objetivo_txt = ajuste("Objetivo", "objetivo_txt", informe.DEF_OBJETIVO, area=True)
        metodo_txt = ajuste("Metodología (un paso por línea)", "metodologia_txt", informe.DEF_METODOLOGIA, area=True)

    st.header("3. Tajos del informe")
    fechas = pd.to_datetime(vol["fecha"])
    rango = st.date_input("Rango de fechas", value=(fechas.min().date(), fechas.max().date()),
                          key=f"rango_{fechas.min().date()}_{fechas.max().date()}")
    desde, hasta = (rango[0], rango[1]) if isinstance(rango, (tuple, list)) and len(rango) == 2 else (
        fechas.min().date(), fechas.max().date())
    en_rango = vol[(fechas.dt.date >= desde) & (fechas.dt.date <= hasta)]
    opciones = {f"{r.labor} · {core.fdate(r.fecha)}": int(r.id) for r in en_rango.itertuples()}
    elegidos = st.multiselect("Tajos a incluir", list(opciones), default=list(opciones),
                              key=f"sel_{desde}_{hasta}_{len(opciones)}")
    ids = [opciones[e] for e in elegidos]

meta = {"cliente": cliente, "unidad": unidad, "titulo": titulo, "periodo": periodo, "dirigido": dirigido,
        "copia": copia, "pie": pie, "firma": firma, "elaborado_por": elaborado, "mina": mina, "objetivo_txt": objetivo_txt, "metodologia_txt": metodo_txt,
        "emision": datetime.now().strftime("%d/%m/%Y")}


CLAVES_CONTENIDO = ("cliente", "periodo", "elaborado_por", "mina", "objetivo_txt", "metodologia_txt", "emision")


def armar(ids, contenido):
    m = dict(contenido)
    return informe.construir(informe.preparar(list(ids), m.get("mina")), m)


def md(t):
    return re.sub(r"\[\[(.+?)\]\]", r":red[**\1**]", str(t))


def cel(c):
    return c[0] if isinstance(c, tuple) else c


tab1, tab2 = st.tabs(["📄 Vista previa del informe", "🗂 Base de datos de tajos"])

# ------------------------------------------------------------ vista previa
with tab1:
    if not ids:
        st.info("Selecciona al menos un tajo en la barra lateral.")
    else:
        contenido = tuple((k, meta.get(k)) for k in CLAVES_CONTENIDO)
        data = armar(tuple(ids), contenido)
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Preparar PDF", type="primary"):
                with st.spinner("Generando PDF..."):
                    st.session_state["pdf"] = formato.a_pdf(data, meta)
            if st.session_state.get("pdf"):
                st.download_button("⬇ Descargar PDF", st.session_state["pdf"], mime="application/pdf",
                                   file_name=f"Informe_Fragmentacion_{periodo}.pdf".replace(" ", "_").replace("/", "-"))
        with c2:
            if st.button("Preparar Word"):
                with st.spinner("Generando Word..."):
                    st.session_state["docx"] = formato.a_docx(data, meta)
            if st.session_state.get("docx"):
                st.download_button("⬇ Descargar Word", st.session_state["docx"],
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                   file_name=f"Informe_Fragmentacion_{periodo}.docx".replace(" ", "_").replace("/", "-"))
        st.caption("Si cambias algo en la barra lateral, vuelve a pulsar Preparar para actualizar el archivo. "
                   "Los gráficos aparecen a medida que se dibujan; la primera vez tarda, luego quedan guardados.")
        st.divider()

        st.markdown("### " + " ".join(meta["titulo"].upper().split()))
        st.table(pd.DataFrame(data["portada"], columns=["Dato", "Valor"]).set_index("Dato"))
        nfig = 0
        for b in data["blocks"]:
            t = b["t"]
            if t == "h1":
                st.header(b["text"])
            elif t == "h2":
                st.subheader(b["text"])
            elif t == "h3":
                st.markdown(f"**{b['text']}**")
            elif t == "p":
                st.markdown(md(b["text"]))
            elif t == "bullets":
                st.markdown("\n".join(f"- {md(i)}" for i in b["items"]))
            elif t == "eq":
                st.code(b["text"], language=None)
            elif t == "kv":
                st.markdown(f"**{b['title']}**")
                st.table(pd.DataFrame([(k, md(v)) for k, v in b["rows"]], columns=["Parámetro", "Valor"]).set_index("Parámetro"))
            elif t == "table":
                st.dataframe(pd.DataFrame([[re.sub(r"\*\*|\[\[|\]\]", "", str(cel(c))) for c in r] for r in b["rows"]],
                                          columns=b["head"]), hide_index=True)
                if b.get("note"):
                    st.caption(b["note"])
            elif t == "img":
                png = informe.png_de(b)
                if png:
                    nfig += 1
                    st.image(png, caption=f"IMAGEN {nfig}: {b['caption']}")
            elif t == "pb":
                st.divider()

# ------------------------------------------------------------ base de datos
with tab2:
    st.write("Edita los datos de cada tajo: cuerpo, tipo de roca, burden, tonelaje, explosivos y **P80 objetivo de la "
             "voladura**. El factor de potencia se calcula solo si lo dejas vacío.")
    ed = st.data_editor(core.voladuras_df(), hide_index=True, disabled=["id", "labor", "fecha"], key="editor")
    if st.button("Guardar cambios"):
        core.guardar_voladuras(ed)
        st.success("Guardado. La vista previa se actualiza sola.")
