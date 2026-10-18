"""
SIMULADOR HIDRÁULICO EN ESTADO ESTABLE - APLICACIÓN WEB (Streamlit)
Generación de perfiles hidráulicos para priorización ILI - multisistema y multiusuario.

Ejecutar en el computador:   streamlit run app.py
"""
import html
import io

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import motor as m

st.set_page_config(page_title="Simulador Hidráulico", page_icon="🛢️", layout="wide")

VERDE, AZUL, ROJO, NARANJA = "#D6DC3F", "#2C4E66", "#FF5B5B", "#F4B183"
st.markdown(f"""
<style>
.titulo {{border:2px solid {AZUL}; color:{AZUL}; text-align:center; font-weight:700; font-size:20px;
          padding:6px; margin-bottom:8px; background:white; font-family:Calibri, Arial}}
table.xl {{border-collapse:collapse; font-family:Calibri, Arial, sans-serif; font-size:14px; margin-bottom:6px}}
table.xl td, table.xl th {{border:1px solid #8ea0b2; padding:3px 9px; white-space:nowrap}}
table.xl th {{color:{AZUL}; background:#ffffff; text-align:center; font-weight:700; white-space:normal}}
table.xl td.in {{background:{VERDE}; color:{AZUL}; font-weight:700; text-align:right}}
table.xl td.out {{background:{AZUL}; color:{VERDE}; font-weight:700; text-align:right}}
table.xl td.lbl {{color:{AZUL}; font-weight:700; border:none; background:transparent}}
table.xl td.est {{color:#0563C1; text-decoration:underline; font-weight:600}}
table.xl td.rot {{background:{ROJO}; color:white; font-weight:700; text-align:right}}
table.xl td.des {{background:{NARANJA}; color:{AZUL}; font-weight:700; text-align:right}}
table.xl td.vacio {{border:none}}
.leyenda span {{padding:2px 10px; margin-right:6px; font-weight:700; border:1px solid #8ea0b2}}
</style>""", unsafe_allow_html=True)

AYUDA = {
    "flujo": "Caudal de operación en miles de barriles por día. Al cambiarlo se recalcula todo.",
    "bph": "Si conoce el caudal en barriles por hora, escríbalo aquí: el programa ajusta el KBDO "
           "(KBDO = bph × 24 / 1000).",
    "iny1": "Inyección (+) o entrega al paso (−) en bph. Se aplica en la estación con INY 1.",
    "iny2": "Inyección (+) o entrega al paso (−) en bph. Se aplica en la estación con INY 2.",
    "producto": "Densidad y viscosidad se toman de la tabla de productos (pestaña Productos = hoja REF).",
    "recibo": "Presión de llegada a la estación (psig). Condición de borde: el perfil se calcula de aguas abajo "
              "hacia aguas arriba.",
    "bypass": "Marque si la estación no bombea: su presión de recibo se iguala a la de descarga.",
    "dr": "% de reducción de fricción por DRA del tramo que inicia en la estación (solo flujo turbulento). "
          "Con %DR > 100 la fricción queda negativa.",
    "fr": "Factor de rotura (1/1.1 = 0.909). Hay rotura de tubo cuando MAOP < Presión int × FR.",
    "iny": "Indique si en esta estación entra INY 1, INY 2 o ninguna.",
    "descarga": "Presión requerida al inicio del tramo para llegar con la presión de recibo de la siguiente.",
    "roturas": "Número de puntos del tramo donde MAOP < Presión int × FR.",
    "remanente": "Mínima presión interna positiva del tramo (punto más crítico para descuelgue).",
    "lleno": "Lleno de línea del tramo en miles de barriles.",
}

ss = st.session_state


# ==========================================================================
# ESTADO DE LA SESIÓN
# ==========================================================================
def bd():
    if "db" not in ss:
        ss.db = m.BaseDatos(m.RUTA_BD)
    return ss.db


def sincronizar_widgets():
    """Pasa a los controles los valores del proyecto (al cargar o después de Buscar objetivo)."""
    p = ss.proy
    ss.w_flujo = float(p.flujo_kbd)
    ss.w_bph = round(p.flujo_kbd * m.KBD_A_BPH, 4)
    ss.w_iny1 = float(p.iny1_bph)
    ss.w_iny2 = float(p.iny2_bph)
    ss.w_prod = p.producto
    ss.w_fecha = p.fecha
    ss.w_resp = p.responsable
    ss.w_km = "; ".join(f"{k:g}" for k in p.km_interpolar)
    ss.ver = ss.get("ver", 0) + 1          # reinicia las tablas editables


def cargar_sistema(nombre):
    ss.proy = bd().cargar(nombre)
    info = bd().info(nombre)
    ss.sistema = nombre
    ss.cargado_ts = info[0] if info else None
    ss.cambios = []
    ss.modificado = False
    ss.sel_sistema = nombre
    ss.pop("obj_estado", None)
    sincronizar_widgets()


def marcar():
    ss.modificado = True


def inicializar():
    if "proy" in ss:
        return
    sistemas = bd().sistemas()
    ss.cambios, ss.modificado, ss.ver = [], False, 0
    if sistemas:
        ultimo = bd().config("ultimo")
        cargar_sistema(ultimo if ultimo in sistemas else sistemas[0])
    else:
        ss.proy = m.Proyecto()
        ss.proy.productos = bd().productos()
        ss.sistema = None


# ---------------------------------------------------------------- callbacks
def cb_flujo():
    ss.w_bph = round(ss.w_flujo * m.KBD_A_BPH, 4)
    ss.proy.flujo_kbd = float(ss.w_flujo)
    marcar()


def cb_bph():
    ss.w_flujo = ss.w_bph / m.KBD_A_BPH
    ss.proy.flujo_kbd = float(ss.w_flujo)
    marcar()


def cb_simple(attr, clave):
    setattr(ss.proy, attr, ss[clave])
    marcar()


def cb_producto():
    p = ss.proy
    p.producto = ss.w_prod
    for pr in p.productos:
        if pr[0] == p.producto:
            p.densidad, p.viscosidad = float(pr[1]), float(pr[2])
    marcar()


def cb_cambiar_sistema():
    nuevo = ss.sel_sistema
    if ss.get("modificado") and nuevo != ss.sistema:
        ss.sistema_pendiente = nuevo
        ss.sel_sistema = ss.sistema           # se queda en el actual hasta que el usuario confirme
    else:
        cargar_sistema(nuevo)
        bd().fijar_config("ultimo", nuevo)


# ==========================================================================
# BUSCAR OBJETIVO (ventanas emergentes iguales a Excel)
# ==========================================================================
@st.dialog("Buscar objetivo")
def dialogo_buscar_objetivo():
    p = ss.proy
    res = m.calcular_proyecto(p)
    salidas, entradas = m.celdas_salida(p), m.celdas_entrada(p)
    i_celda = st.selectbox("Definir la celda:", range(len(salidas)), format_func=lambda k: salidas[k][0])
    celda = salidas[i_celda][1]
    actual = m.evaluar_celda(p, celda, res)
    conteo = celda[0] in m.ES_CONTEO
    valor = st.number_input("Con el valor:", value=0.0 if conteo else round(float(actual), 2),
                            step=1.0 if conteo else 10.0, format="%.0f" if conteo else "%.2f",
                            key=f"obj_valor_{i_celda}")
    # sugerencia de la celda a cambiar (como haría un ingeniero)
    t, i = celda
    i = 0 if i is None else i
    claves = [e[1] for e in entradas]
    sug = ("dr", i)
    if conteo:
        j = i + 1
        while j < len(p.estaciones) and p.estaciones[j].paso_directo:
            j += 1
        sug = ("recibo", j) if ("recibo", j) in claves else sug
    i_var = st.selectbox("Para cambiar la celda:", range(len(entradas)), format_func=lambda k: entradas[k][0],
                         index=claves.index(sug) if sug in claves else 0, key=f"obj_var_{i_celda}")
    st.caption(f"Valor actual de la celda: {actual:,.2f}." + (
        "  Para Rotura de tubo / Descuelgue se busca el valor LÍMITE (el más cercano al actual) de la celda a "
        "cambiar con el que el conteo queda en el valor indicado." if conteo else
        "  Igual que en Excel: se cambia la celda indicada hasta que la celda definida tome el valor."))
    c1, c2 = st.columns(2)
    if c1.button("Aceptar", type="primary", use_container_width=True):
        var = entradas[i_var][1]
        with st.spinner("Buscando…"):
            r = m.buscar_objetivo(p, celda, valor, var)
        anterior = m.leer_variable(p, var)
        if r.encontrado:
            m.fijar_variable(p, var, r.valor_variable)
            marcar()
        ss.obj_estado = dict(celda=salidas[i_celda][0], valor=valor, var=var, nombre_var=entradas[i_var][0],
                             anterior=anterior, res=r)
        ss.sync = True
        st.rerun()
    if c2.button("Cancelar", use_container_width=True):
        st.rerun()


@st.dialog("Estado de la búsqueda de objetivo")
def dialogo_estado():
    e = ss.obj_estado
    r = e["res"]
    if r.encontrado:
        st.write(f"La búsqueda con la celda **{e['celda']}** ha encontrado una solución.")
    else:
        st.error(f"La búsqueda con la celda **{e['celda']}** NO ha encontrado una solución.")
    filas = [("Valor objetivo", f"{e['valor']:,.2f}"), ("Valor actual", f"{r.valor_obtenido:,.2f}"),
             (e["nombre_var"], f"{r.valor_variable:,.4f}  (antes {e['anterior']:,.4f})")]
    if e["var"][0] in ("flujo", "bph"):
        q = r.valor_variable if e["var"][0] == "flujo" else r.valor_variable / m.KBD_A_BPH
        filas.append(("Flujo resultante", f"{q:,.4f} KBDO = {q * m.KBD_A_BPH:,.1f} bph"))
    st.table(pd.DataFrame(filas, columns=["", "Valor"]).set_index(""))
    st.caption(r.mensaje + "  Aceptar conserva el valor en la hoja; Cancelar lo devuelve al anterior.")
    c1, c2 = st.columns(2)
    if c1.button("Aceptar", type="primary", use_container_width=True):
        ss.pop("obj_estado")
        st.rerun()
    if c2.button("Cancelar", use_container_width=True):
        if r.encontrado:
            m.fijar_variable(ss.proy, e["var"], e["anterior"])
            ss.sync = True
        ss.pop("obj_estado")
        st.rerun()


# ==========================================================================
# GRÁFICOS Y TABLAS
# ==========================================================================
def figura_perfil(p, res, alto=600):
    fig = go.Figure()
    ss._ley_rot = False
    for k, (t, r) in enumerate(zip(p.tramos, res)):
        km = [pt.km for pt in t.puntos]
        txt = [pt.punto for pt in t.puntos]
        primero = k == 0
        fig.add_trace(go.Scatter(x=km, y=r.J, name="Presión geodésica (psig)", legendgroup="g", showlegend=primero,
                                 line=dict(color="#7F6000", width=1.2), fill="tozeroy",
                                 fillcolor="rgba(217,201,163,0.35)", text=txt,
                                 hovertemplate="km %{x:.3f}<br>P. geodésica %{y:,.1f} psig<br>%{text}<extra></extra>"))
        fig.add_trace(go.Scatter(x=km, y=r.N, name="Presión int+geod (psig)", legendgroup="n", showlegend=primero,
                                 line=dict(color="#2E75B6", width=2.2), customdata=r.O,
                                 hovertemplate="km %{x:.3f}<br>P. int+geod %{y:,.1f}<br>P. int %{customdata:,.1f} psig"
                                               "<extra></extra>"))
        fig.add_trace(go.Scatter(x=km, y=r.Q, name="MAOP+geod (psig)", legendgroup="q", showlegend=primero,
                                 line=dict(color="#C00000", width=1, dash="dash"), customdata=r.P,
                                 hovertemplate="km %{x:.3f}<br>MAOP+geod %{y:,.1f}<br>MAOP %{customdata:,.1f} psig"
                                               "<extra></extra>"))
        rot = [j for j, s in enumerate(r.S) if s]
        if rot and k > 0 and any(any(x.S) for x in res[:k]):
            ss._ley_rot = True
        if rot:
            fig.add_trace(go.Scatter(x=[km[j] for j in rot], y=[r.N[j] for j in rot], mode="markers",
                                     name="Rotura de tubo", legendgroup="r", showlegend=not ss.get("_ley_rot"),
                                     marker=dict(symbol="x", size=10, color="red"),
                                     hovertemplate="ROTURA km %{x:.3f}<extra></extra>"))
        des = [j for j, s in enumerate(r.R) if s]
        if des:
            fig.add_trace(go.Scatter(x=[km[j] for j in des], y=[r.N[j] for j in des], mode="markers",
                                     name="Descuelgue", legendgroup="d", marker=dict(symbol="triangle-down",
                                                                                   size=9, color="orange")))
    kms = m.km_estaciones(p)
    for i, e in enumerate(p.estaciones):
        if i < len(res):
            y, txt = res[i].N[0], f"<b>{e.nombre}</b><br>{kms[i]:.2f} km<br>Desc: {res[i].presion_descarga:,.0f}"
        else:
            y, txt = res[-1].N[-1], f"<b>{e.nombre}</b><br>{kms[i]:.2f} km<br>Rec: {res[-1].presion_recibo:,.0f}"
        fig.add_vline(x=kms[i], line=dict(color="gray", width=1, dash="dot"))
        fig.add_annotation(x=kms[i], y=y, text=txt, showarrow=False, xanchor="left", yanchor="bottom",
                           font=dict(size=10, color=AZUL), xshift=4, yshift=4)
    fig.update_layout(height=alto, template="plotly_white", margin=dict(l=60, r=20, t=60, b=110),
                      title=dict(text=f"Perfil hidráulico - {p.sistema} | {p.producto} | {p.flujo_kbd:,.2f} KBDO",
                                 font=dict(color=AZUL, size=16), x=0.5, xanchor="center"),
                      xaxis_title="Distancia (km)", yaxis_title="Presión (psi)",
                      legend=dict(orientation="h", y=-0.15, x=0.5, xanchor="center", yanchor="top"))
    return fig


def celda(valor, clase, ayuda="", dec=2):
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        valor = f"{valor:,.{dec}f}"
    t = f' title="{html.escape(ayuda)}"' if ayuda else ""
    return f'<td class="{clase}"{t}>{html.escape(str(valor))}</td>'


def html_resumen(p, res):
    """Hoja Resumen con la misma estructura y colores de la plantilla Excel."""
    n = len(p.tramos)
    kms = m.km_estaciones(p)
    fl = m.flujos_por_tramo(p)
    h = ['<table class="xl">']
    h.append("<tr>" + celda("Flujo (KBDO)", "lbl") + celda(p.flujo_kbd, "in", AYUDA["flujo"], 4)
             + celda(p.flujo_kbd * m.KBD_A_BPH, "in", AYUDA["bph"]) + celda("bph", "lbl")
             + celda("INY 1 (bph)", "lbl") + celda(p.iny1_bph, "in", AYUDA["iny1"])
             + celda("INY 2 (bph)", "lbl") + celda(p.iny2_bph, "in", AYUDA["iny2"]) + "</tr>")
    h.append("<tr>" + celda("Productos", "lbl") + celda("Densidad Kg/m3", "lbl") + celda(p.densidad, "out")
             + "</tr>")
    h.append("<tr>" + celda(p.producto, "in", AYUDA["producto"]) + celda("Viscosidad cSt", "lbl")
             + celda(p.viscosidad, "out", dec=3) + "</tr></table>")
    h.append('<table class="xl"><tr><td class="vacio"></td>'
             + "".join(f"<th>{html.escape(e.nombre)}</th>" for e in p.estaciones[:n]) + "</tr>")
    h.append("<tr>" + celda("%DR", "lbl") + "".join(celda(e.dr, "in", AYUDA["dr"], 4) for e in p.estaciones[:n])
             + "</tr>")
    h.append("<tr>" + celda("FR", "lbl") + "".join(celda(e.fr, "in", AYUDA["fr"], 4) for e in p.estaciones[:n])
             + "</tr></table>")
    cols = ["Inicio Segmento", "Ubicación (km)", "Presión recibo (psig)", "Presión descarga (psig)", "Flujo (Kbdo)",
            "Rotura de tubo", "Lleno de línea (kbls)", "Tramo", "Presión Int remanente psig (Descuelgue)",
            "INYECCIÓN", "Bypass"]
    h.append('<table class="xl"><tr>' + "".join(f"<th>{c}</th>" for c in cols) + "</tr>")
    for i, e in enumerate(p.estaciones):
        r = res[i] if i < n else None
        fila = celda(e.nombre, "est") + celda(kms[i], "out")
        if i == 0:
            fila += celda("—", "out")
        else:
            fila += celda(e.presion_recibo, "out" if e.paso_directo else "in", AYUDA["recibo"])
        if r:
            fila += (celda(r.presion_descarga, "out", AYUDA["descarga"]) + celda(fl[i], "out")
                     + celda(r.roturas, "rot" if r.roturas else "out", AYUDA["roturas"], 0)
                     + celda(r.lleno_kbls, "out", AYUDA["lleno"]) + celda(p.tramos[i].nombre, "out")
                     + celda(r.presion_remanente or 0, "des" if r.descuelgues else "out", AYUDA["remanente"])
                     + celda(e.inyeccion, "in", AYUDA["iny"]))
        else:
            fila += celda("", "out") * 6 + celda("", "vacio")
        fila += celda("Sí" if e.paso_directo else ("No" if 0 < i < n else ""), "in" if 0 < i < n else "vacio",
                      AYUDA["bypass"])
        h.append(f"<tr>{fila}</tr>")
    h.append("<tr>" + celda("Total (km)", "lbl") + celda(kms[-1], "out") + '<td class="vacio" colspan="3"></td>'
             + celda("Line fill", "lbl") + celda(sum(r.lleno_kbls for r in res), "out")
             + '<td class="vacio"></td>' + celda("Roturas / Descuelgues", "lbl")
             + celda(f"{sum(r.roturas for r in res)} / {sum(r.descuelgues for r in res)}", "out") + "</tr></table>")
    return "".join(h)


# ==========================================================================
# PÁGINA
# ==========================================================================
inicializar()
if ss.pop("sync", False):
    sincronizar_widgets()
db = bd()
p = ss.proy

# ---------------------------------------------------------------- barra lateral
with st.sidebar:
    st.markdown(f"<h3 style='color:{AZUL};margin-bottom:0'>🛢️ Simulador hidráulico</h3>"
                "<span style='color:gray'>Estado estable · Priorización ILI</span>", unsafe_allow_html=True)
    st.text_input("Su nombre (para el historial)", key="usuario", placeholder="Nombre y apellido")
    sistemas = db.sistemas()
    if sistemas:
        if ss.get("sel_sistema") not in sistemas:
            ss.sel_sistema = ss.sistema if ss.sistema in sistemas else sistemas[0]
        st.selectbox("Sistema", sistemas, key="sel_sistema", on_change=cb_cambiar_sistema)
    if ss.get("sistema_pendiente"):
        st.warning(f"'{ss.sistema}' tiene cambios sin guardar. ¿Qué desea hacer antes de abrir "
                   f"'{ss.sistema_pendiente}'?")
        a, b = st.columns(2)
        if a.button("Descartar cambios"):
            nuevo = ss.pop("sistema_pendiente")
            cargar_sistema(nuevo)
            st.rerun()
        if b.button("Seguir aquí"):
            ss.pop("sistema_pendiente")
            st.rerun()

    if ss.sistema:
        info = db.info(ss.sistema)
        if info:
            st.caption(f"Guardado en la base de datos: {info[0]} por {info[1]}")
        if ss.modificado:
            st.warning("Hay cambios sin guardar.")
        nota = st.text_input("Descripción del cambio (historial)", key="nota_guardar",
                             placeholder="Ej.: actualización espesores ILI 2026")
        otro = info and ss.cargado_ts and info[0] != ss.cargado_ts
        forzar = True
        if otro:
            st.error(f"Otro usuario ({info[1]}) guardó este sistema el {info[0]} después de que usted lo abrió.")
            forzar = st.checkbox("Guardar de todas formas (sobrescribe lo de él)")
        c1, c2 = st.columns(2)
        if c1.button("💾 Guardar", type="primary", use_container_width=True, disabled=not ss.modificado):
            if not ss.get("usuario", "").strip():
                st.error("Escriba su nombre antes de guardar.")
            elif forzar:
                n = db.guardar(p, nota, ss.cambios, usuario=ss.usuario.strip())
                ss.cambios, ss.modificado = [], False
                ss.cargado_ts = db.info(ss.sistema)[0]
                st.toast(f"Sistema guardado ({n} cambios registrados en el historial).", icon="✅")
                st.rerun()
        if c2.button("↩️ Descartar", use_container_width=True, disabled=not ss.modificado):
            cargar_sistema(ss.sistema)
            st.rerun()
        if st.button("🎯 Buscar objetivo…", use_container_width=True):
            dialogo_buscar_objetivo()

    with st.expander("📥 Importar plantillas Excel a la base de datos"):
        archivos = st.file_uploader("Plantillas del simulador (.xlsm / .xlsx)", type=["xlsm", "xlsx"],
                                    accept_multiple_files=True)
        reemplazar = st.checkbox("Reemplazar si el sistema ya existe")
        if archivos and st.button("Importar", use_container_width=True):
            if not ss.get("usuario", "").strip():
                st.error("Escriba su nombre arriba antes de importar.")
            else:
                ok, msgs = [], []
                barra = st.progress(0.0)
                for k, f in enumerate(archivos, 1):
                    try:
                        q = m.importar_plantilla_excel(f.getvalue())
                        if not q.tramos:
                            raise ValueError("no se encontraron tramos. " + " ".join(q.avisos_importacion))
                        for av in q.avisos_importacion:
                            msgs.append(f"⚠️ {f.name}: {av}")
                        nombre = q.sistema.strip() or f.name.rsplit(".", 1)[0]
                        q.sistema = nombre
                        if db.existe(nombre) and not reemplazar:
                            msgs.append(f"⏭️ {nombre}: ya existía (no se reemplazó)")
                        else:
                            db.guardar(q, f"Importado de {f.name}", usuario=ss.usuario.strip())
                            ok.append(nombre)
                            msgs.append(f"✅ {nombre}")
                    except Exception as ex:  # noqa: BLE001
                        msgs.append(f"❌ {f.name}: {ex}")
                    barra.progress(k / len(archivos))
                st.write("\n\n".join(msgs))
                if ok and not ss.sistema:
                    cargar_sistema(ok[0])
                    st.rerun()

    with st.expander("💾 Copia de la base de datos"):
        st.caption("Descargue simuladores.db para tener una copia. En Streamlit Cloud los cambios se pierden "
                   "cuando la aplicación se reinicia: suba esta copia a GitHub para conservarlos.")
        try:
            with open(db.ruta, "rb") as fh:
                st.download_button("⬇️ Descargar simuladores.db", fh.read(), file_name="simuladores.db",
                                   use_container_width=True)
        except OSError:
            st.error("No se pudo leer la base de datos.")

    if ss.sistema:
        with st.expander("🗂️ Administrar sistema"):
            nuevo_nombre = st.text_input("Guardar como nuevo sistema / renombrar", value=ss.sistema)
            a, b = st.columns(2)
            if a.button("Guardar como", use_container_width=True):
                if db.existe(nuevo_nombre):
                    st.error("Ya existe ese nombre.")
                elif ss.get("usuario", "").strip():
                    p.sistema = nuevo_nombre
                    db.guardar(p, f"Copia de {ss.sistema}", ss.cambios, usuario=ss.usuario.strip())
                    cargar_sistema(nuevo_nombre)
                    st.rerun()
                else:
                    st.error("Escriba su nombre arriba.")
            if b.button("Renombrar", use_container_width=True):
                if db.existe(nuevo_nombre):
                    st.error("Ya existe ese nombre.")
                else:
                    db.renombrar(ss.sistema, nuevo_nombre)
                    cargar_sistema(nuevo_nombre)
                    st.rerun()
            if st.checkbox(f"Quiero eliminar '{ss.sistema}' de la base de datos"):
                if st.button("Eliminar definitivamente", type="secondary"):
                    db.eliminar(ss.sistema)
                    for k in ("proy", "sistema"):
                        ss.pop(k, None)
                    st.rerun()

if not ss.sistema:
    st.markdown('<div class="titulo">SIMULADOR EN ESTADO ESTABLE PARA LA GENERACIÓN DE PERFILES HIDRÁULICOS '
                'PARA PRIORIZACIÓN ILI</div>', unsafe_allow_html=True)
    st.info("La base de datos no tiene sistemas. Use **Importar plantillas Excel** en la barra lateral.")
    st.stop()

# ---------------------------------------------------------------- cálculo
try:
    res = m.calcular_proyecto(p)
except Exception as ex:  # noqa: BLE001
    st.error(f"Error en el cálculo: {ex}")
    st.stop()
for i, e in enumerate(p.estaciones):                 # bypass: recibo = descarga
    if e.paso_directo and i < len(res):
        e.presion_recibo = res[i].presion_descarga
if ss.get("obj_estado"):
    dialogo_estado()

st.markdown('<div class="titulo">FORMATO<br>SIMULADOR EN ESTADO ESTABLE PARA LA GENERACIÓN DE PERFILES '
            'HIDRÁULICOS PARA PRIORIZACIÓN ILI</div>', unsafe_allow_html=True)
c = st.columns([3, 2, 2])
c[0].markdown(f"**Sistema:** <span style='color:{AZUL};font-size:18px;font-weight:700'>{html.escape(p.sistema)}"
              "</span>", unsafe_allow_html=True)
c[1].text_input("Fecha de elaboración", key="w_fecha", on_change=cb_simple, args=("fecha", "w_fecha"))
c[2].text_input("Responsable elaboración", key="w_resp", on_change=cb_simple, args=("responsable", "w_resp"))

tabs = st.tabs(["📋 Resumen", "📏 Tramos", "🔧 Actualizar espesores", "📍 Interpolación", "📈 Curva del sistema",
                "🧪 Productos (REF)", "🕓 Historial", "❓ Ayuda"])

# ================================================================ RESUMEN
with tabs[0]:
    st.markdown('<div class="leyenda"><span style="background:#D6DC3F;color:#2C4E66">Celda modificable</span>'
                '<span style="background:#2C4E66;color:#D6DC3F">Celda calculada</span>'
                '<span style="color:gray;border:none">Pase el mouse sobre una celda o un título para ver su ayuda.'
                '</span></div>', unsafe_allow_html=True)
    st.markdown("##### ✏️ Datos de operación")
    a = st.columns(5)
    a[0].number_input("Flujo (KBDO)", key="w_flujo", min_value=0.0, step=1.0, format="%.4f", on_change=cb_flujo,
                      help=AYUDA["flujo"])
    a[1].number_input("Flujo (bph)", key="w_bph", min_value=0.0, step=50.0, format="%.2f", on_change=cb_bph,
                      help=AYUDA["bph"])
    a[2].number_input("INY 1 (bph)", key="w_iny1", step=50.0, format="%.2f", on_change=cb_simple,
                      args=("iny1_bph", "w_iny1"), help=AYUDA["iny1"])
    a[3].number_input("INY 2 (bph)", key="w_iny2", step=50.0, format="%.2f", on_change=cb_simple,
                      args=("iny2_bph", "w_iny2"), help=AYUDA["iny2"])
    nombres_prod = [x[0] for x in p.productos]
    if ss.w_prod not in nombres_prod:
        nombres_prod.append(ss.w_prod)
    a[4].selectbox("Producto", nombres_prod, key="w_prod", on_change=cb_producto, help=AYUDA["producto"])

    st.markdown("##### ✏️ Datos de estaciones")
    n = len(p.tramos)
    df_est = pd.DataFrame({
        "Estación": [e.nombre for e in p.estaciones],
        "Presión recibo (psig)": [None if i == 0 else e.presion_recibo for i, e in enumerate(p.estaciones)],
        "Bypass": [bool(e.paso_directo) for e in p.estaciones],
        "%DR": [e.dr if i < n else None for i, e in enumerate(p.estaciones)],
        "FR": [e.fr if i < n else None for i, e in enumerate(p.estaciones)],
        "Inyección": [e.inyeccion if i < n else None for i, e in enumerate(p.estaciones)],
    })
    ed = st.data_editor(
        df_est, key=f"ed_est_{ss.ver}", hide_index=True, use_container_width=True, num_rows="fixed",
        column_config={
            "Estación": st.column_config.TextColumn(disabled=True),
            "Presión recibo (psig)": st.column_config.NumberColumn(help=AYUDA["recibo"] + " (Con Bypass se calcula.)",
                                                                   format="%.2f"),
            "Bypass": st.column_config.CheckboxColumn(help=AYUDA["bypass"]),
            "%DR": st.column_config.NumberColumn(help=AYUDA["dr"], format="%.4f"),
            "FR": st.column_config.NumberColumn(help=AYUDA["fr"], format="%.4f"),
            "Inyección": st.column_config.SelectboxColumn(options=list(m.INYECCIONES), help=AYUDA["iny"]),
        })
    cambio = False
    for i, e in enumerate(p.estaciones):
        fila = ed.iloc[i]
        bp = bool(fila["Bypass"]) and 0 < i < n
        if bp != e.paso_directo:
            e.paso_directo, cambio = bp, True
        v = fila["Presión recibo (psig)"]
        if i > 0 and not e.paso_directo and pd.notna(v) and float(v) != e.presion_recibo:
            e.presion_recibo, cambio = float(v), True
        if i < n:
            for col, attr in (("%DR", "dr"), ("FR", "fr")):
                v = fila[col]
                if pd.notna(v) and float(v) != getattr(e, attr):
                    setattr(e, attr, float(v))
                    cambio = True
            if fila["Inyección"] in m.INYECCIONES and fila["Inyección"] != e.inyeccion:
                e.inyeccion, cambio = fila["Inyección"], True
    if cambio:
        marcar()
        ss.ver += 1
        st.rerun()

    st.markdown("##### 📋 Hoja Resumen")
    st.markdown(html_resumen(p, res), unsafe_allow_html=True)
    rot = sum(r.roturas for r in res)
    if rot:
        st.error(f"⚠️ {rot} puntos con rotura de tubo. " + ", ".join(
            f"{p.estaciones[i].nombre}: {r.roturas}" for i, r in enumerate(res) if r.roturas))
    st.plotly_chart(figura_perfil(p, res), use_container_width=True, config={"displaylogo": False})
    if st.button("📄 Preparar resultados en Excel"):
        xl = io.BytesIO()
        m.exportar_resultados_excel(p, res, xl)
        ss.excel = xl.getvalue()
    if ss.get("excel"):
        st.download_button("⬇️ Descargar resultados en Excel", ss.excel, file_name=f"resultados_{p.sistema}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           on_click=lambda: ss.pop("excel", None))

# ================================================================ TRAMOS
COLS_EDIT = {"Punto/ubicación": "punto", "dist.del reg. [km]": "km", "altura [m]": "altura", "esp. [in]": "espesor",
             "diám. ext. [in]": "diam_ext", "SMYS [PSI]": "smys", "Rating de Válvula (psi)": "rating"}
with tabs[1]:
    nombres_tr = [f"Tramo {t.nombre}: {p.estaciones[i].nombre} → {p.estaciones[i + 1].nombre}"
                  for i, t in enumerate(p.tramos)]
    it = st.radio("Tramo", range(len(p.tramos)), format_func=lambda k: nombres_tr[k], horizontal=True,
                  label_visibility="collapsed")
    t, r = p.tramos[it], res[it]
    e0 = p.estaciones[it]
    q = m.flujos_por_tramo(p)[it]
    filas = [("Caudal (kBDO)", f"{q:,.2f}"), ("Caudal (m3/s)", f"{q * m.KBD_A_M3S:.4f}"),
             ("bph", f"{q * m.KBD_A_BPH:,.2f}"), ("Densidad (kg/m3)", f"{p.densidad:,.2f}"),
             ("Viscosidad (cSt)", f"{p.viscosidad:.3f}"), ("ºAPI", f"{141.5e3 / p.densidad - 131.5:.2f}"),
             ("Presión Descarga (psig)", f"{r.presion_descarga:,.2f}"),
             ("Presión de Entrada (psig)", f"{r.presion_recibo:,.2f}"), ("% Reducción Fricción", f"{e0.dr:.4f}"),
             ("Lleno de Línea (bls)", f"{r.T[-1]:,.2f}"), ("Factor rotura", f"{e0.fr:.4f}"),
             ("Rotura tubo", str(r.roturas)), ("Descuelgue", str(r.descuelgues))]
    st.markdown('<table class="xl"><tr>' + "".join(f"<th>{a}</th>" for a, _ in filas) + "</tr><tr>" + "".join(
        celda(b, "rot" if (a == "Rotura tubo" and r.roturas) else "out") for a, b in filas) + "</tr></table>",
                unsafe_allow_html=True)
    o = st.columns([1, 1, 1.4, 1, 3])
    rug = o[0].number_input("✏️ Rugosidad interna (in)", value=float(t.rugosidad_in), format="%.5f", min_value=0.0,
                            key=f"rug_{it}_{ss.ver}")
    eff = o[1].number_input("✏️ Eficiencia de la tubería (%)", value=float(t.eficiencia), min_value=1.0,
                            max_value=200.0, key=f"eff_{it}_{ss.ver}")
    if rug != t.rugosidad_in or eff != t.eficiencia:
        t.rugosidad_in, t.eficiencia = rug, eff
        marcar()
        st.rerun()
    o[2].markdown("<br>", unsafe_allow_html=True)
    solo = o[2].toggle("Ver solo filas con rotura / descuelgue")
    csv_perfil = io.StringIO()
    pd.DataFrame([[x.punto, x.km, x.altura, x.espesor, x.diam_ext, x.smys, x.rating] for x in t.puntos],
                 columns=m.COLUMNAS_CSV).to_csv(csv_perfil, sep=";", index=False)
    o[3].markdown("<br>", unsafe_allow_html=True)
    o[3].download_button("⬇️ Perfil (CSV)", csv_perfil.getvalue().encode("utf-8-sig"),
                         file_name=f"perfil_{p.sistema}_tramo_{t.nombre}.csv")
    der = st.container()
    with der:
        df = pd.DataFrame({
            "Punto/ubicación": [x.punto for x in t.puntos], "dist.del reg. [km]": [x.km for x in t.puntos],
            "altura [m]": [x.altura for x in t.puntos], "esp. [in]": [x.espesor for x in t.puntos],
            "diám. ext. [in]": [x.diam_ext for x in t.puntos], "SMYS [PSI]": [x.smys for x in t.puntos],
            "Presión geodésica (psig)": r.J, "Velocidad [m/s]": r.K, "Nre": r.L, "Pérdidas fricción (psi/km)": r.M,
            "Presión int+geod (psig)": r.N, "Presión int (psig)": r.O, "MAOP (psig)": r.P, "MAOP+geod (psig)": r.Q,
            f"Descuelgue: {r.descuelgues}": r.R, f"Rotura tubo: {r.roturas}": r.S, "Lleno de línea": r.T,
            "Rating de Válvula (psi)": [x.rating or None for x in t.puntos]})
        df.index = range(3, len(df) + 3)
        if solo:
            vista = df[df[f"Rotura tubo: {r.roturas}"] | df[f"Descuelgue: {r.descuelgues}"]]
            st.dataframe(vista.style.format(precision=2).apply(
                lambda s: [f"background-color:{ROJO};color:white" if v else "" for v in s],
                subset=[f"Rotura tubo: {r.roturas}"]), use_container_width=True, height=600)
        else:
            st.caption("✏️ Columnas editables: Punto, km, altura, espesor, diámetro, SMYS y Rating de válvula "
                       "(puede pegar desde Excel con Ctrl+V; agregue o borre filas al final de la tabla). "
                       "Las demás columnas son calculadas.")
            cfg = {c: st.column_config.NumberColumn(disabled=True, format="%.2f") for c in df.columns
                   if c not in COLS_EDIT}
            cfg.update({c: st.column_config.CheckboxColumn(disabled=True) for c in df.columns
                        if c.startswith(("Descuelgue", "Rotura"))})
            cfg["Velocidad [m/s]"] = st.column_config.NumberColumn(disabled=True, format="%.4f")
            cfg["dist.del reg. [km]"] = st.column_config.NumberColumn("✏️ km", format="%.6f", required=True)
            cfg["esp. [in]"] = st.column_config.NumberColumn("✏️ esp. [in]", format="%.4f", min_value=0.01)
            cfg["altura [m]"] = st.column_config.NumberColumn("✏️ altura [m]", format="%.3f")
            cfg["diám. ext. [in]"] = st.column_config.NumberColumn("✏️ diám. ext. [in]", format="%.3f",
                                                                    min_value=0.1)
            cfg["SMYS [PSI]"] = st.column_config.NumberColumn("✏️ SMYS [PSI]", format="%.0f")
            cfg["Rating de Válvula (psi)"] = st.column_config.NumberColumn("✏️ Rating válvula (psi)", format="%.0f",
                                                                            help="Vacío = MAOP por espesor")
            cfg["Punto/ubicación"] = st.column_config.TextColumn("✏️ Punto/ubicación")
            clave = f"ed_tr_{it}_{ss.ver}"
            ed_t = st.data_editor(df, key=clave, column_config=cfg, num_rows="dynamic", height=640,
                                  use_container_width=True)
            est = ss.get(clave, {})
            if est.get("edited_rows") or est.get("added_rows") or est.get("deleted_rows"):
                nuevos = []
                for _, f in ed_t.iterrows():
                    if pd.isna(f["dist.del reg. [km]"]):
                        continue

                    def num(col, defecto=0.0):
                        v = f[col]
                        return defecto if pd.isna(v) else float(v)
                    nuevos.append(m.PuntoPerfil("" if pd.isna(f["Punto/ubicación"]) else str(f["Punto/ubicación"]),
                                                num("dist.del reg. [km]"), num("altura [m]"), num("esp. [in]"),
                                                num("diám. ext. [in]"), num("SMYS [PSI]"),
                                                num("Rating de Válvula (psi)")))
                if len(nuevos) >= 2:
                    t.puntos = nuevos
                    marcar()
                    ss.ver += 1
                    st.rerun()
                else:
                    st.error("El tramo debe tener al menos 2 puntos con km.")

# ================================================================ ESPESORES
with tabs[2]:
    st.markdown("Asigne el espesor medido en una nueva inspección a un rango de km. Si los km no existen en el "
                "perfil se crean los puntos (altura interpolada) y el cambio de espesor en escalón, igual que en "
                "la plantilla. Los cambios quedan pendientes hasta **Guardar** y se registran en el Historial.")
    with st.form("espesor"):
        f1, f2, f3, f4 = st.columns([3, 1, 1, 1])
        itr = f1.selectbox("Tramo", range(len(p.tramos)), format_func=lambda k: (
            f"{nombres_tr[k]} (km {p.tramos[k].puntos[0].km:.3f} a {p.tramos[k].puntos[-1].km:.3f})"))
        desde = f2.number_input("km desde", format="%.6f")
        hasta = f3.number_input("km hasta", format="%.6f")
        esp = f4.number_input("Nuevo espesor (in)", min_value=0.0, format="%.4f")
        nota_e = st.text_input("Fuente / nota", placeholder="Ej.: ILI MFL 2026")
        if st.form_submit_button("Aplicar al rango", type="primary"):
            if esp <= 0 or hasta <= desde:
                st.error("Revise el rango de km y el espesor.")
            else:
                cambios = m.actualizar_espesor_rango(p.tramos[itr], desde, hasta, esp)
                for km, ant, nuevo in cambios:
                    ss.cambios.append((p.tramos[itr].nombre, km, "Espesor (in)", f"{ant:.10g}", f"{nuevo:.10g}",
                                       nota_e or f"Espesor km {desde:g}-{hasta:g}"))
                if cambios:
                    marcar()
                    ss.ver += 1
                st.success(f"{len(cambios)} filas modificadas en el tramo {p.tramos[itr].nombre}.")
    st.markdown("**Importar muchos rangos** — archivo CSV o Excel con columnas: `km_desde ; km_hasta ; espesor_in ; "
                "nota` (el tramo se identifica por el km).")
    arch = st.file_uploader("Archivo de rangos", type=["csv", "txt", "xlsx", "xlsm"], key=f"rangos_{ss.ver}")
    if arch and st.button("Aplicar archivo"):
        total, log = 0, []
        for a_, b_, e_, nt in m.leer_rangos_inspeccion((arch.name, arch.getvalue())):
            for i, tr in enumerate(p.tramos):
                ini, fin = max(min(a_, b_), tr.puntos[0].km), min(max(a_, b_), tr.puntos[-1].km)
                if fin > ini:
                    cambios = m.actualizar_espesor_rango(tr, ini, fin, e_)
                    for km, ant, nuevo in cambios:
                        ss.cambios.append((tr.nombre, km, "Espesor (in)", f"{ant:.10g}", f"{nuevo:.10g}",
                                           nt or f"Archivo {arch.name}"))
                    total += len(cambios)
                    log.append(f"Tramo {tr.nombre}: km {ini:g}–{fin:g} → {e_:g} in ({len(cambios)} filas)")
        if total:
            marcar()
            ss.ver += 1
        st.success(f"{total} filas modificadas.\n\n" + "\n\n".join(log))
    if ss.cambios:
        st.markdown("**Cambios pendientes de guardar**")
        st.dataframe(pd.DataFrame(ss.cambios, columns=["Tramo", "km", "Campo", "Anterior", "Nuevo", "Nota"]),
                     use_container_width=True, hide_index=True)

# ================================================================ INTERPOLACIÓN
with tabs[3]:
    st.text_input("Distancias del registro a interpolar (km, separadas por ; o espacio)", key="w_km")
    try:
        p.km_interpolar = [float(x.replace(",", ".")) for x in ss.w_km.replace(";", " ").split()]
    except ValueError:
        st.error("Lista de km no válida.")
    filas = []
    for km, d in zip(p.km_interpolar, m.interpolar_lista(p, res, p.km_interpolar)):
        if d:
            filas.append([d["km"], d["dist_tramo"], d["altura"], d["p_geod"], d["maop"], d["p_int"], d["maop_rem"],
                          d["p_int_geod"], d["velocidad"], d["nre"], d["tramo"], d["inicio"], d["fin"], d["fila"]])
        else:
            filas.append([km] + [None] * 9 + ["Fuera de rango", "", "", None])
    dfi = pd.DataFrame(filas, columns=["dist.del reg. [km]", "dist. tramo [km]", "altura [m]",
                                       "Presión geodésica (psig)", "MAOP (psig)", "Presión int (psig)",
                                       "MAOP remanente (psig)", "Presión int+geod (psig)", "Velocidad [m/s]", "Nre",
                                       "Tramo", "Inicio tramo", "Fin tramo", "# Fila"])
    st.dataframe(dfi.style.format(precision=2, na_rep=""), use_container_width=True, hide_index=True)

# ================================================================ CURVA DEL SISTEMA
with tabs[4]:
    k = st.columns(4)
    li = k[0].number_input("Flujo mínimo (KBDO)", value=0.0, min_value=0.0)
    ls = k[1].number_input("Flujo máximo (KBDO)", value=round(p.flujo_kbd * 1.25, 1), min_value=0.1)
    pasos = k[2].number_input("Puntos", value=10, min_value=2, max_value=50)
    modo = k[3].selectbox("%DR en la curva", ["Sin DRA (%DR = 0)", "%DR actual de cada tramo"])
    st.caption("Con %DR mayor que 100 la fricción queda negativa y la presión de descarga baja al aumentar el "
               "caudal; por eso la curva se calcula por defecto sin DRA.")
    curva = m.curva_sistema(p, li, ls, int(pasos), 0.0 if modo.startswith("Sin") else None)
    nombres_e = [e.nombre for e in p.estaciones[:len(p.tramos)]]
    dfc = pd.DataFrame([[q] + ps for q, ps in curva], columns=["KBDO"] + nombres_e)
    figc = go.Figure([go.Scatter(x=dfc["KBDO"], y=dfc[nm], name=nm, mode="lines+markers") for nm in nombres_e])
    figc.add_vline(x=p.flujo_kbd, line=dict(dash="dash", color="gray"), annotation_text="Operación")
    figc.update_layout(template="plotly_white", height=480, xaxis_title="Flujo (KBDO)",
                       yaxis_title="Presión de descarga (psig)", title=f"Curva del sistema - {p.sistema} ({modo})")
    st.plotly_chart(figc, use_container_width=True, config={"displaylogo": False})
    st.dataframe(dfc.style.format(precision=2), use_container_width=True, hide_index=True)

# ================================================================ PRODUCTOS
with tabs[5]:
    st.caption("Tabla común a todos los sistemas (hoja REF). Se guarda con 💾 Guardar.")
    dfp = pd.DataFrame(p.productos, columns=["Producto", "Densidad Kg/m3", "Viscosidad cSt 1", "Temp 1 °F",
                                             "Densidad2 Kg/m3", "Viscosidad2 cSt", "Temp 2 °F"])
    edp = st.data_editor(dfp, key=f"ed_prod_{ss.ver}", num_rows="dynamic", use_container_width=True, hide_index=True)
    if ss.get(f"ed_prod_{ss.ver}", {}).get("edited_rows") or ss.get(f"ed_prod_{ss.ver}", {}).get("added_rows") \
            or ss.get(f"ed_prod_{ss.ver}", {}).get("deleted_rows"):
        p.productos = [[None if pd.isna(x) else (x if j == 0 else float(x)) for j, x in enumerate(fila)]
                       for fila in edp.itertuples(index=False) if not pd.isna(fila[0])]
        cb_producto() if ss.w_prod in [x[0] for x in p.productos] else marcar()
        ss.ver += 1
        st.rerun()

# ================================================================ HISTORIAL
with tabs[6]:
    hist = pd.DataFrame(db.historial(p.sistema), columns=["Fecha", "Usuario", "Tramo / estación", "km", "Campo",
                                                          "Valor anterior", "Valor nuevo", "Nota"])
    st.dataframe(hist, use_container_width=True, hide_index=True, height=600)
    st.download_button("⬇️ Exportar historial (CSV)", hist.to_csv(sep=";", index=False).encode("utf-8-sig"),
                       file_name=f"historial_{p.sistema}.csv")

# ================================================================ AYUDA
with tabs[7]:
    st.markdown("""
**Cómo usar el simulador**

1. Escriba **su nombre** en la barra lateral (queda en el historial) y escoja el **Sistema**.
2. **Resumen**: cambie Flujo (en KBDO o en bph), INY 1/2, Producto y la tabla de estaciones (presión de recibo,
   Bypass, %DR, FR, Inyección). Todo se recalcula al instante. Debajo están la hoja Resumen con los colores de la
   plantilla y el perfil hidráulico (puede hacer zoom y pasar el mouse sobre las curvas).
3. **🎯 Buscar objetivo** (barra lateral): igual que en Excel — *Definir la celda*, *Con el valor*, *Para cambiar
   la celda*. Ejemplos: Presión descarga SEBASTOPOL = 1900 cambiando %DR; cambiando Flujo (bph) para ajustar el
   KBDO; Rotura de tubo = 0 cambiando la Presión de recibo de la estación siguiente (máxima presión sin roturas).
4. **Tramos**: perfil completo de cada tramo. Las columnas ✏️ se pueden editar o pegar desde Excel.
5. **Actualizar espesores**: por rango de km o con un archivo de rangos de una inspección.
6. **💾 Guardar**: guarda en la base de datos compartida y registra los cambios en el **Historial**. Si otro
   usuario guardó el mismo sistema mientras usted trabajaba, el programa le avisa.
7. **Importar plantillas Excel** (barra lateral): suba una o varias plantillas .xlsm para agregar sistemas.
""")
