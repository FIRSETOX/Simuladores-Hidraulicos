"""
SIMULADOR HIDRÁULICO EN ESTADO ESTABLE - APLICACIÓN WEB (Streamlit)
Generación de perfiles hidráulicos para priorización ILI - multisistema y multiusuario.

Ejecutar en el computador:   streamlit run app.py
"""
import dataclasses
import html
import io
import math

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import importlib

import motor as m

# Funciones del motor que necesita esta versión de app.py. Si el servidor tiene cargado un motor.py viejo
# (Streamlit no siempre recarga los módulos importados al actualizar el repositorio), se recarga aquí.
REQUERIDOS_MOTOR = ("presion_minima_recibo", "tolerancia_perfil", "verificar_perfil", "leer_perfil_presion",
                    "buscar_objetivo", "importar_plantilla_excel", "simular_fuga", "calcular_red", "proyecto_ruta",
                    "rutas_red", "Ramal", "posiciones_esquema", "subred_ramal",
                    "nombres_curva", "flujo_operacion", "curva_entre_estaciones",
                    "flujo_en_estacion")


def _motor_incompleto():
    faltan = [f for f in REQUERIDOS_MOTOR if not hasattr(m, f)]
    campos_est = {f.name for f in dataclasses.fields(m.Estacion)}
    campos_proy = {f.name for f in dataclasses.fields(m.Proyecto)}
    faltan += [c for c in ("delta_bph", "clientes") if c not in campos_est]
    faltan += [c for c in ("perfiles_externos", "margen_descuelgue_psi", "ramales") if c not in campos_proy]
    return faltan


if _motor_incompleto():
    m = importlib.reload(m)

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
    "p_min_recibo": "Presión de recibo MÍNIMA para que el tramo que llega a esta estación no se descuelgue: la "
                    "presión convertida a cabeza (m) alcanza justo para pasar el punto más alto (presión interna = "
                    "margen en el punto crítico). Depende del caudal, %DR, producto y perfil del tramo.",
    "holgura": "Presión de recibo actual − presión mínima sin descuelgue. Negativa (naranja) = el tramo se "
               "descuelga (columna separada) con la presión actual.",
    "margen_desc": "Presión interna mínima que se exige en el punto crítico. 0 = la cabeza de presión es igual a la "
                   "altura del punto (límite teórico). Puede usar, p.ej., la presión de vapor del producto + un margen.",
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


VERSION_CODIGO = "2026-10-20c"


def motor_desactualizado():
    """True si motor.py en el servidor es más viejo que app.py (no se subió el motor.py nuevo)."""
    return bool(_motor_incompleto())


def inicializar():
    if motor_desactualizado():
        st.error("⚠️ El archivo **motor.py** del servidor está desactualizado respecto a **app.py** (le falta: "
                 + ", ".join(_motor_incompleto()) + "). Suba a GitHub el **motor.py** nuevo junto con app.py, "
                 "espere 1-2 minutos y recargue la página. Si sigue igual: *Manage app → Reboot app*.")
        st.stop()
    # objetos creados con una versión anterior del código (sesión abierta durante una actualización):
    # se vuelve a cargar el sistema desde la base de datos para que tengan todos los campos nuevos.
    if "proy" in ss and ss.get("version_codigo") != VERSION_CODIGO:
        sist = ss.get("sistema")
        datos = None
        if ss.get("modificado"):
            try:
                datos = m.Proyecto.from_dict(m.asdict(ss.proy))   # conserva los cambios sin guardar
            except Exception:  # noqa: BLE001
                datos = None
        for k in list(ss.keys()):
            del ss[k]
        ss.version_codigo = VERSION_CODIGO
        ss.cambios, ss.modificado, ss.ver = [], False, 0
        sistemas = bd().sistemas()
        if sist in sistemas or sistemas:
            cargar_sistema(sist if sist in sistemas else sistemas[0])
        if datos is not None and ss.get("sistema") == sist:
            datos.productos = ss.proy.productos
            ss.proy, ss.modificado = datos, True
        if "proy" in ss:
            return
    ss.version_codigo = VERSION_CODIGO
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
COLORES_PERFILES = ["#2CA02C", "#FF7F0E", "#9467BD", "#8C564B", "#E377C2", "#17BECF", "#BCBD22"]
CURVAS = {"geod": "Presión geodésica", "hid": "Presión int+geod", "maop": "MAOP+geod",
          "pint": "Presión interna (perfil MOP del simulador)", "barlow": "Presión de diseño Barlow / MAOP",
          "perfiles": "Perfiles cargados (MOP / transitorios)", "minima": "Línea mínima sin descuelgue"}


def figura_perfil(p, res, curvas=("geod", "hid", "maop"), verifs=None, alto=620, titulo=None):
    """Perfil hidráulico. curvas: claves de CURVAS. verifs: verificaciones de perfiles cargados."""
    fig = go.Figure()
    total = sum(len(t.puntos) for t in p.tramos)
    Sc = go.Scattergl if total > 4000 else go.Scatter
    ya = set()
    minimos = m.presion_minima_recibo(p, res) if "minima" in curvas else []
    tols = sorted({v["tolerancia_pct"] for v in (verifs or {}).values() if v["tolerancia_pct"] > 0}) \
        if "perfiles" in curvas else []

    def ley(g):
        if g in ya:
            return False
        ya.add(g)
        return True

    for k, (t, r) in enumerate(zip(p.tramos, res)):
        km = [pt.km for pt in t.puntos]
        txt = [pt.punto for pt in t.puntos]
        if "geod" in curvas:
            fig.add_trace(Sc(x=km, y=r.J, name="Presión geodésica (psig)", legendgroup="g", showlegend=ley("g"),
                             line=dict(color="#7F6000", width=1.2), fill="tozeroy", fillcolor="rgba(217,201,163,0.35)",
                             text=txt, hovertemplate="km %{x:.3f}<br>P. geodésica %{y:,.1f} psig<br>%{text}"
                                                     "<extra></extra>"))
        if "hid" in curvas:
            fig.add_trace(Sc(x=km, y=r.N, name="Presión int+geod (psig)", legendgroup="n", showlegend=ley("n"),
                             line=dict(color="#2E75B6", width=2.2), customdata=r.O,
                             hovertemplate="km %{x:.3f}<br>P. int+geod %{y:,.1f}<br>P. int %{customdata:,.1f} psig"
                                           "<extra></extra>"))
        if "maop" in curvas:
            fig.add_trace(Sc(x=km, y=r.Q, name="MAOP+geod (psig)", legendgroup="q", showlegend=ley("q"),
                             line=dict(color="#C00000", width=1, dash="dash"), customdata=r.P,
                             hovertemplate="km %{x:.3f}<br>MAOP+geod %{y:,.1f}<br>MAOP %{customdata:,.1f} psig"
                                           "<extra></extra>"))
        if "pint" in curvas:
            fig.add_trace(Sc(x=km, y=r.O, name="Presión interna - simulador (psig)", legendgroup="o",
                             showlegend=ley("o"), line=dict(color="#0B3C8C", width=2, dash="dot"), text=txt,
                             hovertemplate="km %{x:.3f}<br>P. interna %{y:,.1f} psig<br>%{text}<extra></extra>"))
        if "barlow" in curvas:
            fig.add_trace(Sc(x=km, y=r.P, name="Presión de diseño Barlow / MAOP (psig)", legendgroup="b",
                             showlegend=ley("b"), line=dict(color="#FF0000", width=1.6), text=txt,
                             hovertemplate="km %{x:.3f}<br>P. diseño %{y:,.1f} psig<br>%{text}<extra></extra>"))
            for tol in tols:
                fig.add_trace(Sc(x=km, y=[v * (1 + tol / 100) for v in r.P], legendgroup=f"bt{tol:g}",
                                 name=f"Presión de diseño + {tol:g} % (transitorios)", showlegend=ley(f"bt{tol:g}"),
                                 line=dict(color="#FF0000", width=1, dash="dot"),
                                 hovertemplate=f"km %{{x:.3f}}<br>P. diseño + {tol:g} % %{{y:,.1f}} psig"
                                               "<extra></extra>"))
        if minimos and minimos[k].get("p_min") is not None:
            d = minimos[k]
            interna = "hid" not in curvas
            ym = [y - j for y, j in zip(d["linea_minima"], r.J)] if interna else d["linea_minima"]
            fig.add_trace(Sc(x=km, y=ym, name="Línea mínima sin descuelgue (psig)", legendgroup="mn",
                             showlegend=ley("mn"), line=dict(color="#00B050", width=1.6, dash="dash"),
                             hovertemplate="km %{x:.3f}<br>Línea mínima %{y:,.1f} psig<extra></extra>"))
            yc = d["margen_psi"] + (0 if interna else d["altura_critica"] * p.densidad * m.PSI_POR_M_KGM3)
            fig.add_trace(go.Scatter(
                x=[d["km_critico"]], y=[yc], mode="markers+text", name="Punto crítico de descuelgue",
                legendgroup="pc", showlegend=ley("pc"), text=[f"P. mín. recibo {d['p_min']:,.0f}"],
                textposition="bottom right", textfont=dict(size=10, color="#00703C"),
                marker=dict(symbol="triangle-up", size=13, color="#00B050", line=dict(color="black", width=1)),
                hovertemplate=(f"PUNTO CRÍTICO DE DESCUELGUE<br>km {d['km_critico']:.3f}  cota "
                               f"{d['altura_critica']:,.1f} m<br>{html.escape(str(d['punto_critico'] or ''))}"
                               f"<br>P. mín. recibo {d['estacion']}: {d['p_min']:,.2f} psig "
                               f"({d['p_min_m']:,.1f} m)<extra></extra>")))
        rot = [j for j, s in enumerate(r.S) if s]
        if rot and ("hid" in curvas or "pint" in curvas):
            ycur = r.N if "hid" in curvas else r.O
            fig.add_trace(go.Scatter(x=[km[j] for j in rot], y=[ycur[j] for j in rot], mode="markers",
                                     name="Rotura de tubo", legendgroup="r", showlegend=ley("r"),
                                     marker=dict(symbol="x", size=10, color="red"),
                                     hovertemplate="ROTURA km %{x:.3f}<extra></extra>"))
        des = [j for j, s in enumerate(r.R) if s]
        if des and "hid" in curvas:
            fig.add_trace(go.Scatter(x=[km[j] for j in des], y=[r.N[j] for j in des], mode="markers",
                                     name="Descuelgue", legendgroup="d", showlegend=ley("d"),
                                     marker=dict(symbol="triangle-down", size=9, color="orange")))
    if "perfiles" in curvas:
        for j, per in enumerate(p.perfiles_externos):
            color = COLORES_PERFILES[j % len(COLORES_PERFILES)]
            fig.add_trace(Sc(x=per["km"], y=per["p"], name=f"{per['nombre']} ({per['tipo']})",
                             line=dict(color=color, width=2, dash="solid" if per["tipo"] == "MOP" else "dashdot"),
                             hovertemplate=f"{per['nombre']}<br>km %{{x:.3f}}<br>P. interna %{{y:,.1f}} psig"
                                           "<extra></extra>"))
            v = (verifs or {}).get(per["nombre"])
            if v and v["excesos"]:
                tol = v["tolerancia_pct"]
                for fuera, col, etq in ((False, "orange", f"dentro del {tol:g} %"), (True, "red", "fuera de tolerancia")):
                    ex = [z for z in v["excesos"] if (z[5] > tol + 1e-9) == fuera]
                    if not ex:
                        continue
                    fig.add_trace(go.Scatter(
                        x=[z[0] for z in ex], y=[z[1] for z in ex], mode="markers",
                        name=f"{per['nombre']}: supera diseño ({etq})",
                        marker=dict(color=col, size=8, symbol="circle-open", line=dict(width=2)),
                        customdata=[[z[2], z[4], z[5], z[3]] for z in ex],
                        hovertemplate=("SUPERA DISEÑO<br>km %{x:.3f}<br>P perfil %{y:,.1f} psig"
                                       "<br>P diseño %{customdata[0]:,.1f} psig<br>Exceso %{customdata[1]:,.1f} psi"
                                       " = %{customdata[2]:.2f} %<br>%{customdata[3]}<extra></extra>")))
    kms = m.km_estaciones(p)
    for i, e in enumerate(p.estaciones):
        base = res[i] if i < len(res) else res[-1]
        if "hid" in curvas:
            y = base.N[0] if i < len(res) else base.N[-1]
        else:
            y = base.O[0] if i < len(res) else base.O[-1]
        txt = (f"<b>{e.nombre}</b><br>{kms[i]:.2f} km<br>Desc: {res[i].presion_descarga:,.0f}" if i < len(res)
               else f"<b>{e.nombre}</b><br>{kms[i]:.2f} km<br>Rec: {res[-1].presion_recibo:,.0f}")
        fig.add_vline(x=kms[i], line=dict(color="gray", width=1, dash="dot"))
        fig.add_annotation(x=kms[i], y=y, text=txt, showarrow=False, xanchor="left", yanchor="bottom",
                           font=dict(size=10, color=AZUL), xshift=4, yshift=4)
    fig.update_layout(height=alto + 40, template="plotly_white", margin=dict(l=60, r=20, t=60, b=170),
                      title=dict(text=titulo or f"Perfil hidráulico - {p.sistema} | {p.producto} | "
                                                f"{p.flujo_kbd:,.2f} KBDO",
                                 font=dict(color=AZUL, size=16), x=0.5, xanchor="center"),
                      xaxis_title="Distancia (km)", yaxis_title="Presión (psi)", hovermode="closest",
                      legend=dict(orientation="h", y=-0.2, x=0.5, xanchor="center", yanchor="top"))
    return fig


def seccion_descuelgue(p, res):
    """Presión mínima de recibo en cada estación para que la línea no se descuelgue."""
    mins = m.presion_minima_recibo(p, res)
    validos = [d for d in mins if d.get("p_min") is not None]
    hay = any(d["descuelga"] for d in validos)
    with st.expander("🔻 Presión mínima de recibo para evitar descuelgue" + (" — ⚠️ HAY DESCUELGUE" if hay else ""),
                     expanded=hay):
        st.caption("Con el caudal, %DR y producto actuales: menor presión de recibo con la que la presión convertida "
                   "a cabeza (m) alcanza para pasar el punto más alto del tramo que llega a la estación. "
                   "P mín = máx[ P geodésica(i) + margen − pérdidas por fricción de i al recibo ] − P geodésica(recibo).")
        c1, c2 = st.columns([1, 3])
        marg = c1.number_input("Margen contra descuelgue (psi)", min_value=0.0, step=5.0, format="%.2f",
                               value=float(getattr(p, "margen_descuelgue_psi", 0.0) or 0.0),
                               key=f"w_margen_{ss.sistema}_{ss.ver}", help=AYUDA["margen_desc"])
        if abs(marg - float(getattr(p, "margen_descuelgue_psi", 0.0) or 0.0)) > 1e-12:
            p.margen_descuelgue_psi = marg
            marcar()
            ss.ver += 1
            st.rerun()
        filas = []
        for d in validos:
            filas.append({
                "Estación (recibo)": d["estacion"] + (" (bypass)" if d["bypass"] else ""),
                "P. mín. recibo (psig)": d["p_min"], "P. mín. recibo (m)": d["p_min_m"],
                "P. recibo actual (psig)": d["p_actual"], "Holgura (psi)": d["holgura"],
                "Estado": (f"⚠️ Se descuelga ({d['n_descuelgue']} puntos, km {d['km_descuelgue'][0]:.3f}–"
                           f"{d['km_descuelgue'][1]:.3f})" if d["km_descuelgue"] else "⚠️ Se descuelga")
                if d["descuelga"] else "✅ Sin descuelgue",
                "km punto crítico": d["km_critico"], "Punto crítico": d["punto_critico"] or "",
                "Cota crítica (m)": d["altura_critica"],
                "Descarga necesaria con el recibo mínimo (psig)": d["descarga_con_minimo"]})
        if filas:
            st.dataframe(pd.DataFrame(filas).style.format(
                {"P. mín. recibo (psig)": "{:,.2f}", "P. mín. recibo (m)": "{:,.1f}", "P. recibo actual (psig)": "{:,.2f}",
                 "Holgura (psi)": "{:,.2f}", "km punto crítico": "{:.3f}", "Cota crítica (m)": "{:,.1f}",
                 "Descarga necesaria con el recibo mínimo (psig)": "{:,.1f}"}),
                use_container_width=True, hide_index=True)
        aplicables = [(i + 1, d) for i, d in enumerate(mins) if d.get("p_min") is not None and not d["bypass"]]
        if aplicables:
            sel = c2.multiselect("Fijar la presión de recibo en su mínimo (redondeado hacia arriba a 0,1 psi)",
                                 [d["estacion"] for _, d in aplicables], key=f"apl_min_{ss.sistema}_{ss.ver}")
            if sel and c2.button("Aplicar presión mínima de recibo"):
                for k, d in aplicables:
                    if d["estacion"] in sel:
                        p.estaciones[k].presion_recibo = math.ceil(max(d["p_min"], 0.0) * 10 - 1e-9) / 10
                marcar()
                ss.ver += 1
                st.rerun()
        st.caption("También puede usar 🎯 Buscar objetivo: *Descuelgue – estación = 0* cambiando la presión de recibo "
                   "(da el mismo valor) o cambiando el flujo (caudal mínimo sin descuelgue).")


def mensaje_verif(nom, v):
    """Mensaje de verificación de un perfil: máximo % de exceso, km y tolerancia."""
    if v["n_puntos"] == 0:
        st.warning(f"⚠️ {nom}: el perfil no coincide con el rango de km del sistema.")
    elif v["cumple"]:
        st.success(f"✅ {nom}: no supera la presión de diseño (margen mínimo {-v['max_exceso_psi']:,.1f} psi = "
                   f"{-v['max_exceso_pct']:.2f} % en km {v['km_max_exceso']:.3f}).")
    else:
        txt = (f"{nom}: supera la presión de diseño en {v['n_excesos']} puntos; máximo **{v['max_exceso_pct']:.2f} %** "
               f"({v['max_exceso_psi']:,.1f} psi) en **km {v['km_max_exceso']:.3f}** — tolerancia {v['tolerancia_pct']:g} %")
        if v["aceptable"]:
            st.warning(f"⚠️ {txt}: DENTRO de lo permitido.")
        else:
            st.error(f"❌ {txt}: **FUERA** ({v['n_fuera']} puntos superan el {v['tolerancia_pct']:g} %).")


def verificaciones(p, res):
    return {per["nombre"]: m.verificar_perfil(p, res, per) for per in p.perfiles_externos}


def celda(valor, clase, ayuda="", dec=2):
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        valor = f"{valor:,.{dec}f}"
    t = f' title="{html.escape(ayuda)}"' if ayuda else ""
    return f'<td class="{clase}"{t}>{html.escape(str(valor))}</td>'


def html_resumen(p, res):
    """Hoja Resumen con la misma estructura y colores de la plantilla Excel."""
    n = len(p.tramos)
    kms = m.km_estaciones(p)
    fl = [r.flujo_kbd for r in res]
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
            "INYECCIÓN", "Bypass", "P. mín. recibo sin descuelgue (psig)", "Holgura recibo (psi)"]
    mins = m.presion_minima_recibo(p, res)
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
        d = mins[i - 1] if 0 < i <= len(mins) else {}
        if d.get("p_min") is not None:
            fila += (celda(d["p_min"], "out", AYUDA["p_min_recibo"])
                     + celda(d["holgura"], "des" if d["descuelga"] else "out", AYUDA["holgura"]))
        else:
            fila += celda("", "vacio") * 2
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

# ==========================================================================
# RED / RAMALES  y  FUGAS: funciones de la interfaz
# ==========================================================================
COLORES_RAMAL = ["#2E75B6", "#C55A11", "#548235", "#7030A0", "#BF8F00", "#00838F", "#AD1457"]


def nombre_cadena(p, cad):
    return "Línea principal" if not cad else f"Ramal {cad}"


def opciones_nodos(p):
    """Estaciones donde puede salir un ramal: [(cadena, índice, etiqueta)]."""
    ops = [("", i, f"Línea principal – {e.nombre}") for i, e in enumerate(p.estaciones)]
    for r in p.ramales:
        ops += [(r.nombre, i, f"Ramal {r.nombre} – {e.nombre}") for i, e in enumerate(r.estaciones) if i > 0]
    return ops


def posiciones_red(p):
    return m.posiciones_esquema(p)


def figura_esquema(p, res, nodo_sel=None, fuga=None):
    """Esquema de la red (estaciones, tramos, ramales, % de flujo y clientes)."""
    pos = posiciones_red(p)
    red = getattr(res, "red", None) or {}
    fig = go.Figure()
    cadenas = [("", p, list(res), 0)] + [(r.nombre, r, list(red[r.nombre]["res"]) if r.nombre in red else [], k + 1)
                                          for k, r in enumerate(p.ramales)]
    xs_all, ys_all = [0.0], [0.0]
    marcadores = []
    for nom, cad, rs, k in cadenas:
        color = AZUL if not nom else COLORES_RAMAL[(k - 1) % len(COLORES_RAMAL)]
        for i, t in enumerate(cad.tramos):
            a, b = pos.get((nom, i)), pos.get((nom, i + 1))
            if a is None or b is None:
                continue
            xs, ys = ([a[0], a[0], b[0]], [a[1], b[1], b[1]]) if (nom and i == 0 and a[1] != b[1]) else \
                ([a[0], b[0]], [a[1], b[1]])
            q = rs[i].flujo_kbd if i < len(rs) else None
            fig.add_trace(go.Scatter(x=xs, y=ys, mode="lines", line=dict(color=color, width=4),
                                     hoverinfo="text", showlegend=False,
                                     text=f"Tramo {t.nombre}" + (f"<br>{q:,.2f} KBDO" if q is not None else "")))
            fig.add_annotation(x=b[0] - 0.12, y=b[1], ax=xs[-2], ay=ys[-2], xref="x", yref="y", axref="x", ayref="y",
                               showarrow=True, arrowhead=3, arrowsize=1.4, arrowwidth=2, arrowcolor=color, text="")
            xm = (xs[-2] + xs[-1]) / 2
            fig.add_annotation(x=xm, y=b[1] + 0.17, showarrow=False, font=dict(size=12, color=color),
                               text=f"<b>{html.escape(str(t.nombre))}</b>" + (f"<br>{q:,.1f} KBDO" if q is not None else ""))
        if nom and nom in red:
            inf = red[nom]
            a = pos.get((nom, 0))
            b = pos.get((nom, 1))
            if a and b:
                etiqueta = f"{inf['pct_calculado']:.1f} %" + (" (hidráulico)" if cad.modo == "Hidráulico" else "")
                fig.add_annotation(x=a[0] + 0.05, y=(a[1] + b[1]) / 2, showarrow=False, xanchor="left",
                                   font=dict(size=12, color=color), text=f"<b>{etiqueta}</b>",
                                   bgcolor="rgba(255,255,255,0.85)")
        xs, ys, txt, cd, simb = [], [], [], [], []
        for i, e in enumerate(cad.estaciones):
            if nom and i == 0:
                continue                                       # el nodo ya está dibujado por su padre
            x, y = pos.get((nom, i), (None, None))
            if x is None:
                continue
            xs.append(x)
            ys.append(y)
            xs_all.append(x)
            ys_all.append(y)
            sale = [r.nombre for r in p.ramales if (r.padre or "") == nom and r.nodo == i]
            info = f"<b>{html.escape(e.nombre)}</b><br>{nombre_cadena(p, nom)}"
            if i > 0 and i - 1 < len(rs):
                info += f"<br>Recibo {rs[i - 1].presion_recibo:,.1f} psig"
            if i < len(rs):
                info += f"<br>Descarga {rs[i].presion_descarga:,.1f} psig"
            for rn in sale:
                if rn in red:
                    info += f"<br>→ {html.escape(rn)}: {red[rn]['flujo']:,.2f} KBDO, P nodo {red[rn]['p_nodo'] or 0:,.1f}"
            txt.append(info)
            cd.append(f"{nom}|{i}")
            simb.append("circle-x" if (e.paso_directo and 0 < i < len(cad.tramos)) else "circle")
            # clientes (entregas)
            clientes = [c.strip() for c in (getattr(e, "clientes", "") or "").split(",") if c.strip()]
            arriba = (y == 0.0)                     # línea principal: clientes arriba; ramales: abajo
            for k2, cli in enumerate(clientes):
                cx, cy = x + 0.35, (y + 0.42 + 0.22 * k2) if arriba else (y - 0.32 - 0.22 * k2)
                fig.add_trace(go.Scatter(x=[x, x, cx], y=[y, cy, cy], mode="lines", showlegend=False,
                                         line=dict(color="#2E75B6", width=1.5, dash="dot"), hoverinfo="skip"))
                fig.add_trace(go.Scatter(x=[cx + 0.06], y=[cy], mode="markers+text", showlegend=False,
                                         marker=dict(symbol="pentagon", size=16, color="#FFE600",
                                                     line=dict(color="#2E75B6", width=1.5)),
                                         text=[f"  {html.escape(cli)}"], textposition="middle right",
                                         textfont=dict(size=11, color="#1F3864"), hoverinfo="skip"))
                xs_all.append(cx + 1.0)
                ys_all.append(cy)
        tamanos = [26 if nodo_sel == c else 20 for c in cd]
        marcadores.append(go.Scatter(
            x=xs, y=ys, mode="markers+text", text=[t.split("<br>")[0] for t in txt], textposition="bottom center",
            textfont=dict(size=12, color="#1F3864"), hovertext=txt, hoverinfo="text", customdata=cd,
            marker=dict(symbol=simb, size=tamanos, color="#1F3864",
                        line=dict(color=["#FF0000" if nodo_sel == c else "white" for c in cd], width=3)),
            selected=dict(marker=dict(opacity=1)), unselected=dict(marker=dict(opacity=1)),
            showlegend=False))
    for tr_ in marcadores:                      # estaciones encima de las líneas (para poder hacer clic)
        fig.add_trace(tr_)
    fig.update_traces(unselected=dict(marker=dict(opacity=1), textfont=dict(color="#1F3864")),
                      selected=dict(marker=dict(opacity=1)), selector=dict(type="scatter"))
    if fuga:
        fig.add_trace(go.Scatter(x=[fuga[0]], y=[fuga[1]], mode="markers+text", showlegend=False,
                                 marker=dict(symbol="star", size=18, color="red"), text=["Fuga"],
                                 textposition="top center"))
    alto = 230 + 120 * max(0, -min(ys_all)) + 110 * max(0, max(ys_all))
    fig.update_layout(height=alto, template="plotly_white", margin=dict(l=10, r=10, t=10, b=10),
                      xaxis=dict(visible=False, range=[min(xs_all) - 0.5, max(xs_all) + 0.9]),
                      yaxis=dict(visible=False, range=[min(ys_all) - 0.7, max(ys_all) + 0.6]), dragmode=False,
                      clickmode="event+select")
    return fig


def cargar_otro_sistema(origen):
    """origen: ('excel', archivo subido) o ('bd', nombre)."""
    if origen[0] == "excel":
        return m.importar_plantilla_excel(origen[1].getvalue())
    return bd().cargar(origen[1])


def desplazar_km(ramal, delta):
    for t in ramal.tramos:
        for pt in t.puntos:
            pt.km += delta


def km_nodo(p, cadena, nodo):
    cad = m.cadena_de(p, cadena)
    if cad is None or not cad.tramos:
        return 0.0
    if nodo == 0:
        return cad.tramos[0].puntos[0].km if cad.tramos[0].puntos else 0.0
    t = cad.tramos[min(nodo, len(cad.tramos)) - 1]
    return max((pt.km for pt in t.puntos), default=0.0)


def tabla_estaciones_cadena(p, nom_cadena, cad, rs):
    """Resultados por estación de una cadena (línea principal o ramal)."""
    pr = p if not nom_cadena else m.proyecto_ramal(p, cad, rs[0].flujo_kbd if rs else 0.0)
    kms = m.km_estaciones(pr)
    try:
        mins = m.presion_minima_recibo(pr, rs)
    except Exception:  # noqa: BLE001
        mins = []
    filas = []
    for i, e in enumerate(cad.estaciones):
        r = rs[i] if i < len(rs) else None
        d = mins[i - 1] if 0 < i <= len(mins) else {}
        filas.append({"Estación": e.nombre, "km": kms[i] if i < len(kms) else None,
                      "Presión recibo (psig)": rs[i - 1].presion_recibo if 0 < i <= len(rs) else None,
                      "Presión descarga (psig)": r.presion_descarga if r else None,
                      "Flujo (KBDO)": r.flujo_kbd if r else None, "Rotura de tubo": r.roturas if r else None,
                      "Descuelgue": r.descuelgues if r else None, "Lleno de línea (kbls)": r.lleno_kbls if r else None,
                      "P. mín. recibo sin descuelgue (psig)": d.get("p_min"), "Holgura recibo (psi)": d.get("holgura"),
                      "Bypass": "Sí" if e.paso_directo else ""})
    return pd.DataFrame(filas)


def editor_perfil_tramo(t, r, clave):
    """Editor de los puntos de un tramo (pegar desde Excel). Devuelve True si cambió."""
    df = pd.DataFrame({
        "Punto/ubicación": [x.punto for x in t.puntos], "km": [x.km for x in t.puntos],
        "altura [m]": [x.altura for x in t.puntos], "esp. [in]": [x.espesor for x in t.puntos],
        "diám. ext. [in]": [x.diam_ext for x in t.puntos], "SMYS [PSI]": [x.smys for x in t.puntos],
        "Rating válvula (psi)": [x.rating or None for x in t.puntos]},
        columns=["Punto/ubicación", "km", "altura [m]", "esp. [in]", "diám. ext. [in]", "SMYS [PSI]",
                 "Rating válvula (psi)"])
    if r is not None and r.N and len(r.N) == len(t.puntos):
        df["Presión int (psig)"] = r.O
        df["MAOP (psig)"] = r.P
        df["Rotura"] = r.S
    cfg = {"km": st.column_config.NumberColumn(format="%.6f", required=True),
           "esp. [in]": st.column_config.NumberColumn(format="%.4f", min_value=0.0),
           "altura [m]": st.column_config.NumberColumn(format="%.3f"),
           "diám. ext. [in]": st.column_config.NumberColumn(format="%.3f", min_value=0.0),
           "SMYS [PSI]": st.column_config.NumberColumn(format="%.0f"),
           "Rating válvula (psi)": st.column_config.NumberColumn(format="%.0f", help="Vacío = MAOP por espesor"),
           "Presión int (psig)": st.column_config.NumberColumn(disabled=True, format="%.1f"),
           "MAOP (psig)": st.column_config.NumberColumn(disabled=True, format="%.1f"),
           "Rotura": st.column_config.CheckboxColumn(disabled=True)}
    ed = st.data_editor(df, key=clave, column_config=cfg, num_rows="dynamic", height=420, use_container_width=True)
    est = ss.get(clave, {})
    if not (est.get("edited_rows") or est.get("added_rows") or est.get("deleted_rows")):
        return False
    nuevos = []
    for _, f in ed.iterrows():
        if pd.isna(f["km"]):
            continue

        def num(col, defecto=0.0):
            v = f[col]
            return defecto if pd.isna(v) else float(v)
        nuevos.append(m.PuntoPerfil("" if pd.isna(f["Punto/ubicación"]) else str(f["Punto/ubicación"]),
                                    num("km"), num("altura [m]"), num("esp. [in]"), num("diám. ext. [in]"),
                                    num("SMYS [PSI]"), num("Rating válvula (psi)")))
    nuevos.sort(key=lambda x: x.km)
    if len(nuevos) >= 2 or not nuevos:
        t.puntos = nuevos
        return True
    st.error("El tramo debe tener al menos 2 puntos con km.")
    return False


def figura_fuga(f, curvas=("geod", "hid", "maop")):
    """Perfil con la fuga (línea azul) y sin la fuga (gris punteado)."""
    pf, rf = f["proy_fuga"], f["res_fuga"]
    p = ss.proy
    if f["cadena"]:
        cam = next((c for n_, c in m.rutas_red(pf) if c and c[-1].nombre == f["cadena"]), [])
        pr, rr = m.proyecto_ruta(pf, rf, cam)
        camb = next((c for n_, c in m.rutas_red(p) if c and c[-1].nombre == f["cadena"]), [])
        pb, rb = m.proyecto_ruta(p, f["res_base"], camb)
    else:
        pr, rr = pf, list(rf)
        pb, rb = p, list(f["res_base"])
    fig = figura_perfil(pr, rr, curvas, titulo=f"Perfil con fuga en km {f['km']:.3f} – {f['q_bph']:,.0f} bph")
    for k, (t, r) in enumerate(zip(pb.tramos, rb)):
        fig.add_trace(go.Scatter(x=[x.km for x in t.puntos], y=r.N, mode="lines", name="Presión int+geod SIN fuga",
                                 legendgroup="sf", showlegend=(k == 0),
                                 line=dict(color="gray", width=1.5, dash="dot"), hoverinfo="skip"))
    # punto de la fuga
    alt = None
    for t, r in zip(pr.tramos, rr):
        for x, n_ in zip(t.puntos, r.N):
            if abs(x.km - f["km"]) < 1e-6:
                alt = n_
                break
    fig.add_trace(go.Scatter(x=[f["km"]], y=[alt], mode="markers+text", name="Fuga",
                             marker=dict(symbol="star", size=18, color="red"), text=[f"Fuga {f['q_bph']:,.0f} bph"],
                             textposition="top center"))
    return fig


_tabs = st.tabs(["📋 Resumen", "🔀 Red / ramales", "💧 Fugas / válvula ilícita", "📉 Perfiles MOP / transitorios",
                 "📏 Tramos", "🔧 Actualizar espesores", "📍 Interpolación", "📈 Curva del sistema",
                 "🧪 Productos (REF)", "🕓 Historial", "❓ Ayuda"])
tabs = [_tabs[0]] + list(_tabs[3:])
tab_red, tab_fuga = _tabs[1], _tabs[2]

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
        "Iny(+)/Entrega(−) bph": [getattr(e, "delta_bph", 0.0) if 0 < i < n else None for i, e in enumerate(p.estaciones)],
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
            "Iny(+)/Entrega(−) bph": st.column_config.NumberColumn(
                help="Caudal que entra (+) o se entrega al paso (−) en esta estación, en bph. Cambia el flujo de "
                     "los tramos aguas abajo.", format="%.2f"),
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
            v = fila["Iny(+)/Entrega(−) bph"]
            if i > 0 and pd.notna(v) and float(v) != getattr(e, "delta_bph", 0.0):
                e.delta_bph, cambio = float(v), True
    if cambio:
        marcar()
        ss.ver += 1
        st.rerun()

    st.markdown("##### 📋 Hoja Resumen")
    if getattr(res, "red", None):
        st.info("🔀 Este sistema tiene ramales: " + "; ".join(
            f"**{nom}** se lleva {inf['flujo']:,.2f} KBDO ({inf['pct_calculado']:.1f} %) en "
            f"{m.cadena_de(p, inf['ramal'].padre).estaciones[inf['ramal'].nodo].nombre}"
            for nom, inf in res.red.items()) + ". La tabla muestra la línea principal; vea la pestaña *Red / ramales*.")
    st.markdown(html_resumen(p, res), unsafe_allow_html=True)
    rot = sum(r.roturas for r in res)
    if rot:
        st.error(f"⚠️ {rot} puntos con rotura de tubo. " + ", ".join(
            f"{p.estaciones[i].nombre}: {r.roturas}" for i, r in enumerate(res) if r.roturas))
    seccion_descuelgue(p, res)
    verifs = verificaciones(p, res)
    opciones = ["geod", "hid", "maop", "pint", "barlow", "minima"] + (["perfiles"] if p.perfiles_externos else [])
    defecto = ["geod", "hid", "maop"] + (["barlow", "perfiles"] if p.perfiles_externos else [])
    curvas = st.multiselect("Curvas del gráfico", opciones, default=defecto, format_func=lambda k: CURVAS[k],
                            key=f"curvas_{ss.sistema}_{len(p.perfiles_externos)}",
                            help="Agregue la Presión interna (perfil MOP) y la Presión de diseño Barlow, o los "
                                 "perfiles cargados en la pestaña 'Perfiles MOP / transitorios'.")
    for nom, v in verifs.items():
        mensaje_verif(nom, v)
    st.plotly_chart(figura_perfil(p, res, curvas, verifs), use_container_width=True, config={"displaylogo": False})
    if st.button("📄 Preparar resultados en Excel"):
        xl = io.BytesIO()
        m.exportar_resultados_excel(p, res, xl)
        ss.excel = xl.getvalue()
    if ss.get("excel"):
        st.download_button("⬇️ Descargar resultados en Excel", ss.excel, file_name=f"resultados_{p.sistema}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           on_click=lambda: ss.pop("excel", None))

# ================================================================ RED / RAMALES
with tab_red:
    st.markdown("Esquema del sistema. **Haga clic en una estación** para agregar un ramal que salga de ahí "
                "(derivación del flujo), o elíjala en la lista. Los ramales pueden repartir el flujo con un "
                "**% fijo** (válvula de control) o por **hidráulica** (el % lo dan las presiones de recibo de "
                "cada destino).")
    nodo_sel = ss.get(f"nodo_sel_{ss.sistema}")
    ev = st.plotly_chart(figura_esquema(p, res, nodo_sel), use_container_width=True, on_select="rerun",
                         selection_mode="points", key=f"esq_{ss.sistema}_{ss.ver}",
                         config={"displaylogo": False, "displayModeBar": False})
    try:
        pts_sel = ev.selection.points if ev and ev.selection else []
    except AttributeError:
        pts_sel = []
    if pts_sel and pts_sel[0].get("customdata"):
        nuevo_sel = pts_sel[0]["customdata"]
        if isinstance(nuevo_sel, list):
            nuevo_sel = nuevo_sel[0]
        if nuevo_sel != nodo_sel:
            ss[f"nodo_sel_{ss.sistema}"] = nuevo_sel
            st.rerun()
    st.caption("● estación con bombeo  ⊗ estación en bypass / punto de conexión  ⬟ entrega a cliente "
               "(columna *Clientes* de las estaciones)  — el % sobre cada ramal es la fracción del flujo que se deriva.")

    red = getattr(res, "red", None) or {}
    if red:
        st.markdown("##### 🔀 Nodos de derivación")
        filas = []
        for nom, inf in red.items():
            ram = inf["ramal"]
            padre = m.cadena_de(p, ram.padre)
            nodo_nom = padre.estaciones[ram.nodo].nombre if padre is not None and ram.nodo < len(padre.estaciones) else ""
            ep, er = inf["estrangula_padre"], inf["estrangula_ramal"]
            filas.append({"Ramal": nom, "Sale de": f"{nombre_cadena(p, ram.padre)} – {nodo_nom}", "Modo": ram.modo,
                          "% del flujo": inf["pct_calculado"], "Flujo ramal (KBDO)": inf["flujo"],
                          "Flujo que sigue (KBDO)": float("nan") if inf["flujo_padre_despues"] is None else inf["flujo_padre_despues"],
                          "Presión en el nodo (psig)": inf["p_nodo"],
                          "P. que pide la línea (psig)": float("nan") if inf["d_padre"] is None else inf["d_padre"],
                          "P. que pide el ramal (psig)": inf["d_ramal"],
                          "Válvula de control": ("—" if ram.modo == "Hidráulico" else
                                                 (f"estrangula {er:,.1f} psi en el ramal" if (er or 0) > 0.05 else
                                                  (f"estrangula {ep:,.1f} psi en la línea" if (ep or 0) > 0.05 else "—"))),
                          "Nodo": "bombea (cabezal)" if inf.get("nodo_bombea") else "bypass (sin bombeo)",
                          "Nota": inf.get("nota", "")})
        st.dataframe(pd.DataFrame(filas).style.format(precision=2, na_rep="—"), use_container_width=True,
                     hide_index=True)
        st.caption("**Porcentaje**: la presión del nodo es la mayor que piden las ramas; la otra rama la reduce con "
                   "su válvula de control (estrangula). **Hidráulico**: sin válvula, la presión en el nodo es la "
                   "misma para las dos ramas y el % sale de las presiones de recibo de cada destino. Si la estación "
                   "del nodo está en bypass, esa presión es la que debe llegar desde aguas arriba.")

    # ---------------- agregar ramal
    ops = opciones_nodos(p)
    etiquetas = [o[2] for o in ops]
    idx_def = 0
    if nodo_sel:
        cad_s, i_s = nodo_sel.split("|")
        for k, o in enumerate(ops):
            if o[0] == cad_s and o[1] == int(i_s):
                idx_def = k
    with st.expander("➕ Agregar ramal / tramos nuevos", expanded=bool(nodo_sel) or not p.ramales):
        a1, a2, a3, a4 = st.columns([2.2, 2, 1.4, 1])
        k_nodo = a1.selectbox("Sale de la estación", range(len(ops)), index=idx_def,
                              format_func=lambda k: etiquetas[k], key=f"nr_nodo_{ss.sistema}_{nodo_sel}")
        nom_r = a2.text_input("Nombre del ramal", value=f"Ramal {len(p.ramales) + 1}", key=f"nr_nom_{ss.sistema}")
        modo_r = a3.radio("Reparto del flujo", m.MODOS_RAMAL, key=f"nr_modo_{ss.sistema}",
                          help="Porcentaje: el ramal se lleva un % fijo del flujo que llega al nodo. Hidráulico: "
                               "el % se calcula con las presiones de recibo de cada destino.")
        pct_r = a4.number_input("% del flujo", 0.0, 100.0, 50.0, step=5.0, key=f"nr_pct_{ss.sistema}",
                                disabled=modo_r == "Hidráulico")
        origen = st.radio("Datos de los tramos nuevos", ["Importar la plantilla Excel del ramal",
                                                         "Copiar de otro sistema guardado",
                                                         "Crear vacío y pegar los perfiles"],
                          horizontal=True, key=f"nr_origen_{ss.sistema}")
        otro = None
        if origen.startswith("Importar"):
            arch = st.file_uploader("Plantilla del ramal (.xlsm / .xlsx)", type=["xlsm", "xlsx"],
                                    key=f"nr_xl_{ss.sistema}_{ss.ver}")
            if arch is not None:
                try:
                    otro = cargar_otro_sistema(("excel", arch))
                except Exception as ex:  # noqa: BLE001
                    st.error(f"No se pudo leer la plantilla: {ex}")
        elif origen.startswith("Copiar"):
            otros = [x for x in bd().sistemas() if x != ss.sistema]
            if otros:
                sel_o = st.selectbox("Sistema", otros, key=f"nr_bd_{ss.sistema}")
                otro = cargar_otro_sistema(("bd", sel_o))
            else:
                st.info("No hay otros sistemas en la base de datos.")
        else:
            nombres_txt = st.text_area("Estaciones del ramal en orden (una por línea, sin la estación del nodo)",
                                       "MIRAFLORES\nPORVENIR\nMONTERREY\nAPIAY\nSAN FERNANDO",
                                       key=f"nr_txt_{ss.sistema}", height=140)
        cad_n, i_n, _ = ops[k_nodo]
        kmn = km_nodo(p, cad_n, i_n)
        if otro is not None:
            nombres_o = [e.nombre for e in otro.estaciones]
            nodo_nombre = m.cadena_de(p, cad_n).estaciones[i_n].nombre
            ini_def = next((k for k, x in enumerate(nombres_o[:-1]) if x.strip().upper() == nodo_nombre.strip().upper()), 0)
            b1, b2 = st.columns([2, 3])
            ini = b1.selectbox("Tomar los tramos a partir de la estación", range(len(nombres_o) - 1), index=ini_def,
                               format_func=lambda k: nombres_o[k], key=f"nr_ini_{ss.sistema}",
                               help="Si la plantilla trae toda la línea (p.ej. desde Sebastopol), elija la estación "
                                    "donde empieza el ramal.")
            km0 = otro.tramos[ini].puntos[0].km if otro.tramos[ini].puntos else 0.0
            b2.caption(f"{otro.sistema}: {len(otro.tramos) - ini} tramos desde {nombres_o[ini]} (km {km0:,.3f}) hasta "
                       f"{nombres_o[-1]}. Producto {otro.producto}; se usará el producto de este sistema.")
            desplazar = st.checkbox(f"Desplazar los km del ramal para que continúe desde el nodo (km {kmn:,.3f})",
                                    value=abs(km0 - kmn) > 0.5, key=f"nr_desp_{ss.sistema}")
        if st.button("➕ Agregar ramal", type="primary", key=f"nr_btn_{ss.sistema}"):
            nombre_ok = nom_r.strip()
            if not nombre_ok or any(r.nombre == nombre_ok for r in p.ramales):
                st.error("Use un nombre de ramal que no exista.")
            elif otro is None and not origen.startswith("Crear"):
                st.error("Primero cargue la plantilla o elija el sistema.")
            else:
                if otro is not None:
                    sub = m.Proyecto(estaciones=otro.estaciones[ini:], tramos=otro.tramos[ini:])
                    nuevo = m.nuevo_ramal_desde_proyecto(p, sub, nombre_ok, cad_n, i_n, modo_r, pct_r)
                    if desplazar:
                        desplazar_km(nuevo, kmn - km0)
                else:
                    nombres_l = [x.strip() for x in nombres_txt.splitlines() if x.strip()]
                    nuevo = m.nuevo_ramal_vacio(p, nombre_ok, cad_n, i_n, nombres_l, modo_r, pct_r)
                p.ramales.append(nuevo)
                ss.cambios.append((f"Ramal {nombre_ok}", None, "Ramal agregado", "",
                                   f"desde {etiquetas[k_nodo]}, {len(nuevo.tramos)} tramos", ""))
                ss.pop(f"nodo_sel_{ss.sistema}", None)
                marcar()
                ss.ver += 1
                st.rerun()

    # ---------------- editar ramales
    if p.ramales:
        st.markdown("##### ✏️ Ramales")
        nombres_r = [r.nombre for r in p.ramales]
        rsel = st.selectbox("Ramal", nombres_r, key=f"ram_sel_{ss.sistema}")
        ram = p.ramales[nombres_r.index(rsel)]
        inf = red.get(ram.nombre, {})
        rs_r = list(inf.get("res", []))
        c1, c2, c3, c4 = st.columns([1.4, 1, 1.6, 1])
        modo_n = c1.radio("Reparto", m.MODOS_RAMAL, index=m.MODOS_RAMAL.index(ram.modo) if ram.modo in m.MODOS_RAMAL else 0,
                          horizontal=True, key=f"ram_modo_{rsel}_{ss.ver}")
        pct_n = c2.number_input("% del flujo", 0.0, 100.0, float(ram.pct), step=1.0, key=f"ram_pct_{rsel}_{ss.ver}",
                                disabled=modo_n == "Hidráulico")
        c3.metric("Flujo del ramal", f"{inf.get('flujo', 0):,.2f} KBDO",
                  f"{inf.get('pct_calculado', 0):.2f} % del flujo en el nodo", delta_color="off")
        if modo_n != ram.modo or abs(pct_n - ram.pct) > 1e-12:
            ram.modo, ram.pct = modo_n, pct_n
            marcar()
            ss.ver += 1
            st.rerun()
        if c4.button("🗑️ Quitar ramal", key=f"ram_del_{rsel}"):
            hijos = [r.nombre for r in p.ramales if r.padre == ram.nombre]
            if hijos:
                st.error(f"Primero quite los ramales que salen de este: {', '.join(hijos)}")
            else:
                p.ramales.remove(ram)
                ss.cambios.append((f"Ramal {rsel}", None, "Ramal quitado", rsel, "", ""))
                marcar()
                ss.ver += 1
                st.rerun()
        if inf.get("nota"):
            st.warning(inf["nota"])
        st.markdown("**Estaciones del ramal** (la primera es el nodo de derivación: su %DR y FR aplican al primer "
                    "tramo del ramal)")
        n_r = len(ram.tramos)
        df_e = pd.DataFrame({
            "Estación": [e.nombre for e in ram.estaciones],
            "Presión recibo (psig)": [None if i == 0 else e.presion_recibo for i, e in enumerate(ram.estaciones)],
            "Bypass": [bool(e.paso_directo) for e in ram.estaciones],
            "%DR": [e.dr if i < n_r else None for i, e in enumerate(ram.estaciones)],
            "FR": [e.fr if i < n_r else None for i, e in enumerate(ram.estaciones)],
            "Iny(+)/Entrega(−) bph": [e.delta_bph if 0 < i < n_r else None for i, e in enumerate(ram.estaciones)],
            "Clientes (entregas)": [e.clientes for e in ram.estaciones]})
        ed_e = st.data_editor(df_e, key=f"ram_est_{rsel}_{ss.ver}", hide_index=True, use_container_width=True,
                              num_rows="fixed",
                              column_config={"Presión recibo (psig)": st.column_config.NumberColumn(format="%.2f"),
                                             "%DR": st.column_config.NumberColumn(format="%.4f"),
                                             "FR": st.column_config.NumberColumn(format="%.4f"),
                                             "Iny(+)/Entrega(−) bph": st.column_config.NumberColumn(format="%.2f"),
                                             "Clientes (entregas)": st.column_config.TextColumn(
                                                 help="Separados por coma; solo para el esquema.")})
        cambio = False
        for i, e in enumerate(ram.estaciones):
            f_ = ed_e.iloc[i]
            if i > 0 and str(f_["Estación"]).strip() and f_["Estación"] != e.nombre:
                e.nombre, cambio = str(f_["Estación"]).strip(), True
            bp = bool(f_["Bypass"]) and 0 < i < n_r
            if bp != e.paso_directo:
                e.paso_directo, cambio = bp, True
            v = f_["Presión recibo (psig)"]
            if i > 0 and not e.paso_directo and pd.notna(v) and float(v) != e.presion_recibo:
                e.presion_recibo, cambio = float(v), True
            for col, attr in (("%DR", "dr"), ("FR", "fr"), ("Iny(+)/Entrega(−) bph", "delta_bph")):
                v = f_[col]
                if pd.notna(v) and float(v) != getattr(e, attr):
                    setattr(e, attr, float(v))
                    cambio = True
            cl = "" if pd.isna(f_["Clientes (entregas)"]) else str(f_["Clientes (entregas)"])
            if cl != e.clientes:
                e.clientes, cambio = cl, True
        if cambio:
            marcar()
            ss.ver += 1
            st.rerun()
        if rs_r:
            st.markdown("**Resultados del ramal**")
            st.dataframe(tabla_estaciones_cadena(p, ram.nombre, ram, rs_r).style.format(precision=2, na_rep=""),
                         use_container_width=True, hide_index=True)
        if ram.tramos:
            st.markdown("**Perfil de los tramos del ramal** (pegue desde Excel: punto, km, altura, espesor, "
                        "diámetro, SMYS, rating)")
            nt = [f"{t.nombre}: {ram.estaciones[i].nombre} → {ram.estaciones[i + 1].nombre}"
                  f"{'' if t.puntos else '  (sin perfil)'}" for i, t in enumerate(ram.tramos)]
            it_r = st.radio("Tramo del ramal", range(len(ram.tramos)), format_func=lambda k: nt[k], horizontal=True,
                            key=f"ram_tr_{rsel}")
            t_r = ram.tramos[it_r]
            o1, o2, o3 = st.columns([1, 1, 3])
            rug = o1.number_input("Rugosidad interna (in)", value=float(t_r.rugosidad_in), format="%.5f",
                                  min_value=0.0, key=f"ram_rug_{rsel}_{it_r}_{ss.ver}")
            eff = o2.number_input("Eficiencia (%)", value=min(max(float(t_r.eficiencia), 1.0), 200.0), min_value=1.0,
                                  max_value=200.0, key=f"ram_eff_{rsel}_{it_r}_{ss.ver}")
            if rug != t_r.rugosidad_in or eff != t_r.eficiencia:
                t_r.rugosidad_in, t_r.eficiencia = rug, eff
                marcar()
                st.rerun()
            if editor_perfil_tramo(t_r, rs_r[it_r] if it_r < len(rs_r) else None, f"ram_pts_{rsel}_{it_r}_{ss.ver}"):
                marcar()
                ss.ver += 1
                st.rerun()

        with st.expander("Clientes (entregas) de las estaciones de la línea principal"):
            df_c = pd.DataFrame({"Estación": [e.nombre for e in p.estaciones],
                                 "Clientes (entregas)": [e.clientes for e in p.estaciones]})
            ed_c = st.data_editor(df_c, key=f"cli_main_{ss.ver}", hide_index=True, disabled=["Estación"],
                                  use_container_width=True)
            if any((("" if pd.isna(v) else str(v)) != e.clientes) for v, e in zip(ed_c["Clientes (entregas)"], p.estaciones)):
                for v, e in zip(ed_c["Clientes (entregas)"], p.estaciones):
                    e.clientes = "" if pd.isna(v) else str(v)
                marcar()
                ss.ver += 1
                st.rerun()

        # ---------------- perfil por ruta
        st.markdown("##### 📈 Perfil hidráulico por ruta")
        rutas = m.rutas_red(p)
        k_ruta = st.selectbox("Ruta", range(len(rutas)), format_func=lambda k: rutas[k][0], key=f"ruta_{ss.sistema}",
                              index=min(1, len(rutas) - 1))
        pr_ruta, rs_ruta = m.proyecto_ruta(p, res, rutas[k_ruta][1])
        if any(not t.puntos for t in pr_ruta.tramos):
            st.info("Hay tramos sin perfil en esta ruta: péguelos en el editor de arriba.")
        else:
            curv_r = st.multiselect("Curvas", ["geod", "hid", "maop", "pint", "barlow", "minima"],
                                    default=["geod", "hid", "maop"], format_func=lambda k: CURVAS[k],
                                    key=f"curv_ruta_{ss.sistema}")
            st.plotly_chart(figura_perfil(pr_ruta, rs_ruta, curv_r,
                                          titulo=f"Perfil hidráulico – {rutas[k_ruta][0]} | {p.producto}"),
                            use_container_width=True, config={"displaylogo": False})

# ================================================================ FUGAS
with tab_fuga:
    st.markdown("Simula una **válvula ilícita** o un **derrame** en un km: el caudal que sale se calcula con la "
                "ecuación de orificio **q = Cd·A·√(2·ΔP/ρ)** usando la presión interna en ese punto, y la línea se "
                "recalcula con la fuga (aguas abajo pasa menos caudal, por lo que la presión en el punto también "
                "baja; se resuelve por iteración).")
    lineas = [""] + [r.nombre for r in p.ramales]
    f1, f2, f3 = st.columns([1.3, 1.6, 1.6])
    esc = f1.radio("Escenario", ["Válvula ilícita", "Derrame"], key=f"fg_esc_{ss.sistema}")
    lin = f2.selectbox("Línea", lineas, format_func=lambda x: nombre_cadena(p, x), key=f"fg_lin_{ss.sistema}")
    modo_f = f3.radio("Durante la fuga", m.MODOS_FUGA, key=f"fg_modo_{ss.sistema}",
                      help="Caudal constante: entra al sistema el mismo caudal (la presión de descarga baja). "
                           "Presión de descarga constante: la estación de aguas arriba mantiene su presión y entra "
                           "más caudal (bombas de curva plana). La realidad suele estar entre las dos.")
    cad_f = m.cadena_de(p, lin)
    kms_f = [t.puntos[0].km for t in cad_f.tramos if t.puntos] + [t.puntos[-1].km for t in cad_f.tramos if t.puntos]
    if not kms_f:
        st.info("Esa línea no tiene perfiles cargados.")
    else:
        kmin, kmax = min(kms_f), max(kms_f)
        g1, g2, g3, g4 = st.columns(4)
        km_f = g1.number_input(f"km de la fuga ({kmin:,.3f} a {kmax:,.3f})", min_value=float(kmin),
                               max_value=float(kmax), value=float(round((kmin + kmax) / 2, 3)), step=1.0,
                               format="%.3f", key=f"fg_km_{ss.sistema}_{lin}")
        d_int = m.diametro_interno_en_km(cad_f, km_f) or 0.0
        if esc == "Válvula ilícita":
            DIAM = {"1/4\"": 0.25, "1/2\"": 0.5, "3/4\"": 0.75, "1\"": 1.0, "1 1/2\"": 1.5, "2\"": 2.0, "3\"": 3.0,
                    "4\"": 4.0}
            dsel = g2.selectbox("Diámetro de la válvula", list(DIAM) + ["Otro"], index=1, key=f"fg_dsel_{ss.sistema}")
            diam = DIAM.get(dsel) or g2.number_input("Diámetro (in)", 0.01, 48.0, 0.5, key=f"fg_dotro_{ss.sistema}")
            contra = g3.number_input("Contrapresión a la salida (psig)", 0.0, 5000.0, 0.0, step=5.0,
                                     key=f"fg_contra_{ss.sistema}",
                                     help="0 = descarga a la atmósfera. Use una contrapresión si la válvula descarga "
                                          "a una manguera o carrotanque (o para representar la pérdida de la manguera).")
            tiempo = g4.number_input("Duración (horas)", 0.0, 10000.0, 24.0, step=1.0, key=f"fg_t_{ss.sistema}") * 60
        else:
            tam = g2.select_slider("Tamaño del orificio (% del diámetro interno)",
                                   [5, 10, 20, 30, 50, 75, 100], value=10, key=f"fg_tam_{ss.sistema}",
                                   help="100 % = rotura total (guillotina).")
            diam = d_int * tam / 100
            g2.caption(f"Diámetro interno en el km: {d_int:.3f} in → orificio de {diam:.3f} in")
            contra = 0.0
            tiempo = g3.number_input("Tiempo hasta detectar y parar el bombeo (min)", 0.0, 10000.0, 30.0, step=5.0,
                                     key=f"fg_t_{ss.sistema}")
        cd_f = g4.number_input("Coeficiente de descarga Cd", 0.1, 1.0, 0.61, step=0.01, key=f"fg_cd_{ss.sistema}",
                               help="0,61 = orificio de borde vivo. Una válvula parcialmente abierta o una conexión "
                                    "con manguera tiene un Cd efectivo menor.")
        if st.button("▶ Simular fuga", type="primary", key=f"fg_btn_{ss.sistema}"):
            try:
                with st.spinner("Calculando la fuga…"):
                    ss[f"fuga_{ss.sistema}"] = m.simular_fuga(p, km_f, diam, cd_f, contra, modo_f, lin, tiempo)
            except Exception as ex:  # noqa: BLE001
                st.error(f"No se pudo simular: {ex}")
        f = ss.get(f"fuga_{ss.sistema}")
        if f:
            h1, h2, h3, h4 = st.columns(4)
            h1.metric("Caudal de la fuga", f"{f['q_bph']:,.1f} bph",
                      f"{f['q_bpd']:,.0f} bbl/día · {f['q_m3h']:,.2f} m³/h", delta_color="off")
            h2.metric("% del caudal que llega al punto", f"{f['pct_fuga']:.2f} %",
                      f"llegan {f['flujo_llega_kbd']:,.2f} KBDO", delta_color="off")
            h3.metric("Presión interna en el punto", f"{f['px_fuga']:,.1f} psig",
                      f"{f['px_fuga'] - f['px_base']:+,.1f} psi (sin fuga {f['px_base']:,.1f})", delta_color="off")
            h4.metric("Volumen perdido", f"{f['volumen_bbl']:,.1f} bbl",
                      f"en {f['tiempo_min'] / 60:,.1f} h" if f['tiempo_min'] >= 120 else f"en {f['tiempo_min']:,.0f} min",
                      delta_color="off")
            if f["modo"].startswith("Caudal"):
                st.info(f"Con caudal constante, la presión de descarga de **{f['estacion_control']}** pasa de "
                        f"{f['d_control_base']:,.1f} a **{f['d_control_fuga']:,.1f} psig** "
                        f"({f['d_control_fuga'] - f['d_control_base']:+,.1f} psi): es la señal que vería la operación.")
            else:
                st.info(f"Para mantener la descarga de **{f['estacion_control']}** en {f['d_control_base']:,.1f} psig "
                        f"el caudal de entrada pasa de {f['flujo_entrada_base']:,.2f} a "
                        f"**{f['flujo_entrada_fuga']:,.2f} KBDO** ({(f['flujo_entrada_fuga'] - f['flujo_entrada_base']) * m.KBD_A_BPH:+,.0f} bph).")
            if f["toda_la_corriente"]:
                st.warning("El orificio es tan grande que **todo el caudal que llega sale por la fuga**: aguas abajo la "
                           "línea queda sin flujo (posible columna separada / vaciado).")
            if f.get("nota"):
                st.warning(f["nota"])
            st.dataframe(pd.DataFrame(f["comparacion"]).style.format(precision=2, na_rep=""), use_container_width=True,
                         hide_index=True)
            st.plotly_chart(figura_fuga(f), use_container_width=True, config={"displaylogo": False})
            if esc == "Válvula ilícita" and st.toggle("Ver sensibilidad al diámetro de la válvula",
                                                      key=f"fg_sens_{ss.sistema}"):
                filas_s = []
                barra = st.progress(0.0, text="Calculando sensibilidad…")
                for kd, dd in enumerate((0.25, 0.5, 0.75, 1.0, 1.5, 2.0)):
                    barra.progress((kd + 1) / 6, text=f"Calculando sensibilidad… válvula de {dd} in")
                    try:
                        x = m.simular_fuga(p, f["km"], dd, f["cd"], f["contrapresion_psi"], f["modo"], f["cadena"],
                                           f["tiempo_min"])
                        filas_s.append({"Diámetro (in)": dd, "Fuga (bph)": x["q_bph"], "Fuga (bbl/día)": x["q_bpd"],
                                        "% del caudal": x["pct_fuga"], "P. en el punto (psig)": x["px_fuga"],
                                        f"Descarga {x['estacion_control']} (psig)": x["d_control_fuga"],
                                        "Entrada (KBDO)": x["flujo_entrada_fuga"]})
                    except Exception:  # noqa: BLE001
                        pass
                barra.empty()
                st.dataframe(pd.DataFrame(filas_s).style.format(precision=2), use_container_width=True,
                             hide_index=True)
            st.caption("Supuestos: estado estable (sin golpe de ariete), orificio con Cd constante, presión de recibo "
                       "de las estaciones fija, válvulas de cheque que impiden el contraflujo desde aguas abajo. Es un "
                       "estimado para priorizar; no reemplaza la simulación transitoria (SPS).")

# ================================================================ PERFILES MOP / TRANSITORIOS
with tabs[1]:
    st.markdown("Cargue perfiles de **presión interna** (perfil MOP, transitorios o disparo de válvulas de seguridad "
                "simulados en SPS) y verifique que estén **por debajo de la presión de diseño (Barlow / MAOP)** del "
                "sistema. El archivo (Excel o CSV) debe tener una columna de abscisa en km (*Abscisa (Km)*) y una "
                "de *Presión Interna* en psig. Sirve para cualquier sistema.")
    st.caption("La presión de diseño es la columna MAOP de las hojas de tramo (0,72·2·SMYS·t/D, o el rating de la "
               "válvula cuando existe); en las plantillas nuevas es la 'Presión de diseño Barlow'.")
    subidos = st.file_uploader("Perfiles de presión (.xlsx / .csv)", type=["xlsx", "xlsm", "csv", "txt"],
                               accept_multiple_files=True, key=f"perfiles_{ss.ver}")
    nuevos = []
    for f in subidos or []:
        try:
            nuevos.append(m.leer_perfil_presion((f.name, f.getvalue())))
        except Exception as ex:  # noqa: BLE001
            st.error(f"❌ {f.name}: {ex}")
    if nuevos:
        for j, per in enumerate(nuevos):
            c1, c2, c3 = st.columns([3, 2, 3])
            per["nombre"] = c1.text_input("Nombre del perfil", value=per["nombre"], key=f"np_{j}_{ss.ver}")
            per["tipo"] = c2.selectbox("Tipo", m.TIPOS_PERFIL, index=m.TIPOS_PERFIL.index(per["tipo"]),
                                       key=f"tp_{j}_{ss.ver}")
            c3.caption(f"{per['archivo']}: {len(per['km'])} puntos, km {per['km'][0]:.3f} a {per['km'][-1]:.3f}, "
                       f"presión {min(per['p']):,.0f} a {max(per['p']):,.0f} psig")
        if st.button("➕ Agregar al sistema", type="primary"):
            existentes = {x["nombre"]: k for k, x in enumerate(p.perfiles_externos)}
            for per in nuevos:
                if per["nombre"] in existentes:
                    p.perfiles_externos[existentes[per["nombre"]]] = per
                else:
                    p.perfiles_externos.append(per)
                ss.cambios.append(("Perfiles", None, f"Perfil {per['tipo']}", "", per["nombre"],
                                   f"Cargado de {per['archivo']}"))
            marcar()
            ss.ver += 1
            st.rerun()

    if not p.perfiles_externos:
        st.info("Este sistema todavía no tiene perfiles cargados.")
    else:
        st.caption("Exceso % = (P perfil − P diseño) / P diseño × 100. **Tolerancia**: % que se permite superar la "
                   "presión de diseño (por defecto 10 % en transitorios según API, 0 % en MOP). Puede cambiar el "
                   "tipo y la tolerancia de cada perfil en la tabla.")
        verifs = verificaciones(p, res)
        filas_e = pd.DataFrame([{"Perfil": per["nombre"], "Tipo": per["tipo"],
                                 "Tolerancia (%)": float(m.tolerancia_perfil(per))} for per in p.perfiles_externos])
        ed = st.data_editor(filas_e, hide_index=True, use_container_width=True, key=f"tolp_{ss.sistema}_{ss.ver}",
                            disabled=["Perfil"],
                            column_config={"Tipo": st.column_config.SelectboxColumn(options=list(m.TIPOS_PERFIL)),
                                           "Tolerancia (%)": st.column_config.NumberColumn(min_value=0.0, step=1.0,
                                                                                           format="%.1f")})
        cambio = False
        for per, (_, fila) in zip(p.perfiles_externos, ed.iterrows()):
            if fila["Tipo"] != per["tipo"]:
                per["tipo"] = fila["Tipo"]
                per.pop("tolerancia_pct", None)
                cambio = True
            elif abs(float(fila["Tolerancia (%)"] or 0) - m.tolerancia_perfil(per)) > 1e-9:
                per["tolerancia_pct"] = max(0.0, float(fila["Tolerancia (%)"] or 0))
                cambio = True
        if cambio:
            marcar()
            ss.ver += 1
            st.rerun()

        filas_v = []
        for per in p.perfiles_externos:
            v = verifs[per["nombre"]]
            res_txt = ("✅ " if v["cumple"] else ("⚠️ " if v["aceptable"] else "❌ ")) + v["resultado"]
            filas_v.append({"Perfil": per["nombre"], "Tipo": per["tipo"], "Tolerancia (%)": v["tolerancia_pct"],
                            "Resultado": res_txt, "Puntos que superan": v["n_excesos"],
                            "Fuera de tolerancia": v["n_fuera"], "Máx. exceso (%)": v["max_exceso_pct"],
                            "Máx. exceso (psi)": v["max_exceso_psi"], "en km": v["km_max_exceso"],
                            "P. perfil (psig)": v["p_en_max"], "P. diseño (psig)": v["diseno_en_max"],
                            "Archivo": per["archivo"], "Cargado": per.get("fecha_carga", "")})
        st.dataframe(pd.DataFrame(filas_v).style.format(precision=2, na_rep="")
                     .format({"en km": "{:.3f}", "Tolerancia (%)": "{:g}"}, na_rep=""),
                     use_container_width=True, hide_index=True)
        for per in p.perfiles_externos:
            mensaje_verif(per["nombre"], verifs[per["nombre"]])
        ver_sim = st.toggle("Mostrar también la presión interna del simulador (perfil MOP calculado)", value=True)
        st.plotly_chart(figura_perfil(p, res, ["barlow", "perfiles"] + (["pint"] if ver_sim else []), verifs,
                                      titulo=f"Perfiles de presión interna vs presión de diseño - {p.sistema}"),
                        use_container_width=True, config={"displaylogo": False})
        sel = st.selectbox("Detalle del perfil", [x["nombre"] for x in p.perfiles_externos])
        v = verifs[sel]
        if v["excesos"]:
            st.markdown("**Tramos de km donde se supera la presión de diseño**")
            dfr = pd.DataFrame([{"km desde": r["km_ini"], "km hasta": r["km_fin"], "Puntos": r["n"],
                                 "Máx. exceso (%)": r["max_pct"], "en km": r["km_max"],
                                 "Máx. exceso (psi)": r["max_psi"],
                                 "Estado": "❌ Fuera de tolerancia" if r["fuera"] else "⚠️ Dentro de tolerancia"}
                                for r in v["rangos"]])
            st.dataframe(dfr.style.format({"km desde": "{:.3f}", "km hasta": "{:.3f}", "en km": "{:.3f}",
                                           "Máx. exceso (%)": "{:.2f}", "Máx. exceso (psi)": "{:,.2f}"}),
                         use_container_width=True, hide_index=True, height=min(400, 40 + 35 * len(dfr)))
            st.markdown("**Puntos que superan la presión de diseño**")
            solo_fuera = st.checkbox("Mostrar solo los puntos fuera de tolerancia", value=False)
            ex = [z for z in v["excesos"] if not solo_fuera or z[5] > v["tolerancia_pct"] + 1e-9]
            dfx = pd.DataFrame([{"km": z[0], "P. perfil (psig)": z[1], "P. diseño (psig)": z[2],
                                 "Exceso (psi)": z[4], "Exceso (%)": z[5], "Estado": z[6], "Punto": z[3]}
                                for z in ex], columns=["km", "P. perfil (psig)", "P. diseño (psig)",
                                                       "Exceso (psi)", "Exceso (%)", "Estado", "Punto"])
            st.dataframe(dfx.style.format({"km": "{:.3f}", "P. perfil (psig)": "{:,.2f}",
                                           "P. diseño (psig)": "{:,.2f}", "Exceso (psi)": "{:,.2f}",
                                           "Exceso (%)": "{:.2f}"}),
                         use_container_width=True, hide_index=True, height=320)
        else:
            st.success("Ningún punto supera la presión de diseño.")
        dft = pd.DataFrame([{"km": z[0], "P perfil (psig)": z[1], "P diseño (psig)": z[2], "Exceso (psi)": z[4],
                             "Exceso (%)": z[5], "Estado": z[6], "Punto": z[3]} for z in v["puntos"]])
        st.download_button("⬇️ Descargar verificación completa (CSV)",
                           dft.to_csv(sep=";", index=False, decimal=",").encode("utf-8-sig"),
                           file_name=f"verificacion_{sel}.csv")
        with st.expander("🗑️ Quitar perfiles"):
            quitar = st.multiselect("Perfiles a quitar", [x["nombre"] for x in p.perfiles_externos])
            if quitar and st.button("Quitar seleccionados"):
                p.perfiles_externos = [x for x in p.perfiles_externos if x["nombre"] not in quitar]
                for nmb in quitar:
                    ss.cambios.append(("Perfiles", None, "Perfil quitado", nmb, "", ""))
                marcar()
                st.rerun()
        st.caption("Los perfiles quedan guardados con el sistema al presionar 💾 Guardar.")

# ================================================================ TRAMOS
COLS_EDIT = {"Punto/ubicación": "punto", "dist.del reg. [km]": "km", "altura [m]": "altura", "esp. [in]": "espesor",
             "diám. ext. [in]": "diam_ext", "SMYS [PSI]": "smys", "Rating de Válvula (psi)": "rating"}
with tabs[2]:
    nombres_tr = [f"Tramo {t.nombre}: {p.estaciones[i].nombre} → {p.estaciones[i + 1].nombre}"
                  for i, t in enumerate(p.tramos)]
    it = st.radio("Tramo", range(len(p.tramos)), format_func=lambda k: nombres_tr[k], horizontal=True,
                  label_visibility="collapsed")
    t, r = p.tramos[it], res[it]
    e0 = p.estaciones[it]
    q = r.flujo_kbd
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
    eff = o[1].number_input("✏️ Eficiencia de la tubería (%)", value=min(max(float(t.eficiencia), 1.0), 200.0),
                            min_value=1.0,
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
        df.index = range(t.fila_inicio, len(df) + t.fila_inicio)
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
with tabs[3]:
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
with tabs[4]:
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
with tabs[5]:
    lineas_c = [""] + [r.nombre for r in p.ramales]
    tipo_c = st.radio("Tipo de curva", ["Línea o ramal completo", "Entre dos estaciones"], horizontal=True,
                      key=f"curva_tipo_{ss.sistema}",
                      help="Entre dos estaciones: presión de descarga necesaria en la estación inicial para llevar "
                           "el caudal hasta la estación final con su presión de recibo (p.ej. Sutamarchán → Miraflores).")
    if p.ramales or tipo_c.startswith("Entre"):
        lin_c = st.radio("Línea", lineas_c, horizontal=True, key=f"curva_lin_{ss.sistema}",
                         format_func=lambda x: "Línea principal" + ("" if tipo_c.startswith("Entre")
                                                                    else " (flujo que entra al sistema)") if not x
                         else f"Ramal {x}" + ("" if tipo_c.startswith("Entre") else " (flujo que entra al ramal en su nodo)"))
    else:
        lin_c = ""
    cad_c = m.cadena_de(p, lin_c)
    entre = tipo_c.startswith("Entre")
    if entre:
        nom_est = [e.nombre for e in cad_c.estaciones]
        if lin_c and nom_est:
            nom_est[0] = f"{nom_est[0]} (nodo)"
        e1, e2, e3, e4 = st.columns([1.5, 1.5, 1.2, 1.6])
        ini_c = e1.selectbox("Desde la estación", range(len(nom_est) - 1), format_func=lambda k: nom_est[k],
                             key=f"c_ini_{ss.sistema}_{lin_c}")
        fin_ops = list(range(ini_c + 1, len(nom_est)))
        fin_c = e2.selectbox("Hasta la estación", fin_ops, format_func=lambda k: nom_est[k],
                             key=f"c_fin_{ss.sistema}_{lin_c}_{ini_c}")
        rs_c = list(res) if not lin_c else list((getattr(res, "red", None) or {}).get(lin_c, {}).get("res", []))
        rec_act = rs_c[fin_c - 1].presion_recibo if fin_c - 1 < len(rs_c) else cad_c.estaciones[fin_c].presion_recibo
        rec_c = e3.number_input(f"Recibo en {cad_c.estaciones[fin_c].nombre} (psig)", value=float(round(rec_act, 2)),
                                step=10.0, key=f"c_rec_{ss.sistema}_{lin_c}_{fin_c}")
        sin_bomb = e4.checkbox("Sin bombeo en las estaciones intermedias", value=True,
                               key=f"c_sb_{ss.sistema}_{lin_c}", disabled=fin_c - ini_c < 2,
                               help="Marcado: el trayecto se toma como una sola tubería (las estaciones intermedias no "
                                    "bombean). Sin marcar: las que bombean fijan su presión de recibo y se muestra la "
                                    "descarga de cada una.")
        q_op = m.flujo_en_estacion(p, res, lin_c, ini_c) or 0.0
    else:
        q_op = m.flujo_operacion(p, res, lin_c) or 0.0
    k = st.columns(4)
    clave_c = f"{ss.sistema}_{lin_c}_{entre}_{ini_c if entre else ''}"
    li = k[0].number_input("Flujo mínimo (KBDO)", value=0.0, min_value=0.0, key=f"c_li_{clave_c}")
    ls = k[1].number_input("Flujo máximo (KBDO)", value=max(round(q_op * 1.25, 1), 100.0 if q_op <= 0 else 0.1),
                           min_value=0.1, key=f"c_ls_{clave_c}")
    pasos = k[2].number_input("Puntos", value=10, min_value=2, max_value=50)
    modo = k[3].selectbox("%DR en la curva", ["Sin DRA (%DR = 0)", "%DR actual de cada tramo"])
    st.caption("Con %DR mayor que 100 la fricción queda negativa y la presión de descarga baja al aumentar el "
               "caudal; por eso la curva se calcula por defecto sin DRA."
               + (" Con ramales, la curva de la línea principal usa el reparto actual (% o hidráulico); la de un "
                  "ramal es su curva propia: caudal que entra al ramal vs. presión de descarga en el nodo hacia el "
                  "ramal y en sus estaciones (lo que deben dar las bombas que lo alimentan)." if p.ramales and not entre else "")
               + (" Entre dos estaciones se respetan las entregas/inyecciones intermedias; no se incluyen los ramales "
                  "que salgan de ellas." if entre else ""))
    tramos_c = cad_c.tramos[ini_c:fin_c] if entre else cad_c.tramos
    if any(not t.puntos for t in tramos_c):
        st.warning("Hay tramos sin perfil: la curva solo es válida cuando todos tienen datos.")
    dr_c = 0.0 if modo.startswith("Sin") else None
    if entre:
        curva, nombres_e = m.curva_entre_estaciones(p, lin_c, ini_c, fin_c, li, ls, int(pasos), dr_c, rec_c, sin_bomb)
        titulo_c = f"{cad_c.estaciones[ini_c].nombre} → {cad_c.estaciones[fin_c].nombre} (recibo {rec_c:,.0f} psig)"
    else:
        curva = m.curva_sistema(p, li, ls, int(pasos), dr_c, lin_c)
        nombres_e = m.nombres_curva(p, lin_c)
        titulo_c = p.sistema if not lin_c else f"{p.sistema} – ramal {lin_c}"
    dfc = pd.DataFrame([[q] + ps for q, ps in curva], columns=["KBDO"] + nombres_e)
    figc = go.Figure([go.Scatter(x=dfc["KBDO"], y=dfc[nm], name=f"Descarga {nm}", mode="lines+markers")
                      for nm in nombres_e])
    figc.add_vline(x=q_op, line=dict(dash="dash", color="gray"), annotation_text=f"Operación {q_op:,.2f} KBDO")
    figc.update_layout(template="plotly_white", height=480,
                       xaxis_title="Flujo del ramal (KBDO)" if (lin_c and not entre) else "Flujo (KBDO)",
                       yaxis_title="Presión de descarga (psig)", title=f"Curva del sistema - {titulo_c} ({modo})")
    st.plotly_chart(figc, use_container_width=True, config={"displaylogo": False})
    st.dataframe(dfc.style.format(precision=2), use_container_width=True, hide_index=True)

# ================================================================ PRODUCTOS
with tabs[6]:
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
with tabs[7]:
    hist = pd.DataFrame(db.historial(p.sistema), columns=["Fecha", "Usuario", "Tramo / estación", "km", "Campo",
                                                          "Valor anterior", "Valor nuevo", "Nota"])
    st.dataframe(hist, use_container_width=True, hide_index=True, height=600)
    st.download_button("⬇️ Exportar historial (CSV)", hist.to_csv(sep=";", index=False).encode("utf-8-sig"),
                       file_name=f"historial_{p.sistema}.csv")

# ================================================================ AYUDA
with tabs[8]:
    st.markdown("""
**Cómo usar el simulador**

1. Escriba **su nombre** en la barra lateral (queda en el historial) y escoja el **Sistema**.
2. **Resumen**: cambie Flujo (en KBDO o en bph), INY 1/2, Producto y la tabla de estaciones (presión de recibo,
   Bypass, %DR, FR, Inyección). Todo se recalcula al instante. Debajo están la hoja Resumen con los colores de la
   plantilla y el perfil hidráulico (puede hacer zoom y pasar el mouse sobre las curvas).
3. **🎯 Buscar objetivo** (barra lateral): igual que en Excel — *Definir la celda*, *Con el valor*, *Para cambiar
   la celda*. Ejemplos: Presión descarga SEBASTOPOL = 1900 cambiando %DR; cambiando Flujo (bph) para ajustar el
   KBDO; Rotura de tubo = 0 cambiando la Presión de recibo de la estación siguiente (máxima presión sin roturas).
4. **Perfiles MOP / transitorios**: suba los perfiles de presión interna (MOP, transitorios o disparo de válvulas de
   seguridad de SPS) y el programa calcula, punto a punto, cuánto superan la presión de diseño (psi y %), el máximo y
   los km donde ocurre; en transitorios se permite hasta el 10 % (editable). En el gráfico del
   Resumen puede agregar la Presión interna, la Presión de diseño y los perfiles cargados.
4b. **Presión mínima de recibo (descuelgue)**: en la hoja Resumen, columnas *P. mín. recibo sin descuelgue* y
   *Holgura*, y el panel 🔻 con el punto crítico (km, cota), el recibo mínimo en psig y en metros de columna, la descarga
   necesaria y un botón para aplicarla. La curva *Línea mínima sin descuelgue* la muestra en el gráfico.
4c. **🔀 Red / ramales**: esquema del sistema. Haga clic en una estación para agregar un ramal (tramos nuevos que salen
   de ahí): importe su plantilla Excel, cópielo de otro sistema o créelo vacío y pegue los perfiles. El flujo se reparte
   con un **% fijo** (válvula de control) o por **hidráulica** (con las presiones de recibo de cada destino). Muestra
   la presión en el nodo, cuánto estrangula la válvula y el perfil hidráulico de cada ruta.
4d. **💧 Fugas / válvula ilícita**: caudal que sale por una válvula ilícita o un derrame en un km (ecuación de orificio
   con la presión de la línea, recalculando la línea con la fuga), volumen perdido y cómo cambian las presiones y
   caudales en las estaciones (con caudal constante o con presión de descarga constante).
5. **Tramos**: perfil completo de cada tramo. Las columnas ✏️ se pueden editar o pegar desde Excel.
6. **Actualizar espesores**: por rango de km o con un archivo de rangos de una inspección.
7. **💾 Guardar**: guarda en la base de datos compartida y registra los cambios en el **Historial**. Si otro
   usuario guardó el mismo sistema mientras usted trabajaba, el programa le avisa.
8. **Importar plantillas Excel** (barra lateral): suba una o varias plantillas .xlsm para agregar sistemas.
""")
