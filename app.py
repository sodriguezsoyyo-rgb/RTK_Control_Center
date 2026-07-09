import unicodedata
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

APP_TITLE = "RTK Control Center"
APP_SUBTITLE = "Seguimiento ejecutivo de nodos, expansiones, actualizaciones y avance operativo"
META_NODOS_DIARIOS_POR_GRUPO = 30
BASE = Path(__file__).parent
DATA_FILE = BASE / "data" / "rtk_local.csv"

PALETTE = {
    "bg": "#070B16", "panel": "#101827", "panel2": "#141F33", "grid": "rgba(148,163,184,.15)",
    "text": "#F8FAFC", "muted": "#94A3B8", "cyan": "#22D3EE", "blue": "#3B82F6",
    "purple": "#8B5CF6", "green": "#10B981", "orange": "#F59E0B", "red": "#FB7185",
}
COLORS = [PALETTE["cyan"], PALETTE["blue"], PALETTE["purple"], PALETTE["green"], PALETTE["orange"], PALETTE["red"]]

COLMAP = {
    "ID": "id", "FECHA DE EXPORTACIÓN": "fecha_exportacion", "CÓDIGO ENERGIS EXPORTACIÓN": "codigo_energis_exportacion",
    "CODIGO ENERGIS EXPORTACION": "codigo_energis_exportacion", "ODT HIJA": "odt_hija", "ESTADO INICIAL ODT": "estado_inicial",
    "GRUPO": "grupo", "FECHA DE IMPORTACIÓN": "fecha_importacion", "FECHA DE IMPORTACION": "fecha_importacion", "DIA": "dia",
    "CÓDIGO ENERGIS IMPORTACIÓN": "codigo_energis_importacion", "CODIGO ENERGIS IMPORTACION": "codigo_energis_importacion",
    "ESTADO FINAL ODT": "estado_final", "BARRIO": "barrio", "ZONA": "zona", "COMUNA": "comuna",
    "EXPANSIONES": "expansiones", "ACTUALIZACIÓN": "actualizacion", "ACTUALIZACION": "actualizacion", "NODOS TOTALES": "nodos_totales",
}

BARRIO_CALI_COORDS = {
    "AGUACATAL": (3.462, -76.559), "BOCHALEMA": (3.343, -76.522), "CANEY": (3.384, -76.516),
    "CIUDAD CAPRI": (3.397, -76.535), "COLSEGUROS": (3.429, -76.523), "CRISTOBAL COLON": (3.422, -76.522),
    "EL INGENIO": (3.383, -76.530), "INGENIO": (3.383, -76.530), "EL JARDIN": (3.415, -76.503),
    "LA HACIENDA": (3.377, -76.527), "PRADOS DE NORTE": (3.482, -76.518), "SAN JOAQUIN": (3.369, -76.527),
    "SAN JUDAS TADEO": (3.432, -76.523), "SAN JUDAS TADEO 1": (3.432, -76.521), "SAN PEDRO": (3.452, -76.533),
    "SANTA ANITA": (3.401, -76.529), "SANTA ELENA": (3.424, -76.516), "SANTO DOMINGO": (3.432, -76.516),
    "TRES CRUCES (CHIPICHAPE)": (3.475, -76.528), "URB. COLSEGUROS": (3.427, -76.520), "VALLE DEL LILI": (3.367, -76.519),
    "VIPASA": (3.486, -76.506),
}
COMUNA_COORDS = {
    "1": (3.452, -76.566), "2": (3.485, -76.527), "3": (3.455, -76.535), "4": (3.469, -76.506), "5": (3.482, -76.496),
    "6": (3.492, -76.481), "7": (3.455, -76.494), "8": (3.442, -76.500), "9": (3.440, -76.523), "10": (3.426, -76.527),
    "11": (3.425, -76.505), "12": (3.438, -76.493), "13": (3.422, -76.485), "14": (3.419, -76.470), "15": (3.403, -76.489),
    "16": (3.404, -76.509), "17": (3.382, -76.527), "18": (3.392, -76.548), "19": (3.420, -76.548), "20": (3.421, -76.570),
    "21": (3.399, -76.468), "22": (3.347, -76.531),
}


def norm_text(value):
    text = "" if pd.isna(value) else str(value).strip().upper()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return " ".join(text.split())


def read_any(file):
    if file is None:
        file = DATA_FILE
    name = str(file).lower()
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(file)
    for enc in ("utf-8-sig", "utf-8", "latin1", "cp1252"):
        try:
            return pd.read_csv(file, sep=None, engine="python", encoding=enc)
        except Exception:
            continue
    return pd.read_csv(file, encoding="latin1")


@st.cache_data(show_spinner=False)
def load_data_from_path(path_str):
    return clean_data(read_any(path_str))


def clean_data(df):
    df = df.copy()
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    rename = {}
    for col in df.columns:
        key = norm_text(col)
        rename[col] = COLMAP.get(col, COLMAP.get(key, key.lower().replace(" ", "_")))
    df = df.rename(columns=rename)

    if "id" in df.columns:
        df = df[pd.to_numeric(df["id"], errors="coerce").notna()]

    for col in ["grupo", "estado_final", "estado_inicial", "barrio", "zona", "comuna", "dia"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().replace({"nan": "Sin Dato", "None": "Sin Dato", "": "Sin Dato"})

    for col in ["fecha_exportacion", "fecha_importacion"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)

    for col in ["expansiones", "actualizacion", "nodos_totales", "odt_hija", "codigo_energis_exportacion", "codigo_energis_importacion"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    if "nodos_totales" not in df.columns and {"expansiones", "actualizacion"}.issubset(df.columns):
        df["nodos_totales"] = df["expansiones"].fillna(0) + df["actualizacion"].fillna(0)
    elif "nodos_totales" in df.columns and {"expansiones", "actualizacion"}.issubset(df.columns):
        calculados = df["expansiones"].fillna(0) + df["actualizacion"].fillna(0)
        df["nodos_totales"] = df["nodos_totales"].where(df["nodos_totales"] > 0, calculados).astype(int)

    if "estado_final" in df.columns:
        s = df["estado_final"].map(norm_text)
        df["estado_categoria"] = "Pendiente"
        df.loc[s.str.contains("TERMIN", na=False), "estado_categoria"] = "Terminada"
        df.loc[s.str.contains("ANAL", na=False), "estado_categoria"] = "Análisis"
        df.loc[s.str.contains("IMPORT", na=False), "estado_categoria"] = "Importada Parcial"
    else:
        df["estado_categoria"] = "Sin Dato"

    return df.reset_index(drop=True)


def apply_layout(fig, height=360):
    fig.update_layout(
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=PALETTE["text"], size=12),
        margin=dict(l=10, r=10, t=15, b=10),
        legend=dict(orientation="h", y=-0.18, x=0, bgcolor="rgba(0,0,0,0)"),
        hoverlabel=dict(bgcolor=PALETTE["panel2"], font_size=13, font_family="Inter"),
    )
    fig.update_xaxes(gridcolor=PALETTE["grid"], zeroline=False)
    fig.update_yaxes(gridcolor=PALETTE["grid"], zeroline=False)
    return fig


def metric_card(label, value, caption, color="cyan"):
    st.markdown(f"""
    <div class='metric-card'>
      <div class='metric-label'>{label}</div>
      <div class='metric-value'>{value}</div>
      <div class='metric-caption'>{caption}</div>
      <div class='metric-pill'></div>
    </div>
    """, unsafe_allow_html=True)


def panel_title(title, subtitle=""):
    st.markdown(f"<div class='panel-title'>{title}</div><div class='panel-subtitle'>{subtitle}</div>", unsafe_allow_html=True)


def kpis(df):
    total = len(df)
    terminadas = int((df["estado_categoria"] == "Terminada").sum()) if "estado_categoria" in df else 0
    analisis = int((df["estado_categoria"] == "Análisis").sum()) if "estado_categoria" in df else 0
    exp = int(pd.to_numeric(df.get("expansiones", 0), errors="coerce").fillna(0).sum())
    act = int(pd.to_numeric(df.get("actualizacion", 0), errors="coerce").fillna(0).sum())
    nodos = int(pd.to_numeric(df.get("nodos_totales", 0), errors="coerce").fillna(0).sum())
    dias = max(df["fecha_importacion"].dropna().dt.date.nunique() if "fecha_importacion" in df else 1, 1)
    grupos = max(df["grupo"].replace("Sin Dato", pd.NA).dropna().nunique() if "grupo" in df else 1, 1)
    meta_total = dias * grupos * META_NODOS_DIARIOS_POR_GRUPO
    return {
        "total": total, "terminadas": terminadas, "analisis": analisis, "expansiones": exp, "actualizacion": act,
        "nodos_totales": nodos, "prod_dia": round(nodos / (dias * grupos), 1), "avance": round((nodos / meta_total) * 100, 1),
        "meta_total": meta_total, "dias": int(dias), "grupos": int(grupos)
    }


def top_value(df, col, value_col="nodos_totales"):
    if df.empty or col not in df.columns:
        return "Sin dato"
    if value_col in df.columns:
        data = df.copy()
        data[value_col] = pd.to_numeric(data[value_col], errors="coerce").fillna(0)
        ranking = data.groupby(col, dropna=True)[value_col].sum().sort_values(ascending=False)
        ranking = ranking[ranking > 0]
        return str(ranking.index[0]) if not ranking.empty else "Sin dato"
    vc = df[col].value_counts()
    return str(vc.index[0]) if not vc.empty else "Sin dato"


def daily_bars(df):
    if df.empty or "fecha_importacion" not in df.columns:
        return apply_layout(go.Figure(), 390)
    data = df.dropna(subset=["fecha_importacion"]).copy()
    data["nodos_totales"] = pd.to_numeric(data.get("nodos_totales", 0), errors="coerce").fillna(0)
    data = data.groupby("fecha_importacion", as_index=False)["nodos_totales"].sum().sort_values("fecha_importacion")
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=data["fecha_importacion"], y=data["nodos_totales"], name="Nodos trabajados",
        marker=dict(color=PALETTE["cyan"], line=dict(color="rgba(255,255,255,.25)", width=1)),
        text=data["nodos_totales"], texttemplate="%{text:,.0f}", textposition="outside",
        hovertemplate="Fecha: %{x|%d/%m/%Y}<br>Nodos: %{y:,.0f}<extra></extra>",
    ))
    fig.update_layout(
        bargap=.28, dragmode="zoom", hovermode="x unified",
        xaxis=dict(
            rangeslider=dict(visible=True, thickness=.08),
            rangeselector=dict(
                buttons=[dict(count=7, label="7d", step="day", stepmode="backward"), dict(count=15, label="15d", step="day", stepmode="backward"), dict(count=1, label="1m", step="month", stepmode="backward"), dict(step="all", label="Todo")],
                bgcolor=PALETTE["panel2"], activecolor=PALETTE["cyan"], font=dict(color=PALETTE["text"]),
            ),
        ),
    )
    fig.update_yaxes(title_text="Nodos")
    return apply_layout(fig, 410)


def donut(df, col):
    if df.empty or col not in df.columns:
        return apply_layout(go.Figure(), 360)
    data = df[col].value_counts().reset_index()
    data.columns = [col, "cantidad"]
    fig = px.pie(data, names=col, values="cantidad", hole=.64, color_discrete_sequence=COLORS)
    fig.update_traces(textinfo="percent+label", hovertemplate="%{label}<br>Asignaciones: %{value}<extra></extra>")
    return apply_layout(fig, 360)


def bar_by_value(df, col, top=10, orientation="h", value_col="nodos_totales", height=360):
    if df.empty or col not in df.columns:
        return apply_layout(go.Figure(), height)
    data = df.copy()
    data[value_col] = pd.to_numeric(data.get(value_col, 0), errors="coerce").fillna(0)
    vc = data.groupby(col, dropna=True)[value_col].sum().sort_values(ascending=False).head(top).reset_index()
    vc.columns = [col, "nodos"]
    if vc.empty:
        return apply_layout(go.Figure(), height)
    if orientation == "h":
        vc = vc.sort_values("nodos", ascending=True)
        fig = px.bar(vc, x="nodos", y=col, orientation="h", text="nodos", color="nodos", color_continuous_scale=[PALETTE["blue"], PALETTE["cyan"]])
        fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside", hovertemplate="%{y}<br>Nodos: %{x:,.0f}<extra></extra>")
        fig.update_xaxes(title_text="Nodos")
        fig.update_yaxes(title_text="")
    else:
        fig = px.bar(vc, x=col, y="nodos", text="nodos", color="nodos", color_continuous_scale=[PALETTE["blue"], PALETTE["cyan"]])
        fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside", hovertemplate="%{x}<br>Nodos: %{y:,.0f}<extra></extra>")
        fig.update_yaxes(title_text="Nodos")
        fig.update_xaxes(title_text="")
    fig.update_layout(showlegend=False, coloraxis_showscale=False)
    return apply_layout(fig, height)


def stacked_by_group(df):
    if df.empty or "grupo" not in df.columns:
        return apply_layout(go.Figure(), 360)
    data = df.copy()
    data["nodos_totales"] = pd.to_numeric(data.get("nodos_totales", 0), errors="coerce").fillna(0)
    data = data.groupby(["grupo", "estado_categoria"], as_index=False)["nodos_totales"].sum()
    fig = px.bar(data, x="grupo", y="nodos_totales", color="estado_categoria", barmode="stack", text="nodos_totales", color_discrete_sequence=COLORS)
    fig.update_traces(texttemplate="%{text:,.0f}", hovertemplate="Grupo: %{x}<br>Nodos: %{y:,.0f}<extra></extra>")
    fig.update_yaxes(title_text="Nodos")
    fig.update_xaxes(title_text="")
    return apply_layout(fig, 360)


def gauge(value):
    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=float(value), number={"suffix": "%", "font": {"size": 44}},
        title={"text": "Avance sobre meta", "font": {"size": 16}},
        gauge={"axis": {"range": [0, 100]}, "bar": {"color": PALETTE["cyan"]}, "bgcolor": PALETTE["panel2"], "borderwidth": 0,
               "steps": [{"range": [0, 60], "color": "rgba(251,113,133,.18)"}, {"range": [60, 85], "color": "rgba(245,158,11,.20)"}, {"range": [85, 100], "color": "rgba(16,185,129,.22)"}],
               "threshold": {"line": {"color": PALETTE["green"], "width": 4}, "thickness": .75, "value": 100}},
    ))
    return apply_layout(fig, 360)


def fallback_coord(barrio, comuna):
    import hashlib
    base = COMUNA_COORDS.get(str(comuna).replace(".0", ""), (3.43, -76.52))
    seed = int(hashlib.md5(f"{barrio}-{comuna}".encode()).hexdigest()[:6], 16)
    return base[0] + ((seed % 100) - 50) / 10000, base[1] + (((seed // 100) % 100) - 50) / 10000


def mapa_barrios(df):
    if df.empty or "barrio" not in df.columns:
        return apply_layout(go.Figure(), 480)
    data = df.copy()
    for col in ["nodos_totales", "expansiones", "actualizacion"]:
        data[col] = pd.to_numeric(data.get(col, 0), errors="coerce").fillna(0)
    grouped = data.groupby(["comuna", "barrio"], dropna=False).agg(
        asignaciones=("barrio", "size"), nodos=("nodos_totales", "sum"), expansiones=("expansiones", "sum"), actualizaciones=("actualizacion", "sum"),
        terminadas=("estado_categoria", lambda x: int((x == "Terminada").sum())), analisis=("estado_categoria", lambda x: int((x == "Análisis").sum())),
        grupos=("grupo", lambda x: ", ".join(sorted(set(map(str, x.dropna()))))),
    ).reset_index()
    lats, lons = [], []
    for _, r in grouped.iterrows():
        key = norm_text(r["barrio"])
        lat, lon = BARRIO_CALI_COORDS.get(key, fallback_coord(r["barrio"], r["comuna"]))
        lats.append(lat); lons.append(lon)
    grouped["lat"] = lats; grouped["lon"] = lons
    grouped["tamano"] = grouped["nodos"].clip(lower=8)
    fig = px.scatter_mapbox(
        grouped, lat="lat", lon="lon", size="tamano", color="comuna", hover_name="barrio",
        hover_data={"comuna": True, "nodos": True, "expansiones": True, "actualizaciones": True, "asignaciones": True, "terminadas": True, "analisis": True, "grupos": True, "lat": False, "lon": False, "tamano": False},
        color_discrete_sequence=COLORS, size_max=42, zoom=11, height=500,
    )
    fig.update_layout(mapbox_style="carto-darkmatter", mapbox_center={"lat": 3.43, "lon": -76.52}, mapbox_zoom=11, legend=dict(orientation="h", y=-.1, x=0))
    return apply_layout(fig, 500)


def filter_data(df):
    st.sidebar.markdown("### Filtros")
    out = df.copy()
    grupos = st.sidebar.multiselect("Grupo", sorted(out["grupo"].dropna().unique())) if "grupo" in out else []
    estados = st.sidebar.multiselect("Estado", sorted(out["estado_categoria"].dropna().unique())) if "estado_categoria" in out else []
    comunas = st.sidebar.multiselect("Comuna", sorted(out["comuna"].dropna().unique())) if "comuna" in out else []
    barrios_base = out[out["comuna"].isin(comunas)] if comunas and "comuna" in out else out
    barrios = st.sidebar.multiselect("Barrio", sorted(barrios_base["barrio"].dropna().unique())) if "barrio" in out else []
    zonas = st.sidebar.multiselect("Zona", sorted(out["zona"].dropna().unique())) if "zona" in out else []
    if grupos: out = out[out["grupo"].isin(grupos)]
    if estados: out = out[out["estado_categoria"].isin(estados)]
    if comunas: out = out[out["comuna"].isin(comunas)]
    if barrios: out = out[out["barrio"].isin(barrios)]
    if zonas: out = out[out["zona"].isin(zonas)]
    if "fecha_importacion" in out and not out["fecha_importacion"].dropna().empty:
        min_d, max_d = out["fecha_importacion"].min().date(), out["fecha_importacion"].max().date()
        start, end = st.sidebar.date_input("Rango de fechas", (min_d, max_d))
        out = out[(out["fecha_importacion"].dt.date >= start) & (out["fecha_importacion"].dt.date <= end)]
    return out


st.set_page_config(page_title=APP_TITLE, page_icon="📍", layout="wide")
css = BASE / "assets" / "styles.css"
if css.exists():
    st.markdown(f"<style>{css.read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)

st.sidebar.markdown("# 📍 RTK Analytics")
st.sidebar.caption("Centro de control operativo")
uploaded = st.sidebar.file_uploader("Cargar Excel/CSV actualizado", type=["csv", "xlsx", "xls"])

try:
    df = clean_data(read_any(uploaded)) if uploaded is not None else load_data_from_path(str(DATA_FILE))
except Exception as e:
    st.error(f"No pude cargar la base de datos: {e}")
    st.stop()

fdf = filter_data(df)

st.markdown(f"""
<div class='header-wrap'>
  <div class='kicker'>EMCALI · Alumbrado Público</div>
  <h1 class='title'>{APP_TITLE}</h1>
  <div class='subtitle'>{APP_SUBTITLE}</div>
</div>
""", unsafe_allow_html=True)

m = kpis(fdf)
c1, c2, c3, c4, c5, c6 = st.columns(6)
with c1: metric_card("Nodos", f"{m['nodos_totales']:,}".replace(",", "."), "Total trabajados", "cyan")
with c2: metric_card("Expansiones", f"{m['expansiones']:,}".replace(",", "."), "Nodos expansión", "blue")
with c3: metric_card("Actualizaciones", f"{m['actualizacion']:,}".replace(",", "."), "Nodos actualización", "purple")
with c4: metric_card("Terminadas", f"{m['terminadas']:,}".replace(",", "."), "Asignaciones cerradas", "good")
with c5: metric_card("En análisis", f"{m['analisis']:,}".replace(",", "."), "Asignaciones pendientes", "warn")
with c6: metric_card("Nodos/grupo/día", f"{m['prod_dia']}", "Meta: 30", "cyan")

st.write("")
left, mid, right = st.columns([1.35, .9, .9])
with left:
    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    panel_title("Tendencia diaria", "Barras interactivas por nodos trabajados; puedes hacer zoom y usar el selector inferior")
    st.plotly_chart(daily_bars(fdf), use_container_width=True, config={"displaylogo": False, "scrollZoom": True})
    st.markdown("</div>", unsafe_allow_html=True)
with mid:
    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    panel_title("Estado de ODT", "Distribución por estado final")
    st.plotly_chart(donut(fdf, "estado_categoria"), use_container_width=True, config={"displaylogo": False})
    st.markdown("</div>", unsafe_allow_html=True)
with right:
    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    panel_title("Meta operativa", f"30 nodos diarios por grupo · meta filtrada: {m['meta_total']:,}".replace(",", "."))
    st.plotly_chart(gauge(min(m["avance"], 100)), use_container_width=True, config={"displaylogo": False})
    st.markdown("</div>", unsafe_allow_html=True)

st.write("")
st.markdown("<div class='panel'>", unsafe_allow_html=True)
panel_title("Mapa interactivo de Cali", "Barrios trabajados con filtros por comuna y barrio")
st.plotly_chart(mapa_barrios(fdf), use_container_width=True, config={"displaylogo": False, "scrollZoom": True})
st.caption("Nota: el Excel actual no trae latitud/longitud por nodo. Por ahora el mapa ubica cada barrio con coordenadas de referencia; cuando conectemos una base geográfica real, se mostrarán los puntos exactos.")
st.markdown("</div>", unsafe_allow_html=True)

st.write("")
a, b = st.columns([1, 1])
with a:
    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    panel_title("Producción por grupo", "Nodos por cuadrilla / levantadores")
    st.plotly_chart(stacked_by_group(fdf), use_container_width=True, config={"displaylogo": False})
    st.markdown("</div>", unsafe_allow_html=True)
with b:
    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    panel_title("Top barrios", f"Barrio con más nodos: {top_value(fdf, 'barrio')}")
    st.plotly_chart(bar_by_value(fdf, "barrio", top=10, orientation="h", height=360), use_container_width=True, config={"displaylogo": False})
    st.markdown("</div>", unsafe_allow_html=True)

st.write("")
a, b, c = st.columns(3)
with a:
    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    panel_title("Zonas", "Nodos trabajados por zona")
    st.plotly_chart(bar_by_value(fdf, "zona", 8, "h"), use_container_width=True, config={"displaylogo": False})
    st.markdown("</div>", unsafe_allow_html=True)
with b:
    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    panel_title("Comunas", "Nodos trabajados por comuna")
    st.plotly_chart(bar_by_value(fdf, "comuna", 10, "h"), use_container_width=True, config={"displaylogo": False})
    st.markdown("</div>", unsafe_allow_html=True)
with c:
    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    panel_title("Días", "Nodos trabajados por día")
    st.plotly_chart(bar_by_value(fdf, "dia", 7, "h"), use_container_width=True, config={"displaylogo": False})
    st.markdown("</div>", unsafe_allow_html=True)

st.write("")
st.markdown("<div class='panel'>", unsafe_allow_html=True)
panel_title("Detalle operativo", "Tabla filtrable para revisión y exportación")
st.dataframe(fdf, use_container_width=True, hide_index=True)
st.download_button("Descargar datos filtrados", fdf.to_csv(index=False).encode("utf-8-sig"), "rtk_filtrado.csv", "text/csv", type="primary")
st.markdown("</div>", unsafe_allow_html=True)
