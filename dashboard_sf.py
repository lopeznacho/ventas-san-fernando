"""
Dashboard Ventas & Abastecimiento — San Fernando
Cargá el archivo de stock maestro + hasta 6 archivos de ventas mensuales.
"""

import re
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from datetime import datetime

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Ventas & Stock — San Fernando",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
.alert-rojo      { background:#3d1010; border-left:4px solid #e74c3c;
                   padding:10px 14px; border-radius:5px; margin:4px 0; }
.alert-amarillo  { background:#3d2e10; border-left:4px solid #f39c12;
                   padding:10px 14px; border-radius:5px; margin:4px 0; }
.alert-verde     { background:#0f3d1a; border-left:4px solid #27ae60;
                   padding:10px 14px; border-radius:5px; margin:4px 0; }
.alert-violeta   { background:#1a1040; border-left:4px solid #9b59b6;
                   padding:10px 14px; border-radius:5px; margin:4px 0; }
</style>
""",
    unsafe_allow_html=True,
)

# ── Constants ─────────────────────────────────────────────────────────────────
MESES_ORDEN = ["Noviembre", "Diciembre", "Enero", "Febrero", "Marzo", "Abril"]
N_MESES_TOTAL = 6  # denominador fijo para promedio

DEFAULT_ROT_ALTA   = 10
DEFAULT_ROT_MEDIA  = 3
DEFAULT_COB_CRITICA    = 1.0
DEFAULT_COB_SOBRESTOCK = 4.0


# ── Helpers ───────────────────────────────────────────────────────────────────

def _norm(text: str) -> str:
    """Normalise string: lowercase, remove accents, collapse spaces."""
    text = str(text).strip().lower()
    for src, dst in zip("áéíóúüñ", "aeiouun"):
        text = text.replace(src, dst)
    return re.sub(r"\s+", "_", re.sub(r"[^a-z0-9_\s]", "", text))


def detect_mapping(df: pd.DataFrame) -> dict:
    """Return best-guess column mapping for known semantic fields."""
    norm_map = {_norm(c): c for c in df.columns}
    candidates = {
        "codigo":      ["codigo", "cod", "codart", "codarticulo", "sku", "id",
                        "articulo_id", "code", "item"],
        "descripcion": ["descripcion", "desc", "nombre", "articulo", "producto",
                        "detalle", "name", "denominacion"],
        "cantidad":    ["cantidad", "cant", "qty", "unidades", "vendido", "venta",
                        "cantidad_vendida", "unid_vendidas", "ventas"],
        "stock":       ["stock", "saldo", "existencia", "stock_actual", "disponible",
                        "cantidad_stock", "inventario"],
        "rubro":       ["rubro", "categoria", "familia", "seccion", "family",
                        "category", "linea"],
        "marca":       ["marca", "brand", "proveedor", "fabricante", "supplier"],
        "precio":      ["precio", "price", "precio_venta", "pvp", "precio_unit"],
    }
    mapping = {}
    for field, names in candidates.items():
        for name in names:
            if name in norm_map:
                mapping[field] = norm_map[name]
                break
    return mapping


def clean_code(val) -> str:
    if pd.isna(val):
        return ""
    return str(val).strip().upper()


def to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace(",", ".", regex=False).str.strip(),
        errors="coerce",
    ).fillna(0.0)


# ── Loaders ───────────────────────────────────────────────────────────────────

def load_stock(file) -> pd.DataFrame | None:
    try:
        df = pd.read_excel(file, dtype=str)
        df.columns = [c.strip() for c in df.columns]
    except Exception as e:
        st.error(f"Error leyendo stock: {e}")
        return None

    m = detect_mapping(df)

    if "codigo" not in m:
        st.error(
            f"**Stock maestro:** no se encontró columna de código.\n\n"
            f"Columnas detectadas: `{list(df.columns)}`\n\n"
            "Renombrá la columna a **Codigo** y volvé a cargar."
        )
        return None

    out = pd.DataFrame()
    out["codigo"] = df[m["codigo"]].apply(clean_code)

    out["descripcion"] = (
        df[m["descripcion"]].str.strip() if "descripcion" in m else out["codigo"]
    )

    out["stock_actual"] = to_num(df[m["stock"]]) if "stock" in m else 0.0
    if "stock" not in m:
        st.warning("⚠️ Stock maestro: no se encontró columna de stock. Se asignó 0.")

    if "rubro" in m:
        out["rubro"] = df[m["rubro"]].str.strip().str.title().fillna("Sin rubro")
    if "marca" in m:
        out["marca"] = df[m["marca"]].str.strip().str.title().fillna("Sin marca")
    if "precio" in m:
        out["precio"] = to_num(df[m["precio"]])

    out = out[out["codigo"] != ""].copy()

    dupes = out[out["codigo"].duplicated(keep=False)]
    if not dupes.empty:
        st.error(
            f"🚨 **Stock maestro:** {dupes['codigo'].nunique()} códigos duplicados.\n\n"
            "Esto es un problema en el archivo fuente. Corregilo antes de continuar."
        )
        with st.expander("Ver duplicados"):
            st.dataframe(dupes.sort_values("codigo"))
        return None

    return out.reset_index(drop=True)


def load_sales(file, mes: str) -> pd.DataFrame | None:
    try:
        df = pd.read_excel(file, dtype=str)
        df.columns = [c.strip() for c in df.columns]
    except Exception as e:
        st.error(f"Error leyendo {mes}: {e}")
        return None

    m = detect_mapping(df)
    missing = [f for f in ("codigo", "cantidad") if f not in m]
    if missing:
        st.warning(
            f"⚠️ **{mes}:** no se encontró columna de `{'` / `'.join(missing)}`.\n\n"
            f"Columnas detectadas: `{list(df.columns)}`"
        )
        return None

    out = pd.DataFrame()
    out["codigo"] = df[m["codigo"]].apply(clean_code)
    out["cantidad"] = to_num(df[m["cantidad"]])
    if "descripcion" in m:
        out["_desc_venta"] = df[m["descripcion"]].str.strip()

    out = out[out["codigo"] != ""].copy()

    # Collapse duplicates with a warning
    dup_codes = out[out["codigo"].duplicated(keep=False)]["codigo"].nunique()
    if dup_codes:
        st.warning(
            f"⚠️ **{mes}:** {dup_codes} códigos duplicados → se suman las cantidades."
        )
        out = out.groupby("codigo", as_index=False)["cantidad"].sum()

    out["mes"] = mes
    return out


# ── Consolidation ─────────────────────────────────────────────────────────────

def build_base(
    stock_df: pd.DataFrame,
    sales_by_month: dict,
    rot_alta: float,
    rot_media: float,
    cob_critica: float,
    cob_sobrestock: float,
) -> tuple[pd.DataFrame, list[dict]]:
    """
    Returns (base_df, orphan_rows).
    base_df has one row per product from the stock master.
    """
    base = stock_df.copy().set_index("codigo")

    meses_disponibles = [m for m in MESES_ORDEN if m in sales_by_month]
    for mes in meses_disponibles:
        base[mes] = 0.0

    orphans = []
    for mes, df_mes in sales_by_month.items():
        for _, row in df_mes.iterrows():
            cod = row["codigo"]
            if cod in base.index:
                base.loc[cod, mes] = row["cantidad"]
            else:
                orphans.append(
                    {
                        "codigo": cod,
                        "mes": mes,
                        "cantidad": row["cantidad"],
                        "descripcion_en_venta": row.get("_desc_venta", ""),
                    }
                )

    base = base.reset_index()

    # ── Metrics ──────────────────────────────────────────────────────────────
    mes_cols = [m for m in MESES_ORDEN if m in base.columns]
    base["total_vendido"]    = base[mes_cols].sum(axis=1) if mes_cols else 0.0
    base["promedio_mensual"] = base["total_vendido"] / N_MESES_TOTAL

    base["cobertura_meses"] = np.where(
        base["promedio_mensual"] > 0,
        base["stock_actual"] / base["promedio_mensual"],
        np.inf,
    )

    # Rotation
    rot_conds   = [base["promedio_mensual"] >= rot_alta,
                   base["promedio_mensual"] >= rot_media,
                   base["promedio_mensual"] >  0]
    rot_choices = ["Alta", "Media", "Baja"]
    base["rotacion"] = np.select(rot_conds, rot_choices, default="Sin movimiento")

    # Stock status
    def _estado(row):
        pm  = row["promedio_mensual"]
        stk = row["stock_actual"]
        cob = row["cobertura_meses"]
        if pm == 0 and stk == 0:    return "Sin stock / Sin venta"
        if pm == 0 and stk > 0:     return "Sobrestock (sin venta)"
        if stk == 0 and pm > 0:     return "Quiebre"
        if cob < cob_critica:        return "Crítico"
        if cob <= cob_sobrestock:    return "Normal"
        return "Sobrestock"

    base["estado_stock"] = base.apply(_estado, axis=1)

    # Buying priority score
    base["prioridad_compra"] = np.where(
        base["promedio_mensual"] > 0,
        base["promedio_mensual"] / (base["stock_actual"] + 0.01),
        0.0,
    )

    # Value columns (if price available)
    if "precio" in base.columns:
        base["valor_stock"]   = base["stock_actual"]   * base["precio"]
        base["valor_vendido"] = base["total_vendido"]  * base["precio"]

    return base, orphans


# ── Styling helpers ───────────────────────────────────────────────────────────

_ESTADO_COLORS = {
    "Quiebre":               "background-color:#3d1010;color:#e74c3c;font-weight:bold",
    "Crítico":               "background-color:#3d2e10;color:#e67e22;font-weight:bold",
    "Normal":                "background-color:#0f3d1a;color:#2ecc71",
    "Sobrestock":            "background-color:#1a1040;color:#9b59b6",
    "Sobrestock (sin venta)":"background-color:#1a1040;color:#8e44ad;font-style:italic",
    "Sin stock / Sin venta": "color:#7f8c8d;font-style:italic",
}

_ROT_COLORS = {
    "Alta":           "color:#27ae60;font-weight:bold",
    "Media":          "color:#f39c12",
    "Baja":           "color:#e74c3c",
    "Sin movimiento": "color:#7f8c8d;font-style:italic",
}


def _style_estado(val):   return _ESTADO_COLORS.get(val, "")
def _style_rot(val):      return _ROT_COLORS.get(val, "")


def _fmt_cob(x):
    if x == np.inf or x > 999:
        return "∞"
    return f"{x:.1f}"


def to_excel(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Dashboard")
    return buf.getvalue()


# ── APP ───────────────────────────────────────────────────────────────────────

def main():
    st.title("🛒 Ventas & Abastecimiento — San Fernando")
    st.caption(
        f"Período base: Noviembre 2024 – Abril 2025 ({N_MESES_TOTAL} meses)  "
        f"•  Actualizado: {datetime.today().strftime('%d/%m/%Y')}"
    )

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("📂 Archivos")

        st.markdown("**1 — Stock maestro** *(todos los productos)*")
        stock_file = st.file_uploader(
            "stock_maestro.xlsx", type=["xlsx", "xls"], key="stock", label_visibility="collapsed"
        )

        st.markdown("**2 — Ventas mensuales** *(uno por mes)*")
        uploaded = {}
        for mes in MESES_ORDEN:
            f = st.file_uploader(mes, type=["xlsx", "xls"], key=mes)
            if f:
                uploaded[mes] = f

        st.divider()
        st.header("⚙️ Umbrales")
        rot_alta   = st.number_input("Rotación ALTA (u/mes ≥)",  value=DEFAULT_ROT_ALTA,   min_value=1, step=1)
        rot_media  = st.number_input("Rotación MEDIA (u/mes ≥)", value=DEFAULT_ROT_MEDIA,  min_value=1, step=1)
        cob_critica    = st.number_input("Cobertura CRÍTICA (< meses)",    value=DEFAULT_COB_CRITICA,    step=0.5, min_value=0.1)
        cob_sobrestock = st.number_input("Cobertura SOBRESTOCK (> meses)", value=DEFAULT_COB_SOBRESTOCK, step=0.5, min_value=1.0)

        st.divider()
        run_btn = st.button("▶ Procesar", type="primary", use_container_width=True)

    # ── Welcome screen ────────────────────────────────────────────────────────
    if not stock_file:
        st.info("👈 Cargá primero el **stock maestro** desde el panel lateral.")
        st.markdown(
            """
### Archivos necesarios
| # | Archivo | Contenido |
|---|---------|-----------|
| 1 | **Stock maestro** | Todos los artículos: código, descripción, stock, y opcionalmente rubro / marca / precio |
| 2–7 | **Ventas Nov → Abr** | Artículos vendidos: código + cantidad por mes |

### Columnas que se detectan automáticamente
**Stock:** `Codigo`, `Descripcion`, `Stock`, `Rubro`, `Marca`, `Precio`
**Ventas:** `Codigo`, `Cantidad` *(o variantes como `Cod`, `SKU`, `Cant`, `Vendido`, etc.)*

> Si el sistema no detecta una columna, te avisa con el nombre exacto encontrado para que puedas renombrarla.
"""
        )
        return

    if not run_btn and "base" not in st.session_state:
        st.info("👈 Cargá los archivos y hacé clic en **Procesar**.")
        return

    # ── Process ───────────────────────────────────────────────────────────────
    if run_btn:
        with st.spinner("Procesando…"):
            stock_df = load_stock(stock_file)
            if stock_df is None:
                return

            sales_by_month = {}
            for mes, f in uploaded.items():
                df_mes = load_sales(f, mes)
                if df_mes is not None:
                    sales_by_month[mes] = df_mes

            base, orphans = build_base(
                stock_df, sales_by_month,
                rot_alta, rot_media, cob_critica, cob_sobrestock
            )

        st.session_state.update(
            base=base,
            orphans=orphans,
            meses=list(sales_by_month.keys()),
            thresholds=dict(
                rot_alta=rot_alta, rot_media=rot_media,
                cob_critica=cob_critica, cob_sobrestock=cob_sobrestock
            ),
        )
        st.success(
            f"✅ {len(base):,} productos cargados  •  "
            f"{len(sales_by_month)} archivos de ventas  •  "
            f"{len(orphans)} códigos huérfanos"
        )

    if "base" not in st.session_state:
        return

    base: pd.DataFrame = st.session_state["base"]
    orphans: list       = st.session_state["orphans"]
    meses_cargados      = st.session_state["meses"]

    mes_cols = [m for m in MESES_ORDEN if m in base.columns]

    # ── Alert banner ──────────────────────────────────────────────────────────
    n_quiebre = (base["estado_stock"] == "Quiebre").sum()
    n_critico = (base["estado_stock"] == "Crítico").sum()
    n_sobre   = base["estado_stock"].str.startswith("Sobrestock").sum()
    n_sin_venta_stock = (
        (base["total_vendido"] == 0) & (base["stock_actual"] > 0)
    ).sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f'<div class="alert-rojo">🚨 <b>{n_quiebre}</b> en QUIEBRE</div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="alert-amarillo">⚠️ <b>{n_critico}</b> CRÍTICOS</div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="alert-violeta">📦 <b>{n_sobre}</b> SOBRESTOCK</div>', unsafe_allow_html=True)
    c4.markdown(f'<div class="alert-amarillo">😴 <b>{n_sin_venta_stock}</b> sin venta c/stock</div>', unsafe_allow_html=True)

    st.divider()

    # ── Filters ───────────────────────────────────────────────────────────────
    st.subheader("🔎 Filtros")
    fc = st.columns(5)

    estados_opts = ["Todos"] + sorted(base["estado_stock"].unique().tolist())
    f_estado = fc[0].selectbox("Estado de stock", estados_opts)

    rot_opts = ["Todas", "Alta", "Media", "Baja", "Sin movimiento"]
    f_rot = fc[1].selectbox("Rotación", rot_opts)

    rubro_opts = (
        ["Todos"] + sorted(base["rubro"].dropna().unique().tolist())
        if "rubro" in base.columns else ["Todos"]
    )
    f_rubro = fc[2].selectbox("Rubro / Categoría", rubro_opts, disabled="rubro" not in base.columns)

    marca_opts = (
        ["Todas"] + sorted(base["marca"].dropna().unique().tolist())
        if "marca" in base.columns else ["Todas"]
    )
    f_marca = fc[3].selectbox("Marca / Proveedor", marca_opts, disabled="marca" not in base.columns)

    f_search = fc[4].text_input("Buscar código / producto")

    filtered = base.copy()
    if f_estado != "Todos":
        filtered = filtered[filtered["estado_stock"] == f_estado]
    if f_rot != "Todas":
        filtered = filtered[filtered["rotacion"] == f_rot]
    if f_rubro != "Todos" and "rubro" in filtered.columns:
        filtered = filtered[filtered["rubro"] == f_rubro]
    if f_marca != "Todas" and "marca" in filtered.columns:
        filtered = filtered[filtered["marca"] == f_marca]
    if f_search:
        mask = filtered["codigo"].str.contains(f_search, case=False, na=False) | \
               filtered["descripcion"].str.contains(f_search, case=False, na=False)
        filtered = filtered[mask]

    st.caption(f"Mostrando **{len(filtered):,}** de **{len(base):,}** productos")

    # ── KPIs ──────────────────────────────────────────────────────────────────
    st.subheader("📊 Métricas del período")
    k = st.columns(5)

    cob_finita = base.loc[base["cobertura_meses"] < np.inf, "cobertura_meses"]

    k[0].metric("Total productos",       f"{len(base):,}")
    k[1].metric("Unidades vendidas",      f"{base['total_vendido'].sum():,.0f}")
    k[2].metric("Productos con venta",    f"{(base['total_vendido'] > 0).sum():,}",
                f"{(base['total_vendido'] > 0).mean()*100:.0f}% del catálogo")
    k[3].metric("Cobertura mediana",
                f"{cob_finita.median():.1f} m" if len(cob_finita) else "—")
    k[4].metric("Meses cargados",         f"{len(meses_cargados)} / {N_MESES_TOTAL}")

    if "precio" in base.columns and "valor_stock" in base.columns:
        k2 = st.columns(3)
        k2[0].metric("Valor total stock",   f"$ {base['valor_stock'].sum():,.0f}")
        k2[1].metric("Valor sobrestock",
                     f"$ {base.loc[base['estado_stock'].str.startswith('Sobrestock'), 'valor_stock'].sum():,.0f}")
        k2[2].metric("Valor vendido total", f"$ {base['valor_vendido'].sum():,.0f}")

    st.divider()

    # ── Tabs ──────────────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📋 Base completa",
        "🚨 Alertas & decisiones",
        "📈 Análisis de ventas",
        "🏆 Ranking de rotación",
        "📤 Exportar",
    ])

    # ═══════════════════════════════════════════════════════════════════════════
    #  TAB 1 — Full consolidated base
    # ═══════════════════════════════════════════════════════════════════════════
    with tab1:
        st.subheader("Base consolidada de productos")

        show_cols = ["codigo", "descripcion"]
        for opt in ("rubro", "marca"):
            if opt in filtered.columns:
                show_cols.append(opt)
        show_cols += mes_cols + [
            "total_vendido", "promedio_mensual", "stock_actual",
            "cobertura_meses", "rotacion", "estado_stock",
        ]
        if "precio" in filtered.columns:
            show_cols += ["precio", "valor_stock"]

        df_show = filtered[show_cols].copy()
        df_show["cobertura_meses"] = df_show["cobertura_meses"].apply(_fmt_cob)

        styled = (
            df_show.style
            .map(_style_estado, subset=["estado_stock"])
            .map(_style_rot,    subset=["rotacion"])
            .format(
                {
                    "total_vendido":    "{:.0f}",
                    "promedio_mensual": "{:.1f}",
                    "stock_actual":     "{:.0f}",
                    **({col: "{:.0f}" for col in mes_cols}),
                },
                na_rep="0",
            )
        )
        st.dataframe(styled, use_container_width=True, height=520)

    # ═══════════════════════════════════════════════════════════════════════════
    #  TAB 2 — Alerts & decisions
    # ═══════════════════════════════════════════════════════════════════════════
    with tab2:
        left, right = st.columns(2)

        # ── Left: urgent buy + watch ──────────────────────────────────────────
        with left:
            st.markdown("### 🔴 Comprar YA — quiebre o crítico con rotación alta/media")
            df_buy = filtered[
                filtered["rotacion"].isin(["Alta", "Media"]) &
                filtered["estado_stock"].isin(["Quiebre", "Crítico"])
            ].sort_values("prioridad_compra", ascending=False)

            if df_buy.empty:
                st.success("No hay productos urgentes con los filtros actuales.")
            else:
                cols = ["codigo", "descripcion", "promedio_mensual",
                        "stock_actual", "cobertura_meses", "estado_stock"]
                cols = [c for c in cols if c in df_buy.columns]
                df_buy_show = df_buy[cols].copy()
                df_buy_show["cobertura_meses"] = df_buy_show["cobertura_meses"].apply(_fmt_cob)
                st.dataframe(
                    df_buy_show.style.map(_style_estado, subset=["estado_stock"]),
                    use_container_width=True,
                )

            st.markdown("### 🟡 Vigilar — normal pero con cobertura < 2 meses")
            df_watch = filtered[
                filtered["rotacion"].isin(["Alta", "Media"]) &
                (filtered["estado_stock"] == "Normal") &
                (filtered["cobertura_meses"] < 2)
            ].sort_values("cobertura_meses")

            if df_watch.empty:
                st.success("No hay productos en zona de alerta próxima.")
            else:
                cols = ["codigo", "descripcion", "promedio_mensual",
                        "stock_actual", "cobertura_meses"]
                cols = [c for c in cols if c in df_watch.columns]
                df_watch_show = df_watch[cols].copy()
                df_watch_show["cobertura_meses"] = df_watch_show["cobertura_meses"].apply(_fmt_cob)
                st.dataframe(df_watch_show, use_container_width=True)

        # ── Right: overstock + no movement ───────────────────────────────────
        with right:
            st.markdown("### 🟣 Sobrestock — capital inmovilizado")
            df_over = filtered[
                filtered["estado_stock"].str.startswith("Sobrestock")
            ].sort_values("stock_actual", ascending=False)

            if df_over.empty:
                st.success("Sin sobrestock con los filtros actuales.")
            else:
                cols = ["codigo", "descripcion", "stock_actual",
                        "promedio_mensual", "cobertura_meses", "estado_stock"]
                if "valor_stock" in df_over.columns:
                    cols.append("valor_stock")
                cols = [c for c in cols if c in df_over.columns]
                df_over_show = df_over[cols].copy()
                df_over_show["cobertura_meses"] = df_over_show["cobertura_meses"].apply(_fmt_cob)
                st.dataframe(
                    df_over_show.style.map(_style_estado, subset=["estado_stock"]),
                    use_container_width=True,
                )

            st.markdown(
                f"### ⚫ Sin venta en {len(meses_cargados) or N_MESES_TOTAL} meses — "
                "con stock disponible"
            )
            df_dead = filtered[
                (filtered["total_vendido"] == 0) & (filtered["stock_actual"] > 0)
            ].sort_values("stock_actual", ascending=False)

            if df_dead.empty:
                st.success("Todos los productos con stock tuvieron al menos alguna venta.")
            else:
                cols = ["codigo", "descripcion", "stock_actual"]
                if "rubro" in df_dead.columns:   cols.insert(2, "rubro")
                if "valor_stock" in df_dead.columns: cols.append("valor_stock")
                cols = [c for c in cols if c in df_dead.columns]
                st.info(f"{len(df_dead)} productos")
                st.dataframe(df_dead[cols], use_container_width=True)

        # ── Decision matrix ───────────────────────────────────────────────────
        st.divider()
        st.markdown("### 🧭 Matriz de decisión")
        matrix = pd.DataFrame([
            ["Quiebre / Crítico + Alta rotación",  "🔴 Comprar de inmediato",     "URGENTE"],
            ["Quiebre / Crítico + Media rotación", "🟠 Comprar pronto",            "ALTA"],
            ["Normal + Alta rotación (cob < 2m)",  "✅ Monitorear semanalmente",   "MEDIA"],
            ["Sobrestock + Sin venta",             "🟣 Revisar precio / liquidar", "BAJA"],
            ["Sobrestock + Baja rotación",         "🟡 No recomprar",              "BAJA"],
            ["Sin stock + Sin venta",              "⚫ Evaluar baja del catálogo", "SIN ACCIÓN"],
        ], columns=["Situación", "Acción recomendada", "Prioridad"])
        st.dataframe(matrix, use_container_width=True, hide_index=True)

    # ═══════════════════════════════════════════════════════════════════════════
    #  TAB 3 — Sales analysis
    # ═══════════════════════════════════════════════════════════════════════════
    with tab3:
        if not mes_cols:
            st.warning("No hay archivos de ventas cargados.")
        else:
            # Monthly totals bar chart
            df_monthly = pd.DataFrame(
                {"Mes": mes_cols, "Unidades": [filtered[m].sum() for m in mes_cols]}
            )
            fig_bar = px.bar(
                df_monthly, x="Mes", y="Unidades",
                title="Unidades vendidas por mes (filtro aplicado)",
                color="Unidades", color_continuous_scale="Blues",
                text="Unidades",
            )
            fig_bar.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
            fig_bar.update_layout(coloraxis_showscale=False)
            st.plotly_chart(fig_bar, use_container_width=True)

            col_a, col_b = st.columns(2)

            with col_a:
                rot_vc = filtered["rotacion"].value_counts().reset_index()
                rot_vc.columns = ["Rotación", "Productos"]
                fig_pie = px.pie(
                    rot_vc, values="Productos", names="Rotación",
                    title="Distribución por nivel de rotación",
                    color="Rotación",
                    color_discrete_map={
                        "Alta": "#27ae60", "Media": "#f39c12",
                        "Baja": "#e74c3c", "Sin movimiento": "#95a5a6",
                    },
                )
                st.plotly_chart(fig_pie, use_container_width=True)

            with col_b:
                est_vc = filtered["estado_stock"].value_counts().reset_index()
                est_vc.columns = ["Estado", "Productos"]
                fig_est = px.bar(
                    est_vc, x="Estado", y="Productos",
                    title="Productos por estado de stock",
                    color="Estado",
                    color_discrete_map={
                        "Quiebre": "#e74c3c", "Crítico": "#e67e22",
                        "Normal": "#27ae60", "Sobrestock": "#9b59b6",
                        "Sobrestock (sin venta)": "#8e44ad",
                        "Sin stock / Sin venta": "#95a5a6",
                    },
                )
                fig_est.update_layout(showlegend=False)
                st.plotly_chart(fig_est, use_container_width=True)

            # Heatmap top 30
            st.subheader("Mapa de calor — Top 30 productos")
            top30 = (
                filtered[filtered["total_vendido"] > 0]
                .nlargest(30, "total_vendido")[["descripcion"] + mes_cols]
                .set_index("descripcion")
            )
            if not top30.empty:
                fig_heat = px.imshow(
                    top30,
                    labels=dict(x="Mes", y="Producto", color="Unidades"),
                    color_continuous_scale="YlOrRd",
                    title="Ventas mensuales — Top 30 por total acumulado",
                    aspect="auto",
                )
                st.plotly_chart(fig_heat, use_container_width=True)

    # ═══════════════════════════════════════════════════════════════════════════
    #  TAB 4 — Ranking
    # ═══════════════════════════════════════════════════════════════════════════
    with tab4:
        top_n = st.slider("Top N productos", 10, 100, 20, step=5)

        df_rank = (
            filtered[filtered["total_vendido"] > 0]
            .nlargest(top_n, "total_vendido")
            .copy()
        )
        df_rank.insert(0, "#", range(1, len(df_rank) + 1))

        # Horizontal bar chart
        fig_hbar = px.bar(
            df_rank.head(20),
            x="total_vendido",
            y="descripcion",
            orientation="h",
            color="rotacion",
            color_discrete_map={
                "Alta": "#27ae60", "Media": "#f39c12", "Baja": "#e74c3c",
            },
            title=f"Top 20 productos — unidades vendidas en el período",
            labels={"total_vendido": "Total vendido", "descripcion": ""},
        )
        fig_hbar.update_layout(yaxis={"categoryorder": "total ascending"}, height=600)
        st.plotly_chart(fig_hbar, use_container_width=True)

        # Scatter: rotation vs. coverage
        st.subheader("Rotación vs. Cobertura actual")
        df_sc = filtered[
            (filtered["total_vendido"] > 0) & (filtered["cobertura_meses"] < 99)
        ].copy()

        if not df_sc.empty:
            thresholds = st.session_state.get("thresholds", {})
            fig_sc = px.scatter(
                df_sc,
                x="promedio_mensual",
                y="cobertura_meses",
                size=df_sc["stock_actual"].clip(lower=1),
                color="estado_stock",
                hover_data=["codigo", "descripcion"],
                color_discrete_map={
                    "Quiebre": "#e74c3c", "Crítico": "#e67e22",
                    "Normal": "#27ae60", "Sobrestock": "#9b59b6",
                },
                title="Promedio mensual vendido vs. Cobertura en meses  (burbuja = stock actual)",
                labels={
                    "promedio_mensual": "Promedio mensual (u/mes)",
                    "cobertura_meses":  "Cobertura (meses)",
                },
            )
            crit = thresholds.get("cob_critica", DEFAULT_COB_CRITICA)
            sobr = thresholds.get("cob_sobrestock", DEFAULT_COB_SOBRESTOCK)
            fig_sc.add_hline(y=crit, line_dash="dash", line_color="#e74c3c",
                             annotation_text=f"Crítico ({crit}m)")
            fig_sc.add_hline(y=sobr, line_dash="dash", line_color="#9b59b6",
                             annotation_text=f"Sobrestock ({sobr}m)")
            st.plotly_chart(fig_sc, use_container_width=True)

        # Ranking table
        rank_cols = ["#", "codigo", "descripcion"]
        for opt in ("rubro", "marca"):
            if opt in df_rank.columns:
                rank_cols.append(opt)
        rank_cols += ["total_vendido", "promedio_mensual", "stock_actual",
                      "cobertura_meses", "rotacion", "estado_stock"]
        rank_cols = [c for c in rank_cols if c in df_rank.columns]

        df_rank_show = df_rank[rank_cols].copy()
        df_rank_show["cobertura_meses"] = df_rank_show["cobertura_meses"].apply(_fmt_cob)

        st.dataframe(
            df_rank_show.style
            .map(_style_estado, subset=["estado_stock"])
            .map(_style_rot,    subset=["rotacion"])
            .format({"total_vendido": "{:.0f}", "promedio_mensual": "{:.1f}",
                     "stock_actual": "{:.0f}"}, na_rep="0"),
            use_container_width=True,
        )

    # ═══════════════════════════════════════════════════════════════════════════
    #  TAB 5 — Export
    # ═══════════════════════════════════════════════════════════════════════════
    with tab5:
        date_str = datetime.today().strftime("%Y%m%d")

        st.markdown("#### Base completa (con filtros aplicados)")
        st.download_button(
            "⬇️ Descargar base filtrada",
            data=to_excel(filtered),
            file_name=f"sf_base_{date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.markdown("#### Solo alertas críticas (quiebre + crítico)")
        df_critical = base[base["estado_stock"].isin(["Quiebre", "Crítico"])].copy()
        st.download_button(
            "⬇️ Descargar alertas críticas",
            data=to_excel(df_critical),
            file_name=f"sf_alertas_{date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.markdown("#### Lista de compras sugerida")
        df_compras = base[
            base["rotacion"].isin(["Alta", "Media"]) &
            base["estado_stock"].isin(["Quiebre", "Crítico"])
        ].copy().sort_values("prioridad_compra", ascending=False)

        if not df_compras.empty:
            df_compras["cant_sugerida_compra"] = (
                df_compras["promedio_mensual"] * 2 - df_compras["stock_actual"]
            ).clip(lower=0).round(0)
            st.download_button(
                "⬇️ Descargar lista de compras",
                data=to_excel(df_compras),
                file_name=f"sf_compras_{date_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("No hay productos que requieran compra urgente.")

        st.markdown("#### Sobrestock (para gestión de capital inmovilizado)")
        df_over_exp = base[base["estado_stock"].str.startswith("Sobrestock")].copy()
        if not df_over_exp.empty:
            st.download_button(
                "⬇️ Descargar sobrestock",
                data=to_excel(df_over_exp),
                file_name=f"sf_sobrestock_{date_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        if orphans:
            st.divider()
            st.markdown("#### ⚠️ Códigos en ventas sin match en stock maestro")
            st.warning(
                f"{len(orphans)} registros de ventas no encontraron su código "
                "en el stock maestro. Revisá si hay variaciones de código."
            )
            df_orp = pd.DataFrame(orphans)
            st.dataframe(df_orp, use_container_width=True)
            st.download_button(
                "⬇️ Descargar códigos huérfanos",
                data=to_excel(df_orp),
                file_name=f"sf_huerfanos_{date_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    main()
