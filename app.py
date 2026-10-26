"""Plataforma Analítica Executiva (C-Level) - Streamlit + Plotly.

Executar com:  streamlit run app.py
Requer o ficheiro ``financeiro.xlsx`` na mesma pasta deste script.
"""
from __future__ import annotations

import io
import math
from datetime import timedelta
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

DATA_PATH = Path(__file__).parent / "financeiro.xlsx"
NO_DISCOUNT = "Nenhum"
DISCOUNT_ORDER = [NO_DISCOUNT, "Low", "Medium", "High"]
MONTHS_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
             "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
ISO3 = {"Canada": "CAN", "Germany": "DEU", "France": "FRA", "Mexico": "MEX",
        "United States of America": "USA"}

# Paleta corporativa: ardósia, cinzento, esmeralda (ganho) e vermelho (perda)
SLATE = "#334E68"
SLATE_LIGHT = "#627D98"
GRAY = "#9FB3C8"
EMERALD = "#059669"
RED = "#B91C1C"
AMBER = "#D97706"
CAT_COLORS = [SLATE, EMERALD, AMBER, "#7C5C8E", SLATE_LIGHT, "#0E7490"]
GRID = "rgba(100,116,139,0.18)"

FILTER_KEYS = {"ano": "f_ano", "pais": "f_pais", "segmento": "f_seg", "produto": "f_prod"}


# --------------------------------------------------------------------------
# Dados
# --------------------------------------------------------------------------
@st.cache_data(show_spinner="A carregar dados...")
def load_data(path: Path = DATA_PATH) -> pd.DataFrame:
    """Lê e higieniza a base financeira."""
    df = pd.read_excel(path)
    df.columns = df.columns.str.strip()
    for col in ("segmento", "pais", "produto", "faixa_desconto", "mes"):
        df[col] = df[col].astype("string").str.strip()
    df["faixa_desconto"] = (
        df["faixa_desconto"].fillna(NO_DISCOUNT).replace({"None": NO_DISCOUNT})
    )
    df["data"] = pd.to_datetime(df["data"])
    df["periodo"] = df["data"].dt.to_period("M").dt.to_timestamp()
    return df.sort_values("data").reset_index(drop=True)


def safe_div(a: float, b: float) -> float:
    return a / b if b else 0.0


# --------------------------------------------------------------------------
# Formatação (padrão pt)
# --------------------------------------------------------------------------
def _swap_sep(text: str) -> str:
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


def brl(value: float, abbreviate: bool = True) -> str:
    """Valor em R$, abreviado (K/M/B) ou completo com separadores de milhar."""
    sign = "-" if value < 0 else ""
    v = abs(value)
    if abbreviate:
        for limit, suffix in ((1e9, "B"), (1e6, "M"), (1e3, "K")):
            if v >= limit:
                return f"{sign}R$ {_swap_sep(f'{v / limit:,.1f}')}{suffix}"
    return f"{sign}R$ {_swap_sep(f'{v:,.2f}')}"


def num(value: float, decimals: int = 0) -> str:
    return _swap_sep(f"{value:,.{decimals}f}")


def pct(value: float) -> str:
    return f"{_swap_sep(f'{value:.1f}')}%"


def delta_pct(cur: float, prev: float) -> str | None:
    """Variação relativa (%) face ao valor de referência."""
    if not prev:
        return None
    return f"{_swap_sep(f'{(cur / prev - 1) * 100:+.1f}')}%"


def delta_pp(cur: float, prev: float) -> str:
    return f"{_swap_sep(f'{cur - prev:+.1f}')} p.p."


# --------------------------------------------------------------------------
# Estilo
# --------------------------------------------------------------------------
def inject_css() -> None:
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
        html, body, [class*="css"] {font-family: 'Inter', sans-serif;}
        .block-container {padding-top: 1.5rem;}
        [data-testid="stMetric"] {
            border: 1px solid rgba(100,116,139,0.3);
            border-left: 4px solid #334E68;
            background: rgba(100,116,139,0.07);
            border-radius: 8px; padding: 12px 14px;
        }
        [data-testid="stMetricLabel"] {text-transform: uppercase;
            letter-spacing: .04em; font-size: .75rem; opacity: .75;}
        [data-testid="stMetricValue"] {font-weight: 700;}
        </style>
        """,
        unsafe_allow_html=True,
    )


def style_fig(fig: go.Figure, height: int = 420, hover: str | None = "x unified") -> go.Figure:
    fig.update_layout(
        height=height, margin=dict(l=10, r=10, t=50, b=10),
        legend=dict(orientation="h", y=-0.18), font=dict(size=13),
        hovermode=hover, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
    )
    fig.update_xaxes(gridcolor=GRID)
    fig.update_yaxes(gridcolor=GRID)
    return fig


# --------------------------------------------------------------------------
# Sidebar: filtros encadeados com seleção múltipla
# --------------------------------------------------------------------------
def _set_all(key: str, options: list) -> None:
    st.session_state[key] = list(options)


def _clear(key: str) -> None:
    st.session_state[key] = []


def multi_filter(label: str, options: list, key: str) -> list:
    """Multiselect com botões Todos/Limpar.

    Se, na execução anterior, todas as opções estavam selecionadas, novas
    opções (por alteração de um filtro anterior) entram selecionadas; caso
    contrário apenas se descartam valores que já não existem.
    """
    state = st.session_state
    prev_opts = state.get(f"{key}_opts")
    if key not in state:
        state[key] = list(options)
    else:
        was_all = prev_opts is not None and set(state[key]) == set(prev_opts)
        state[key] = (list(options) if was_all
                      else [v for v in state[key] if v in options])
    state[f"{key}_opts"] = list(options)

    st.sidebar.multiselect(label, options, key=key)
    c1, c2 = st.sidebar.columns(2)
    c1.button("Todos", key=f"{key}_all", on_click=_set_all, args=(key, options),
              width="stretch")
    c2.button("Limpar", key=f"{key}_none", on_click=_clear, args=(key,),
              width="stretch")
    return state[key]


def apply_filters(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, tuple]:
    """Renderiza a sidebar e devolve (filtrado, período anterior, (início, fim))."""
    st.sidebar.header("Filtros")
    dmin, dmax = df["data"].min().date(), df["data"].max().date()
    period = st.sidebar.date_input("Intervalo de datas", value=(dmin, dmax),
                                   min_value=dmin, max_value=dmax)
    start, end = period if isinstance(period, tuple) and len(period) == 2 else (dmin, dmax)
    dates = df["data"].dt.date

    out = df[(dates >= start) & (dates <= end)]
    anos = multi_filter("Ano", sorted(df["ano"].unique().tolist()), FILTER_KEYS["ano"])
    out = out[out["ano"].isin(anos)]
    paises = multi_filter("País", sorted(out["pais"].unique().tolist()), FILTER_KEYS["pais"])
    out = out[out["pais"].isin(paises)]
    segs = multi_filter("Segmento", sorted(out["segmento"].unique().tolist()),
                        FILTER_KEYS["segmento"])
    out = out[out["segmento"].isin(segs)]
    prods = multi_filter("Produto", sorted(out["produto"].unique().tolist()),
                         FILTER_KEYS["produto"])
    out = out[out["produto"].isin(prods)]

    # período imediatamente anterior (mesma duração e dimensões), só se a base o cobrir
    prev = df.iloc[0:0]
    if not out.empty:
        first, last = out["data"].min().date(), out["data"].max().date()
        span = (last - first).days + 1
        prev_end = first - timedelta(days=1)
        prev_start = prev_end - timedelta(days=span - 1)
        if prev_start >= dmin:
            prev = df[(dates >= prev_start) & (dates <= prev_end) & df["pais"].isin(paises)
                      & df["segmento"].isin(segs) & df["produto"].isin(prods)]

    st.sidebar.divider()
    st.sidebar.metric("Registos filtrados", f"{num(len(out))} / {num(len(df))}")
    st.sidebar.progress(min(len(out) / len(df), 1.0))
    return out, prev, (start, end)


# --------------------------------------------------------------------------
# Métricas
# --------------------------------------------------------------------------
def compute_kpis(df: pd.DataFrame) -> dict:
    venda, lucro = df["venda"].sum(), df["lucro"].sum()
    bruta, desc = df["venda_bruta"].sum(), df["desconto"].sum()
    return {
        "venda": venda, "lucro": lucro, "margem": safe_div(lucro, venda) * 100,
        "unidades": df["unidades_vendidas"].sum(), "ticket": safe_div(venda, len(df)),
        "taxa_desc": safe_div(desc, bruta) * 100, "desconto": desc, "n": len(df),
    }


def render_kpis(cur: dict, prev: dict | None) -> None:
    def d(key, kind="pct"):
        if prev is None or not prev["n"]:
            return None
        return delta_pct(cur[key], prev[key]) if kind == "pct" else delta_pp(cur[key], prev[key])

    r1 = st.columns(3)
    r1[0].metric("Faturamento líquido", brl(cur["venda"]), d("venda"),
                 help=brl(cur["venda"], False))
    r1[1].metric("Lucro operacional", brl(cur["lucro"]), d("lucro"),
                 help=brl(cur["lucro"], False))
    r1[2].metric("Margem operacional", pct(cur["margem"]), d("margem", "pp"),
                 help="Lucro / Faturamento líquido")
    r2 = st.columns(3)
    r2[0].metric("Unidades vendidas", num(cur["unidades"]), d("unidades"))
    r2[1].metric("Ticket médio / transação", brl(cur["ticket"]), d("ticket"))
    r2[2].metric("Taxa média de desconto", pct(cur["taxa_desc"]), d("taxa_desc", "pp"),
                 delta_color="inverse", help="Desconto / Venda bruta")
    if prev is None or not prev["n"]:
        st.caption("Deltas indisponíveis: a base não cobre por completo o período anterior "
                   "de igual duração.")
    else:
        st.caption("Deltas face ao período imediatamente anterior de igual duração.")


# --------------------------------------------------------------------------
# Separador 1: Cockpit
# --------------------------------------------------------------------------
def chart_timeseries(df: pd.DataFrame) -> go.Figure:
    m = df.groupby("periodo", as_index=False)[["venda", "lucro"]].sum()
    m["margem"] = m["lucro"] / m["venda"] * 100
    fig = go.Figure()
    fig.add_bar(x=m["periodo"], y=m["venda"], name="Faturamento", marker_color=SLATE,
                opacity=0.85, hovertemplate="Faturamento: R$ %{y:,.0f}<extra></extra>")
    fig.add_scatter(x=m["periodo"], y=m["lucro"], name="Lucro", mode="lines+markers",
                    line=dict(color=EMERALD, width=3),
                    hovertemplate="Lucro: R$ %{y:,.0f}<extra></extra>")
    fig.add_scatter(x=m["periodo"], y=m["margem"], name="Margem (%)", yaxis="y2",
                    mode="lines", line=dict(color=AMBER, width=2, dash="dot"),
                    hovertemplate="Margem: %{y:.1f}%<extra></extra>")
    fig.update_layout(
        title="Faturamento vs. Lucro e retenção de margem",
        yaxis=dict(title="R$", tickformat="~s"),
        yaxis2=dict(title="Margem (%)", overlaying="y", side="right",
                    showgrid=False, ticksuffix="%"),
    )
    return style_fig(fig)


def chart_seasonality(df: pd.DataFrame) -> go.Figure:
    g = df.groupby(["ano", "mes_nr"], as_index=False)["venda"].sum()
    g["mes_lbl"] = g["mes_nr"].map(lambda i: MONTHS_PT[i - 1])
    g["ano"] = g["ano"].astype(str)
    fig = px.line(g, x="mes_lbl", y="venda", color="ano", markers=True,
                  color_discrete_sequence=CAT_COLORS,
                  category_orders={"mes_lbl": MONTHS_PT},
                  title="Sazonalidade: faturamento por mês do ano")
    fig.update_traces(hovertemplate="R$ %{y:,.0f}")
    fig.update_layout(yaxis=dict(title="R$", tickformat="~s"), xaxis_title=None,
                      legend_title="Ano")
    return style_fig(fig, 340)


def build_alerts(df: pd.DataFrame, max_rate: float) -> pd.DataFrame:
    """Pares (produto, segmento) com margem negativa ou desconto excessivo."""
    g = df.groupby(["produto", "segmento"]).agg(
        venda=("venda", "sum"), lucro=("lucro", "sum"),
        bruta=("venda_bruta", "sum"), desconto=("desconto", "sum"),
    ).reset_index()
    g["margem"] = g["lucro"] / g["venda"] * 100
    g["taxa_desc"] = g["desconto"] / g["bruta"] * 100

    def label(row) -> str:
        found = []
        if row["margem"] < 0:
            found.append("Margem negativa")
        if row["taxa_desc"] > max_rate:
            found.append("Desconto excessivo")
        return " + ".join(found)

    g["alerta"] = g.apply(label, axis=1)
    return g[g["alerta"] != ""].sort_values("margem")


def render_alerts(df: pd.DataFrame) -> None:
    st.subheader("🚨 Alertas e anomalias")
    max_rate = st.slider("Limite de taxa de desconto considerada excessiva (%)",
                         2.0, 15.0, 8.0, 0.5)
    alerts = build_alerts(df, max_rate)
    neg_tx = int((df["lucro"] < 0).sum())
    seg = df.groupby("segmento")[["lucro", "venda"]].sum()
    seg_neg = seg[seg["lucro"] < 0]

    for name, row in seg_neg.iterrows():
        st.error(f"Segmento **{name}** opera com prejuízo: lucro {brl(row['lucro'])} "
                 f"({pct(row['lucro'] / row['venda'] * 100)} de margem).")
    if neg_tx:
        st.warning(f"**{num(neg_tx)}** transações ({pct(neg_tx / len(df) * 100)}) "
                   "com lucro negativo no período.")
    if alerts.empty and seg_neg.empty and not neg_tx:
        st.success("Nenhuma anomalia detetada no período filtrado.")
    if not alerts.empty:
        st.dataframe(
            alerts[["produto", "segmento", "alerta", "venda", "lucro", "margem", "taxa_desc"]],
            hide_index=True, width="stretch",
            column_config={
                "produto": "Produto", "segmento": "Segmento", "alerta": "Alerta",
                "venda": st.column_config.NumberColumn("Faturamento", format="R$ %.0f"),
                "lucro": st.column_config.NumberColumn("Lucro", format="R$ %.0f"),
                "margem": st.column_config.NumberColumn("Margem", format="%.1f%%"),
                "taxa_desc": st.column_config.NumberColumn("Taxa desc.", format="%.1f%%"),
            },
        )


def tab_cockpit(df: pd.DataFrame, prev: pd.DataFrame) -> None:
    render_kpis(compute_kpis(df), compute_kpis(prev) if len(prev) else None)
    st.write("")
    st.plotly_chart(chart_timeseries(df), width="stretch")
    st.plotly_chart(chart_seasonality(df), width="stretch")
    render_alerts(df)


# --------------------------------------------------------------------------
# Separador 2: Matriz BCG
# --------------------------------------------------------------------------
def quadrant_name(vol: float, margem: float, vol_cut: float, mar_cut: float) -> str:
    if vol >= vol_cut and margem >= mar_cut:
        return "Produtos Estrela"
    if vol < vol_cut and margem >= mar_cut:
        return "Operações de Alto Rendimento"
    if vol >= vol_cut:
        return "Oportunidades de Volume"
    return "Críticos / Sub-otimizados"


def portfolio_table(df: pd.DataFrame, color_by: str) -> tuple[pd.DataFrame, float, float]:
    g = df.groupby(["produto", color_by], as_index=False).agg(
        unidades=("unidades_vendidas", "sum"), venda=("venda", "sum"), lucro=("lucro", "sum"))
    g["margem"] = g["lucro"] / g["venda"] * 100
    vol_cut, mar_cut = g["unidades"].mean(), g["margem"].mean()
    g["quadrante"] = [quadrant_name(u, m, vol_cut, mar_cut)
                      for u, m in zip(g["unidades"], g["margem"])]
    return g, vol_cut, mar_cut


def chart_bcg(g: pd.DataFrame, color_by: str, vol_cut: float, mar_cut: float) -> go.Figure:
    plot = g[g["venda"] > 0]  # bolhas exigem tamanho positivo
    fig = px.scatter(
        plot, x="unidades", y="margem", size="venda", color=color_by, text="produto",
        size_max=48, color_discrete_sequence=CAT_COLORS,
        custom_data=["produto", color_by, "venda", "lucro", "quadrante"],
    )
    fig.update_traces(
        textposition="top center", marker=dict(opacity=0.75, line=dict(width=1, color="white")),
        hovertemplate=("<b>%{customdata[0]}</b> · %{customdata[1]}"
                       "<br>Unidades: %{x:,.0f}<br>Margem: %{y:.1f}%"
                       "<br>Faturamento: R$ %{customdata[2]:,.0f}"
                       "<br>Lucro: R$ %{customdata[3]:,.0f}"
                       "<br><i>%{customdata[4]}</i><extra></extra>"),
    )
    fig.add_vline(x=vol_cut, line=dict(color=SLATE_LIGHT, dash="dash"))
    fig.add_hline(y=mar_cut, line=dict(color=SLATE_LIGHT, dash="dash"))
    fig.add_hline(y=0, line=dict(color=RED, width=1))
    corners = [(0.01, 0.99, "left", "top", "Operações de Alto Rendimento", EMERALD),
               (0.99, 0.99, "right", "top", "Produtos Estrela", EMERALD),
               (0.01, 0.01, "left", "bottom", "Críticos / Sub-otimizados", RED),
               (0.99, 0.01, "right", "bottom", "Oportunidades de Volume", AMBER)]
    for x, y, xa, ya, text, color in corners:
        fig.add_annotation(xref="paper", yref="paper", x=x, y=y, xanchor=xa, yanchor=ya,
                           text=f"<b>{text}</b>", showarrow=False,
                           font=dict(color=color, size=13))
    fig.update_layout(title="Matriz de portfólio: volume × margem (tamanho = faturamento)",
                      xaxis_title="Unidades vendidas", yaxis_title="Margem de lucro (%)",
                      yaxis_ticksuffix="%", legend_title=None)
    return style_fig(fig, 580, hover="closest")


def tab_portfolio(df: pd.DataFrame) -> None:
    c1, c2 = st.columns([1, 2])
    label = c1.radio("Cor das bolhas", ["Segmento", "País"], horizontal=True)
    color_by = "segmento" if label == "Segmento" else "pais"
    g, vol_cut, mar_cut = portfolio_table(df, color_by)
    c2.caption(f"Linhas de corte = médias das bolhas: {num(vol_cut)} unidades e "
               f"{pct(mar_cut)} de margem. Cada bolha é um par produto × {label.lower()}.")
    st.plotly_chart(chart_bcg(g, color_by, vol_cut, mar_cut), width="stretch")

    counts = g.groupby("quadrante").agg(
        itens=("produto", "size"), faturamento=("venda", "sum"), lucro=("lucro", "sum"))
    st.subheader("Resumo por quadrante")
    st.dataframe(counts.sort_values("faturamento", ascending=False), width="stretch",
                 column_config={
                     "quadrante": "Quadrante",
                     "itens": "Nº de bolhas",
                     "faturamento": st.column_config.NumberColumn("Faturamento", format="R$ %.0f"),
                     "lucro": st.column_config.NumberColumn("Lucro", format="R$ %.0f"),
                 })


# --------------------------------------------------------------------------
# Separador 3: Simulador What-If
# --------------------------------------------------------------------------
def project(base: dict, vol: float = 0, price: float = 0, disc: float = 0,
            cost: float = 0) -> dict:
    """Projeta faturamento/lucro a partir de agregados.

    Os argumentos são variações relativas (0.05 = +5%). O desconto acompanha
    volume e preço (é % da venda bruta) e a política de descontos altera a
    taxa; o custo varia com o volume.
    """
    gross = base["bruta"] * (1 + vol) * (1 + price)
    discount = base["desconto"] * (1 + vol) * (1 + price) * (1 + disc)
    venda = gross - discount
    custo = base["custo"] * (1 + vol) * (1 + cost)
    lucro = venda - custo
    return {"venda": venda, "custo": custo, "lucro": lucro,
            "margem": safe_div(lucro, venda) * 100}


def chart_waterfall(base: dict, vol: float, price: float, disc: float, cost: float) -> go.Figure:
    chain = [("Volume (elasticidade)", dict(vol=vol)),
             ("Preço", dict(vol=vol, price=price)),
             ("Descontos", dict(vol=vol, price=price, disc=disc)),
             ("Custos", dict(vol=vol, price=price, disc=disc, cost=cost))]
    prev_val = project(base)["lucro"]
    labels, values, measures = ["Lucro atual"], [prev_val], ["absolute"]
    for name, kwargs in chain:
        val = project(base, **kwargs)["lucro"]
        labels.append(name)
        values.append(val - prev_val)
        measures.append("relative")
        prev_val = val
    labels.append("Lucro projetado")
    values.append(prev_val)
    measures.append("total")

    fig = go.Figure(go.Waterfall(
        x=labels, y=values, measure=measures,
        text=[brl(v) for v in values], textposition="outside",
        increasing=dict(marker=dict(color=EMERALD)),
        decreasing=dict(marker=dict(color=RED)),
        totals=dict(marker=dict(color=SLATE)),
        connector=dict(line=dict(color=GRAY)),
        hovertemplate="%{x}: R$ %{y:,.0f}<extra></extra>",
    ))
    fig.update_layout(title="Cascata: do lucro atual ao projetado",
                      yaxis=dict(title="R$", tickformat="~s"), showlegend=False)
    return style_fig(fig, 420, hover="closest")


def chart_compare(cur: dict, proj: dict) -> go.Figure:
    cats, keys = ["Faturamento", "Custo", "Lucro"], ["venda", "custo", "lucro"]
    fig = go.Figure()
    fig.add_bar(x=cats, y=[cur[k] for k in keys], name="Atual", marker_color=GRAY,
                hovertemplate="R$ %{y:,.0f}<extra>Atual</extra>")
    fig.add_bar(x=cats, y=[proj[k] for k in keys], name="Projetado", marker_color=SLATE,
                hovertemplate="R$ %{y:,.0f}<extra>Projetado</extra>")
    fig.update_layout(title="Atual vs. projetado", barmode="group",
                      yaxis=dict(title="R$", tickformat="~s"))
    return style_fig(fig, 420, hover="closest")


def tab_simulator(df: pd.DataFrame) -> None:
    st.markdown("Ajuste as alavancas para estimar o impacto sobre o resultado do "
                "recorte atual (filtros da barra lateral).")
    c1, c2, c3, c4 = st.columns(4)
    price = c1.slider("Preço médio de venda (%)", -20, 20, 0, 1) / 100
    cost = c2.slider("Custo de fabrico/operação (%)", -20, 20, 0, 1) / 100
    disc = c3.slider("Política de descontos (%)", -50, 50, 0, 5,
                     help="Altera a taxa de desconto: -50% = metade do desconto atual.") / 100
    elasticity = c4.slider("Elasticidade-preço da procura", -3.0, 0.0, 0.0, 0.1,
                           help="0 = volume inalterado. -1 = +10% de preço reduz o "
                                "volume em 10%.")
    vol = elasticity * price

    base = {"bruta": df["venda_bruta"].sum(), "desconto": df["desconto"].sum(),
            "custo": df["custo"].sum()}
    cur = project(base)
    proj = project(base, vol, price, disc, cost)

    lucro_delta = delta_pct(proj["lucro"], cur["lucro"]) or "0,0%"
    m = st.columns(4)
    m[0].metric("Faturamento projetado", brl(proj["venda"]),
                delta_pct(proj["venda"], cur["venda"]) or "0,0%",
                help=f"Atual: {brl(cur['venda'])}")
    m[1].metric("Lucro projetado", brl(proj["lucro"]), lucro_delta,
                help=f"Atual: {brl(cur['lucro'])}")
    m[2].metric("Margem projetada", pct(proj["margem"]),
                delta_pp(proj["margem"], cur["margem"]),
                help=f"Atual: {pct(cur['margem'])}")
    m[3].metric("Variação absoluta do lucro", brl(proj["lucro"] - cur["lucro"]), lucro_delta)

    left, right = st.columns(2)
    left.plotly_chart(chart_waterfall(base, vol, price, disc, cost), width="stretch")
    right.plotly_chart(chart_compare(cur, proj), width="stretch")

    cmp = pd.DataFrame({
        "Indicador": ["Faturamento", "Custo", "Lucro", "Margem (%)"],
        "Cenário atual": [cur["venda"], cur["custo"], cur["lucro"], cur["margem"]],
        "Cenário projetado": [proj["venda"], proj["custo"], proj["lucro"], proj["margem"]],
    })
    cmp["Variação"] = cmp["Cenário projetado"] - cmp["Cenário atual"]
    st.dataframe(cmp.round(2), hide_index=True, width="stretch")
    st.caption("Premissas: o custo é variável com o volume; o desconto escala com preço e "
               "volume. Estimativa linear, sem efeitos de concorrência.")


# --------------------------------------------------------------------------
# Separador 4: Regional e exportação
# --------------------------------------------------------------------------
def chart_treemap(df: pd.DataFrame) -> go.Figure:
    g = df.groupby(["pais", "segmento"], as_index=False)[["venda", "lucro"]].sum()
    g = g[g["venda"] > 0]
    g["margem"] = g["lucro"] / g["venda"] * 100
    fig = px.treemap(g, path=[px.Constant("Global"), "pais", "segmento"], values="venda",
                     color="margem", color_continuous_scale="RdYlGn", color_continuous_midpoint=0,
                     custom_data=["lucro", "margem"],
                     title="Faturamento por país e segmento (cor = margem %)")
    fig.update_traces(
        hovertemplate=("<b>%{label}</b><br>Faturamento: R$ %{value:,.0f}"
                       "<br>Lucro: R$ %{customdata[0]:,.0f}"
                       "<br>Margem: %{customdata[1]:.1f}%<extra></extra>"),
        texttemplate="%{label}<br>R$ %{value:,.2s}")
    fig.update_layout(coloraxis_colorbar=dict(title="Margem %", ticksuffix="%"))
    return style_fig(fig, 460, hover=None)


def chart_map(df: pd.DataFrame) -> go.Figure:
    g = df.groupby("pais", as_index=False)[["venda", "lucro"]].sum()
    g["iso"] = g["pais"].map(ISO3)
    g["margem"] = g["lucro"] / g["venda"] * 100
    fig = px.choropleth(g, locations="iso", color="venda", hover_name="pais",
                        custom_data=["lucro", "margem"],
                        color_continuous_scale=[[0, "#D9E2EC"], [1, SLATE]],
                        title="Mapa de faturamento por país")
    fig.update_traces(hovertemplate=("<b>%{hovertext}</b><br>Faturamento: R$ %{z:,.0f}"
                                     "<br>Lucro: R$ %{customdata[0]:,.0f}"
                                     "<br>Margem: %{customdata[1]:.1f}%<extra></extra>"))
    fig.update_geos(fitbounds="locations", showframe=False, showcountries=True,
                    bgcolor="rgba(0,0,0,0)")
    fig.update_layout(coloraxis_colorbar=dict(title="R$", tickformat="~s"))
    return style_fig(fig, 460, hover=None)


def kpi_summary(df: pd.DataFrame, span: tuple) -> pd.DataFrame:
    k = compute_kpis(df)
    rows = [
        ("Período", f"{span[0]} a {span[1]}"), ("Transações", k["n"]),
        ("Faturamento líquido (R$)", round(k["venda"], 2)),
        ("Lucro operacional (R$)", round(k["lucro"], 2)),
        ("Margem operacional (%)", round(k["margem"], 2)),
        ("Unidades vendidas", round(k["unidades"], 2)),
        ("Ticket médio (R$)", round(k["ticket"], 2)),
        ("Desconto total (R$)", round(k["desconto"], 2)),
        ("Taxa média de desconto (%)", round(k["taxa_desc"], 2)),
    ]
    return pd.DataFrame(rows, columns=["Indicador", "Valor"]).astype({"Valor": "string"})


def to_excel_bytes(data: pd.DataFrame, summary: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        summary.to_excel(writer, index=False, sheet_name="Resumo KPIs")
        data.to_excel(writer, index=False, sheet_name="Dados filtrados")
    return buf.getvalue()


def explorer(df: pd.DataFrame) -> pd.DataFrame:
    """Filtros adicionais sobre a tabela de dados (governança/auditoria)."""
    c1, c2, c3 = st.columns([2, 2, 1])
    faixas = [f for f in DISCOUNT_ORDER if f in df["faixa_desconto"].unique()]
    sel = c1.multiselect("Faixa de desconto", faixas, default=faixas)
    lo, hi = float(df["venda"].min()), float(df["venda"].max())
    rng = (c2.slider("Faturamento da transação (R$)", lo, hi, (lo, hi))
           if hi > lo else (lo, hi))
    only_neg = c3.checkbox("Só lucro negativo")
    out = df[df["faixa_desconto"].isin(sel) & df["venda"].between(*rng)]
    return out[out["lucro"] < 0] if only_neg else out


def tab_regional(df: pd.DataFrame, span: tuple) -> None:
    left, right = st.columns(2)
    left.plotly_chart(chart_treemap(df), width="stretch")
    right.plotly_chart(chart_map(df), width="stretch")

    st.subheader("Explorador de dados")
    table = explorer(df)
    if table.empty:
        st.warning("Nenhum registo corresponde aos filtros da tabela.")
        return
    export = table.drop(columns=["periodo"]).copy()
    export["data"] = export["data"].dt.date
    summary = kpi_summary(table, span)

    c1, c2, c3, c4 = st.columns(4)
    page_size = c1.selectbox("Linhas por página", [10, 25, 50, 100], index=1)
    pages = max(1, math.ceil(len(export) / page_size))
    page = c2.number_input("Página", 1, pages, 1)
    c3.download_button("⬇️ Dados (CSV)", export.to_csv(index=False).encode("utf-8-sig"),
                       "financeiro_filtrado.csv", "text/csv", width="stretch")
    c4.download_button("⬇️ Resumo KPIs (CSV)", summary.to_csv(index=False).encode("utf-8-sig"),
                       "resumo_kpis.csv", "text/csv", width="stretch")
    start = (page - 1) * page_size
    st.dataframe(export.iloc[start:start + page_size], hide_index=True, width="stretch")
    st.caption(f"Página {page} de {pages} · {num(len(export))} registos")
    st.download_button(
        "⬇️ Excel (dados + resumo)", to_excel_bytes(export, summary),
        "financeiro_filtrado.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def main() -> None:
    st.set_page_config(page_title="Plataforma Analítica Executiva", page_icon="📈",
                       layout="wide")
    inject_css()
    st.title("📈 Plataforma Analítica Executiva")

    if not DATA_PATH.exists():
        st.error(f"Ficheiro `{DATA_PATH.name}` não encontrado junto da aplicação.")
        return

    df, prev, span = apply_filters(load_data())
    if df.empty:
        st.warning("Nenhum registo corresponde aos filtros selecionados. "
                   "Ajuste a barra lateral (ex.: botão **Todos**) para ver os dados.")
        return

    t1, t2, t3, t4 = st.tabs(["📊 Cockpit Executivo & KPIs",
                              "🎯 Matriz Estratégica de Portfólio",
                              "🔮 Simulador de Cenários",
                              "🌍 Performance Regional & Exportação"])
    with t1:
        tab_cockpit(df, prev)
    with t2:
        tab_portfolio(df)
    with t3:
        tab_simulator(df)
    with t4:
        tab_regional(df, span)


if __name__ == "__main__":
    main()
