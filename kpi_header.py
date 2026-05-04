import streamlit as st


def render_kpi_header(kpis: list[dict]):
    """Render a responsive row of Streamlit metric cards."""
    if not kpis:
        return
    columns = st.columns(len(kpis))
    for column, kpi in zip(columns, kpis):
        column.metric(
            label=kpi.get("label", ""),
            value=kpi.get("value", "-"),
            delta=kpi.get("delta"),
            delta_color=kpi.get("delta_color", "normal"),
            help=kpi.get("help"),
        )
