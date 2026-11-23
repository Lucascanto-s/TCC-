"""Constantes visuais e layout único dos gráficos do EduIA Analytics."""

from __future__ import annotations

import plotly.graph_objects as go

COLORS = {
    "blue": "#1D4ED8",
    "orange": "#D97706",
    "green": "#059669",
    "red": "#DC2626",
    "purple": "#7C3AED",
    "charcoal": "#374151",
    "gray": "#9CA3AF",
    "background": "#FFFFFF",
    "card": "#F8FAFC",
    "border": "#E2E8F0",
    "title": "#0F172A",
    "text": "#1E293B",
    "muted": "#64748B",
    "sidebar": "#0F172A",
}

# Fontes do sistema: funcionam sem internet no dia da apresentação.
FONT_FAMILY = "Segoe UI, Inter, Roboto, Helvetica Neue, Arial, sans-serif"


def apply_layout(fig: go.Figure, height: int = 500) -> go.Figure:
    """Aplica o contrato visual único a todos os gráficos exibidos."""
    fig.update_layout(
        template="plotly_white",
        height=height,
        paper_bgcolor=COLORS["background"],
        plot_bgcolor=COLORS["background"],
        font=dict(size=16, color=COLORS["text"], family=FONT_FAMILY),
        title=dict(font=dict(size=20, color=COLORS["title"], family=FONT_FAMILY), x=0, xanchor="left"),
        margin=dict(l=10, r=30, t=80, b=30),
        hoverlabel=dict(font=dict(size=15, family=FONT_FAMILY), bgcolor="white"),
        legend=dict(font=dict(size=15, color=COLORS["text"])),
    )
    fig.update_xaxes(automargin=True, tickfont=dict(size=15, color=COLORS["text"]), showgrid=False, zeroline=False)
    fig.update_yaxes(automargin=True, tickfont=dict(size=16, color=COLORS["text"]), gridcolor="#EEF2F7", zeroline=False)
    return fig


PLOTLY_CONFIG = {
    "displayModeBar": "hover",
    "displaylogo": False,
    "modeBarButtonsToRemove": ["zoom2d", "pan2d", "select2d", "lasso2d", "zoomIn2d", "zoomOut2d", "autoScale2d", "resetScale2d"],
    "toImageButtonOptions": {"format": "png", "scale": 2},
    "responsive": True,
}
