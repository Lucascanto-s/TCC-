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
    "card": "#F5F7FA",
    "border": "#CBD5E1",
    "title": "#0F172A",
    "text": "#111111",
    "sidebar": "#0F172A",
    # Aliases legados usados pelas telas que serão revisadas depois.
    "ink": "#111111",
    "muted": "#374151",
    "sky": "#1D4ED8",
    "lilac": "#9CA3AF",
}

FONT_SIZES = {
    "title": 28,
    "subtitle": 22,
    "body": 18,
    "axis": 16,
    "value": 18,
    "kpi": 36,
    "hover": 16,
}

GRAPH_PALETTE = [
    COLORS["blue"], COLORS["orange"], COLORS["green"],
    COLORS["red"], COLORS["purple"], COLORS["charcoal"],
]


def apply_layout(fig: go.Figure, height: int = 500) -> go.Figure:
    """Aplica o contrato visual único a todos os gráficos exibidos."""
    fig.update_layout(
        template="plotly_white",
        height=height,
        paper_bgcolor=COLORS["background"],
        plot_bgcolor=COLORS["background"],
        font=dict(size=FONT_SIZES["body"], color=COLORS["text"], family="DM Sans"),
        title=dict(font=dict(size=24, color=COLORS["title"], family="Space Grotesk")),
        margin=dict(l=260, r=60, t=80, b=60),
        hoverlabel=dict(font=dict(size=FONT_SIZES["hover"], color=COLORS["text"]), bgcolor=COLORS["background"]),
        legend=dict(font=dict(size=FONT_SIZES["axis"], color=COLORS["text"])),
    )
    fig.update_xaxes(automargin=True, tickfont=dict(size=FONT_SIZES["axis"], color=COLORS["text"]), title_font=dict(size=FONT_SIZES["axis"], color=COLORS["title"]), showgrid=False)
    fig.update_yaxes(automargin=True, tickfont=dict(size=FONT_SIZES["axis"], color=COLORS["text"]), title_font=dict(size=FONT_SIZES["axis"], color=COLORS["title"]), gridcolor="#E5E7EB")
    return fig


PLOTLY_CONFIG = {
    "displayModeBar": False,
    "responsive": True,
}
