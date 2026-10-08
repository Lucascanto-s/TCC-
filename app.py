"""EduIA Analytics
Sistema de análise dos impactos da Inteligência Artificial no desempenho escolar.

Para executar:
    streamlit run app.py

O aplicativo aceita planilhas .xlsx com nomes de colunas em português ou inglês.
Quando nenhum arquivo é enviado, uma base demonstrativa é exibida para facilitar
a apresentação e a validação visual do TCC.
"""

from __future__ import annotations

import io
import re
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from streamlit_option_menu import option_menu

from theme import COLORS, FONT_SIZES, GRAPH_PALETTE, PLOTLY_CONFIG, apply_layout


# -----------------------------------------------------------------------------
# Configuração visual e constantes do projeto
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="EduIA Analytics",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap');
    :root { --ink: #111111; --title: #0f172a; --line: #cbd5e1; }
    html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; color: var(--ink); font-size: 18px; }
    h1 { font-family: 'Space Grotesk', sans-serif; font-size: 28px !important; color: var(--title); }
    h2, h3 { font-family: 'Space Grotesk', sans-serif; font-size: 22px !important; color: var(--title); }
    .stApp { background: #ffffff; }
    [data-testid="stSidebar"] { background: #0f172a; }
    section[data-testid="stSidebar"] { width: 320px !important; }
    [data-testid="stSidebar"] * { color: #ffffff; font-size: 18px; }
    [data-testid="stSidebar"] .nav-link { padding: 14px 12px !important; }
    [data-testid="stSidebar"] .nav-link-selected { background: #ffffff !important; color: #0f172a !important; font-weight: 700; }
    .brand { padding: 1.1rem 0 .7rem; }
    .brand-mark { display: inline-flex; width: 42px; height: 42px; align-items: center; justify-content: center;
        border-radius: 12px; background: linear-gradient(135deg, #38bdf8, #8b5cf6); font-size: 1.35rem; }
    .brand-title { font-family: 'Space Grotesk'; font-size: 1.28rem; font-weight: 700; margin-left: .55rem; vertical-align: middle; }
    .eyebrow { color: #1d4ed8; font-size: 16px; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; }
    .hero { border-radius: 8px; padding: 1.6rem 1.8rem; color: #111111; margin-bottom: 1.25rem;
        background: #f5f7fa; border: 1px solid #cbd5e1; position: relative; overflow: hidden; }
    .hero:after { content: ''; position: absolute; right: -40px; top: -75px; width: 220px; height: 220px; border: 30px solid rgba(255,255,255,.11); border-radius: 50%; }
    .hero h1, .hero h2 { margin: 0; color: #0f172a; font-size: 28px !important; }
    .hero p { margin: .45rem 0 0; color: #111111; max-width: 900px; font-size: 18px; }
    .kpi { border: 1px solid #cbd5e1; border-left: 6px solid #1d4ed8; border-radius: 8px; background: #f5f7fa; padding: 1rem 1.1rem; min-height: 126px; }
    .kpi-label { color: #111111; font-size: 16px; font-weight: 700; }
    .kpi-value { color: #0f172a; font: 700 2rem 'Space Grotesk'; margin-top: .3rem; line-height: 1.15; overflow-wrap: anywhere; }
    .section-note { color: #111111; margin-top: -.65rem; margin-bottom: 1rem; font-size: 18px; }
    .stCaption, [data-testid="stCaptionContainer"] { font-size: 16px !important; color: #111111 !important; }
    .stDownloadButton button { border-radius: 4px; font-size: 18px; }
    .question-banner { background: #f5f7fa; border: 1px solid #cbd5e1; border-left: 8px solid #1d4ed8; padding: 14px 18px; margin: 8px 0 14px; }
    .question-banner h2 { margin: 4px 0; line-height: 1.2; overflow-wrap: anywhere; }
    .question-banner p { margin: 0; font-size: 18px; color: #111111; }
    .insight-box, .academic-box { border: 1px solid #cbd5e1; border-left: 6px solid #059669; background: #f5f7fa; color: #111111; padding: 14px 16px; font-size: 18px; line-height: 1.35; margin: 8px 0 14px; }
    .academic-box { border-left-color: #1d4ed8; }
    .secondary-note { color: #111111; font-size: 18px; line-height: 1.3; margin: 8px 0; overflow-wrap: anywhere; }
    .chart-selector-label { color: #0f172a; font-size: 18px; font-weight: 700; margin: 4px 0 6px; }
    [data-testid="stRadio"] [role="radiogroup"] { gap: 8px; flex-wrap: wrap; }
    [data-testid="stRadio"] label { min-height: 48px; padding: 12px 16px !important; border: 1px solid #cbd5e1; border-radius: 8px; background: #f5f7fa; color: #111111 !important; font-size: 18px !important; font-weight: 600; }
    [data-testid="stRadio"] label:has(input:checked) { background: #1d4ed8 !important; color: #ffffff !important; border-color: #1d4ed8 !important; }
    [data-testid="stRadio"] label:has(input:checked) p { color: #ffffff !important; }
    [data-testid="stRadio"] label p { color: #111111 !important; font-size: 18px !important; }
    @media (max-width: 700px) { .hero { padding: 1.25rem; } .kpi-value { font-size: 2rem; } }
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# Dados: leitura, limpeza e base demonstrativa
# -----------------------------------------------------------------------------
def normalizar_nome(nome: str) -> str:
    """Converte o nome de uma coluna para um identificador comparável."""
    nome = unicodedata.normalize("NFKD", str(nome).strip().lower())
    nome = "".join(caractere for caractere in nome if not unicodedata.combining(caractere))
    nome = re.sub(r"[^a-z0-9 ]", "", nome)
    return re.sub(r"\s+", "_", nome)


def base_demo() -> pd.DataFrame:
    """Cria dados plausíveis para que o dashboard funcione antes do upload."""
    escolas = ["Colégio Estadual Jardim das Américas", "Colégio Estadual do Paraná", "Instituto de Educação do Paraná", "Colégio Estadual Santa Felicidade"]
    series = ["9º ano", "1ª série EM", "2ª série EM", "3ª série EM"]
    frequencias = ["Diariamente", "Semanalmente", "Raramente", "Nunca"]
    finalidades = ["Pesquisas e trabalhos", "Tirar dúvidas", "Resumos e revisão", "Produção de textos", "Exercícios"]
    linhas = []
    for indice in range(96):
        linhas.append({
            "Escola": escolas[indice % len(escolas)], "Série": series[indice % len(series)],
            "Frequência IA": frequencias[(indice * 3) % len(frequencias)],
            "Finalidade IA": finalidades[indice % len(finalidades)],
            "Redação Paraná": "Sim" if indice % 3 != 0 else "Não",
            "Fluência Paraná": "Sim" if indice % 4 != 0 else "Não",
            "IA substitui professor?": ["Não", "Parcialmente", "Sim"][(indice + 1) % 3],
            "Pensamento próprio": ["Aumentou", "Permaneceu igual", "Diminuiu"][(indice * 2) % 3],
        })
    return pd.DataFrame(linhas)


def padronizar_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Mapeia cabeçalhos curtos ou perguntas completas para o modelo interno."""
    # Cada lista contém termos que costumam aparecer nas perguntas do Forms.
    # O maior número de termos coincidentes vence, evitando depender do texto exato.
    palavras_chave = {
        "escola": [("escola",), ("colegio",), ("instituicao",), ("estabelecimento",)],
        "serie": [("serie",), ("ano", "escolar"), ("qual", "ano"), ("turma",)],
        "frequencia": [("frequencia",), ("frequente", "ia"), ("vezes", "ia"), ("utiliza", "ia")],
        "finalidade": [("finalidade",), ("utiliza", "ia", "para"), ("usa", "ia", "para"), ("atividade", "ia")],
        "redacao": [("redacao", "parana"), ("redacao",)],
        "fluencia": [("fluencia", "parana"), ("fluencia",)],
        "professor": [("substitui", "professor"), ("professor", "ia"), ("professor", "inteligencia", "artificial")],
        "pensamento": [("pensamento",), ("autonomia",), ("critico", "ia"), ("proprio", "ia")],
    }
    nomes = {normalizar_nome(coluna): coluna for coluna in dataframe.columns}
    renomear: dict[object, str] = {}
    usados: set[object] = set()
    ausentes: list[str] = []
    for destino, grupos in palavras_chave.items():
        candidatos = []
        for nome_normalizado, nome_original in nomes.items():
            if nome_original in usados:
                continue
            tokens = set(nome_normalizado.split("_"))
            pontuacao = max((sum(token in tokens for token in grupo) for grupo in grupos), default=0)
            # Para perguntas completas, também permite encontrar termos no texto.
            texto = nome_normalizado.replace("_", " ")
            pontuacao = max(pontuacao, max((sum(token in texto for token in grupo) for grupo in grupos), default=0))
            if pontuacao:
                candidatos.append((pontuacao, len(nome_normalizado), nome_original))
        if candidatos:
            _, _, escolhido = max(candidatos, key=lambda item: (item[0], -item[1]))
            renomear[escolhido] = destino
            usados.add(escolhido)
        else:
            ausentes.append(destino)
    resultado = dataframe.rename(columns=renomear).copy()

    # A pergunta 2 permite múltiplas escolhas e concentra as plataformas
    # estaduais numa única coluna. Criamos os indicadores usados nos gráficos.
    coluna_ferramentas = next(
        (coluna for coluna in resultado.columns if "ferrament" in normalizar_nome(coluna) or "plataforma" in normalizar_nome(coluna)),
        None,
    )
    if coluna_ferramentas is not None:
        ferramentas = resultado[coluna_ferramentas].fillna("").astype(str).map(normalizar_nome)
        resultado["ferramentas"] = resultado[coluna_ferramentas]
        resultado["redacao"] = ferramentas.map(lambda valor: "Sim" if "redacao_parana" in valor else "Não")
        resultado["fluencia"] = ferramentas.map(lambda valor: "Sim" if "fluencia_parana" in valor else "Não")
    else:
        resultado["ferramentas"] = "Não informado"

    # Mantém as perguntas numeradas sem alterar o texto original das colunas.
    # Isso permite apresentar os gráficos com a mesma redação do questionário.
    for numero in range(1, 9):
        coluna_pergunta = next((coluna for coluna in dataframe.columns if normalizar_nome(coluna).startswith(f"{numero}_")), None)
        resultado[f"pergunta_{numero}"] = dataframe[coluna_pergunta] if coluna_pergunta is not None else "Não informado"
    for coluna in palavras_chave:
        if coluna not in resultado.columns:
            resultado[coluna] = "Não informado"
    resultado.attrs["colunas_ausentes"] = ausentes
    campos_exibicao = list(palavras_chave) + ["ferramentas"] + [f"pergunta_{numero}" for numero in range(1, 9)]
    return resultado[campos_exibicao]


def carregar_dados() -> tuple[pd.DataFrame, bool]:
    """Carrega a base principal Total_master.xlsx que acompanha o aplicativo."""
    arquivo_master = Path(__file__).with_name("Total_master.xlsx")
    if not arquivo_master.exists():
        st.warning("A base Total_master.xlsx não foi encontrada. Exibindo dados demonstrativos.")
        return base_demo(), False
    try:
        # O arquivo possui uma aba por instituição e a aba Principal. Ler apenas
        # a primeira aba fazia o dashboard exibir somente uma escola.
        abas = pd.read_excel(arquivo_master, sheet_name=None, engine="openpyxl")
        tabelas = [tabela.dropna(how="all") for tabela in abas.values() if not tabela.dropna(how="all").empty]
        if not tabelas:
            raise ValueError("Nenhuma aba com respostas foi encontrada")
        planilha = pd.concat(tabelas, ignore_index=True)
        resultado = padronizar_dataframe(planilha)
        resultado.attrs["colunas_originais"] = [str(coluna) for coluna in planilha.columns]
        resultado.attrs["arquivo_fonte"] = arquivo_master.name
        resultado.attrs["abas_lidas"] = list(abas)
        return resultado, True
    except Exception as erro:  # Erro amigável na interface, sem derrubar o app.
        st.error(f"Não foi possível ler a planilha: {erro}")
        return base_demo(), False


def opcoes(dataframe: pd.DataFrame, coluna: str) -> list[str]:
    """Retorna valores válidos para filtros, removendo vazios."""
    return sorted(dataframe[coluna].dropna().astype(str).replace("", "Não informado").unique().tolist())


# -----------------------------------------------------------------------------
# Componentes reutilizáveis de apresentação
# -----------------------------------------------------------------------------
def titulo(pagina: str, descricao: str) -> None:
    st.markdown(f'<div class="eyebrow">EduIA Analytics / {pagina}</div>', unsafe_allow_html=True)
    st.title(pagina)
    st.markdown(f'<p class="section-note">{descricao}</p>', unsafe_allow_html=True)


def kpi(label: str, value: str, accent: str = "#2563EB") -> None:
    st.markdown(f'<div class="kpi" style="border-left-color:{accent}"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div></div>', unsafe_allow_html=True)


def grafico(fig, altura: int = 390, margem: tuple[int, int, int, int] | None = None) -> None:
    """Exibe gráficos somente após passarem pelo layout central do tema."""
    figura = apply_layout(fig, altura)
    if margem:
        esquerda, direita, topo, base = margem
        figura.update_layout(margin=dict(l=esquerda, r=direita, t=topo, b=base))
    st.plotly_chart(figura, width="stretch", config=PLOTLY_CONFIG)


def abreviar_rotulo(valor: object, limite: int = 32) -> str:
    """Quebra rótulos longos sem perder o texto completo no tooltip."""
    texto = str(valor)
    palavras = texto.split()
    linhas: list[str] = []
    linha = ""
    for palavra in palavras:
        if len(linha) + len(palavra) + 1 > limite and linha:
            linhas.append(linha)
            linha = palavra
        else:
            linha = f"{linha} {palavra}".strip()
    if linha:
        linhas.append(linha)
    return "<br>".join(linhas[:3])


def card_metricas(respostas: pd.DataFrame) -> None:
    """Exibe os indicadores de concentração da pergunta selecionada."""
    validas = respostas[respostas["resposta"] != "Não informado"]
    total = int(validas["quantidade"].sum())
    dominante = validas.loc[validas["quantidade"].idxmax()] if not validas.empty else {"resposta": "-", "percentual": 0, "quantidade": 0}
    menor = validas.loc[validas["quantidade"].idxmin()] if not validas.empty else {"resposta": "-", "percentual": 0, "quantidade": 0}
    def valor_curto(valor: object) -> str:
        return rotulo_curto(valor)

    metricas = st.columns(4)
    valores = [
        ("Total de respostas", f"{total}", COLORS["blue"]),
        ("Mais escolhida", valor_curto(dominante["resposta"]), COLORS["purple"]),
        ("Menos escolhida", valor_curto(menor["resposta"]), COLORS["orange"]),
        ("Percentual dominante", f"{dominante['percentual']:.1f}%", COLORS["green"]),
    ]
    for coluna, (rotulo, valor, cor) in zip(metricas, valores):
        with coluna:
            kpi(rotulo, valor, cor)
            if rotulo in {"Mais escolhida", "Menos escolhida"}:
                original = dominante["resposta"] if rotulo == "Mais escolhida" else menor["resposta"]
                st.markdown(f'<div title="{original}" style="font-size:16px;color:#111111;line-height:1.25">{original}</div>', unsafe_allow_html=True)


def gerar_insights(respostas: pd.DataFrame, pergunta: str) -> tuple[str, str]:
    """Gera uma leitura objetiva e uma interpretação acadêmica reutilizável."""
    validas = respostas[respostas["resposta"] != "Não informado"]
    if validas.empty:
        return "Não há respostas válidas para gerar insights.", "A base não contém observações suficientes para uma interpretação acadêmica."
    dominante = validas.loc[validas["quantidade"].idxmax()]
    menor = validas.loc[validas["quantidade"].idxmin()]
    concentracao = validas.nlargest(2, "quantidade")["percentual"].sum()
    insight = f"A resposta mais escolhida foi '{dominante['resposta']}' ({dominante['percentual']:.1f}%). A menor frequência foi '{menor['resposta']}' ({menor['percentual']:.1f}%)."
    if concentracao >= 70:
        insight += f" Existe forte concentração nas duas categorias principais ({concentracao:.1f}% das respostas)."
    else:
        insight += " As respostas apresentam distribuição relativamente diversificada entre as categorias."
    academica = f"Na pergunta '{pergunta[:110]}', os dados indicam predominância de '{dominante['resposta']}'."
    if dominante["percentual"] >= 50:
        academica += " A concentração majoritária sugere um padrão de comportamento relevante para a amostra analisada."
    else:
        academica += " Como nenhuma categoria supera metade da amostra, recomenda-se interpretar o resultado considerando a heterogeneidade dos participantes."
    return insight, academica


def grafico_pergunta_principal(respostas: pd.DataFrame) -> go.Figure:
    """Constrói automaticamente a melhor visualização legível para respostas categóricas."""
    cores = [COLORS["orange"] if indice == 0 else COLORS["blue"] for indice in range(len(respostas))]
    rotulos = respostas["resposta"].map(abreviar_rotulo)
    figura = go.Figure(go.Bar(x=respostas["percentual"], y=rotulos, orientation="h", marker=dict(color=cores, line=dict(color="white", width=1)), text=respostas["percentual"].map(lambda valor: f"{valor:.1f}%"), textfont=dict(size=18, color=COLORS["text"], family="DM Sans"), textposition="outside", customdata=respostas[["resposta", "quantidade"]], hovertemplate="<b>%{customdata[0]}</b><br>Respostas: %{customdata[1]}<br>Percentual: %{x:.1f}%<extra></extra>"))
    dominante = respostas.iloc[0]
    figura.update_layout(title=f"{dominante['percentual']:.1f}%: {dominante['resposta']}", showlegend=False)
    figura.update_yaxes(autorange="reversed")
    figura.update_xaxes(range=[0, 100], ticksuffix="%", title="Percentual de respondentes")
    return apply_layout(figura, max(440, min(760, 210 + len(respostas) * 58)))


def grafico_donut(respostas: pd.DataFrame) -> go.Figure:
    """Cria Donut sem legenda lateral e agrupa categorias menores que 5%."""
    principais = respostas[respostas["percentual"] >= 5].copy()
    pequenas = respostas[respostas["percentual"] < 5]
    if not pequenas.empty:
        principais = pd.concat([principais, pd.DataFrame([{"resposta": "Outras", "quantidade": pequenas["quantidade"].sum(), "percentual": pequenas["percentual"].sum()}])], ignore_index=True)
    dominante = principais.loc[principais["quantidade"].idxmax()]
    figura = go.Figure(go.Pie(labels=principais["resposta"].map(lambda valor: abreviar_rotulo(valor, 24)), values=principais["quantidade"], hole=.62, marker=dict(colors=GRAPH_PALETTE[:len(principais)], line=dict(color="white", width=2)), textposition="outside", textinfo="label+percent", textfont=dict(size=16, color=COLORS["text"]), customdata=principais["resposta"], hovertemplate="<b>%{customdata}</b><br>Respostas: %{value}<br>Percentual: %{percent}<extra></extra>"))
    figura.update_layout(title=f"{dominante['percentual']:.1f}%: {dominante['resposta']}", showlegend=False, annotations=[dict(text=f"{dominante['percentual']:.1f}%", x=.5, y=.5, font=dict(size=28, color=COLORS["title"]), showarrow=False)])
    return apply_layout(figura, 500)


def pergunta_ordenada(coluna: str, respostas: pd.DataFrame) -> bool:
    """Define se linha, área e dispersão possuem uma ordem interpretável."""
    valores = respostas["resposta"].map(normalizar_nome)
    prefixos = ("nao_utilizo", "raramente", "as_vezes", "frequentemente", "sempre")
    return coluna in {"pergunta_1", "pergunta_4"} or valores.map(lambda valor: any(valor.startswith(prefixo) for prefixo in prefixos)).all()


def opcoes_grafico(coluna: str, respostas: pd.DataFrame, multipla: bool) -> list[str]:
    """Retorna apenas os tipos coerentes com a pergunta atual."""
    tipos = ["📊 Barras horizontais", "▥ Barras verticais"]
    if not multipla:
        tipos.append("◉ Donut")
    if pergunta_ordenada(coluna, respostas):
        tipos.extend(["╱ Linha", "▰ Área", "✦ Dispersão com tendência"])
    return tipos


def grafico_barras_verticais(respostas: pd.DataFrame) -> go.Figure:
    """Barras verticais com rótulos quebrados e valores percentuais."""
    rotulos = respostas["resposta"].map(lambda valor: abreviar_rotulo(valor, 18))
    figura = go.Figure(go.Bar(x=rotulos, y=respostas["percentual"], marker=dict(color=[COLORS["orange"] if indice == 0 else COLORS["blue"] for indice in range(len(respostas))]), text=respostas["percentual"].map(lambda valor: f"{valor:.1f}%"), textfont=dict(size=18, color=COLORS["text"]), textposition="outside", customdata=respostas["resposta"], hovertemplate="<b>%{customdata}</b><br>Percentual: %{y:.1f}%<extra></extra>"))
    figura.update_layout(title=f"{respostas.iloc[0]['percentual']:.1f}%: {respostas.iloc[0]['resposta']}", showlegend=False)
    figura.update_xaxes(tickangle=0, title="Categoria")
    figura.update_yaxes(range=[0, 100], ticksuffix="%", title="Percentual de respondentes")
    return apply_layout(figura, max(440, min(700, 300 + len(respostas) * 36)))


def grafico_linha_area(respostas: pd.DataFrame, area: bool = False) -> go.Figure:
    """Linha ou área com pontos legíveis para escalas ordenadas."""
    rotulos = respostas["resposta"].map(lambda valor: abreviar_rotulo(valor, 18))
    trace = go.Scatter(x=rotulos, y=respostas["percentual"], mode="lines+markers+text", text=respostas["percentual"].map(lambda valor: f"{valor:.1f}%"), textposition="top center", textfont=dict(size=18, color=COLORS["text"]), line=dict(color=COLORS["blue"], width=4), marker=dict(size=10, color=COLORS["orange"]), fill="tozeroy" if area else None, fillcolor="rgba(29,78,216,.16)" if area else None, customdata=respostas["resposta"], hovertemplate="<b>%{customdata}</b><br>Percentual: %{y:.1f}%<extra></extra>")
    figura = go.Figure(trace)
    figura.update_layout(title="Tendência das respostas", showlegend=False)
    figura.update_yaxes(range=[0, 100], ticksuffix="%", title="Percentual de respondentes")
    return apply_layout(figura, 500)


def grafico_dispersao_tendencia(respostas: pd.DataFrame) -> go.Figure:
    """Scatter Plot com OLS, selo de direção, R² e equação."""
    dados = respostas.reset_index(drop=True).copy()
    dados["ordem"] = range(1, len(dados) + 1)
    figura = px.scatter(dados, x="ordem", y="percentual", trendline="ols", title="Relação entre ordem e percentual", labels={"ordem": "Ordem da escala", "percentual": "Percentual (%)"}, hover_data={"resposta": True, "ordem": False, "percentual": ":.1f"})
    resultados = px.get_trendline_results(figura)
    slope, pvalue, r2, equacao = 0.0, 1.0, 0.0, "y = 0.00x + 0.00"
    if not resultados.empty:
        modelo = resultados.iloc[0]["px_fit_results"]
        slope = float(modelo.params[1])
        pvalue = float(modelo.pvalues[1])
        r2 = float(modelo.rsquared)
        equacao = f"y = {slope:.2f}x + {float(modelo.params[0]):.2f}"
    status = "▲ Tendência crescente" if slope > .05 and pvalue <= .05 else "▼ Tendência decrescente" if slope < -.05 and pvalue <= .05 else "▬ Tendência estável"
    cor = COLORS["green"] if status.startswith("▲") else COLORS["red"] if status.startswith("▼") else COLORS["charcoal"]
    figura.update_traces(marker=dict(size=12, color=COLORS["blue"]), selector=dict(mode="markers"))
    figura.update_traces(line=dict(width=4, color=COLORS["orange"]), selector=dict(mode="lines"))
    figura.add_annotation(x=.02, y=1.14, xref="paper", yref="paper", text=f"{status} · R² = {r2:.2f} · {equacao} · p = {pvalue:.3f}", showarrow=False, font=dict(size=18, color=cor))
    figura.update_yaxes(range=[0, 100], ticksuffix="%")
    return apply_layout(figura, 500)


def grafico_principal_por_tipo(respostas: pd.DataFrame, tipo: str) -> go.Figure:
    """Seleciona a figura sem oferecer configuração visual avançada."""
    if tipo == "▥ Barras verticais":
        return grafico_barras_verticais(respostas)
    if tipo == "◉ Donut":
        return grafico_donut(respostas)
    if tipo == "╱ Linha":
        return grafico_linha_area(respostas)
    if tipo == "▰ Área":
        return grafico_linha_area(respostas, area=True)
    if tipo == "✦ Dispersão com tendência":
        return grafico_dispersao_tendencia(respostas)
    return grafico_pergunta_principal(respostas)


def percentual(dataframe: pd.DataFrame, coluna: str) -> pd.DataFrame:
    """Resume uma coluna com contagem e percentual, pronto para exibição."""
    resumo = dataframe[coluna].fillna("Não informado").astype(str).value_counts().rename_axis("resposta").reset_index(name="quantidade")
    resumo["percentual"] = (resumo["quantidade"] / max(len(dataframe), 1) * 100).round(1)
    return resumo


def ordenar_respostas(respostas: pd.DataFrame) -> pd.DataFrame:
    """Ordena escalas conhecidas naturalmente e demais respostas por frequência."""
    escala = [("nao_utilizo", 0), ("raramente", 1), ("as_vezes", 2), ("frequentemente", 3), ("sempre", 4)]
    normalizadas = respostas["resposta"].map(normalizar_nome)
    if normalizadas.map(lambda valor: any(valor.startswith(prefixo) for prefixo, _ in escala)).any():
        resultado = respostas.copy()
        resultado["_ordem"] = normalizadas.map(lambda valor: next((ordem for prefixo, ordem in escala if valor.startswith(prefixo)), 99))
        return resultado.sort_values(["_ordem", "quantidade"]).drop(columns="_ordem")
    return respostas.sort_values("quantidade")


ROTULOS_CURTOS = {
    "Inteligências Artificiais generativas (ChatGPT, Gemini, Copilot)": "IA generativa",
    "Plataformas estaduais/educacionais (Redação Paraná, Fluência Paraná, etc.)": "Plataformas do Estado",
    "Editores e geradores visuais/áudio (Canva, NotebookLM, geradores de áudio/slides)": "Editores e geradores",
    "Frequentemente (3 a 4 vezes por semana)": "Frequentemente",
    "Às vezes (1 a 2 vezes por semana)": "Às vezes",
    "Raramente (1 a 2 vezes por mês)": "Raramente",
    "Não Utilizo": "Não utilizo",
}


def rotulo_curto(valor: object) -> str:
    """Retorna um rótulo curto, legível e com no máximo três palavras."""
    texto = re.sub(r"\s+", " ", str(valor)).strip()
    for original, curto in ROTULOS_CURTOS.items():
        if texto.startswith(original) or original.startswith(texto):
            return curto
    palavras = re.findall(r"[\wÀ-ÿ]+", texto)
    return " ".join(palavras[:3]) if len(palavras) > 3 else texto


def opcoes_multiplas(valor: object) -> list[str]:
    """Extrai opções conhecidas mesmo quando há vírgulas dentro dos parênteses."""
    texto = re.sub(r"\s+", " ", str(valor)).strip()
    encontradas = [original for original in ROTULOS_CURTOS if original.lower() in texto.lower()]
    if encontradas:
        return encontradas
    separador = ";" if ";" in texto else ","
    return [parte.strip() for parte in texto.split(separador) if parte.strip()]


def eh_multipla(dataframe: pd.DataFrame, coluna: str) -> bool:
    """Detecta respostas com múltiplas opções na mesma célula."""
    valores = dataframe[coluna].dropna().astype(str)
    return coluna == "pergunta_2" or valores.map(lambda valor: len(opcoes_multiplas(valor)) > 1).mean() >= .5


def resumir_pergunta(dataframe: pd.DataFrame, coluna: str) -> tuple[pd.DataFrame, bool, list[dict[str, str]]]:
    """Resume escolha única ou menções de múltipla escolha com o denominador correto."""
    multipla = eh_multipla(dataframe, coluna)
    total_respondentes = max(len(dataframe), 1)
    nota_completa: list[dict[str, str]] = []
    if multipla:
        contagem: dict[str, int] = {}
        for resposta in dataframe[coluna].fillna("Não informado"):
            for opcao in opcoes_multiplas(resposta):
                chave = rotulo_curto(opcao)
                contagem[chave] = contagem.get(chave, 0) + 1
                if not any(item["curto"] == chave for item in nota_completa):
                    nota_completa.append({"curto": chave, "completo": str(opcao)})
        resumo = pd.DataFrame([{"resposta": chave, "quantidade": valor, "percentual": round(valor / total_respondentes * 100, 1)} for chave, valor in contagem.items()])
    else:
        bruto = dataframe[coluna].fillna("Não informado").astype(str).value_counts().rename_axis("completo").reset_index(name="quantidade")
        bruto["resposta"] = bruto["completo"].map(rotulo_curto)
        bruto["percentual"] = (bruto["quantidade"] / total_respondentes * 100).round(1)
        resumo = bruto[["resposta", "quantidade", "percentual"]]
        nota_completa = [{"curto": rotulo_curto(valor), "completo": valor} for valor in bruto["completo"].drop_duplicates()]
    resumo = resumo.sort_values("quantidade", ascending=False).reset_index(drop=True)
    if len(resumo) > 6:
        principais = resumo.head(5)
        outras = resumo.iloc[5:]["quantidade"].sum()
        resumo = pd.concat([principais, pd.DataFrame([{"resposta": "Outras", "quantidade": outras, "percentual": round(outras / total_respondentes * 100, 1)}])], ignore_index=True)
    return resumo, multipla, nota_completa


def percentual_uso_diario(respostas: pd.Series) -> float:
    """Calcula uso diário tanto para a base demo quanto para o questionário real."""
    valores = respostas.fillna("").astype(str).str.strip().str.lower()
    return valores.isin(["sempre", "diariamente"]).mean() * 100


# -----------------------------------------------------------------------------
# Páginas analíticas
# -----------------------------------------------------------------------------
def pagina_dashboard(dataframe: pd.DataFrame, origem_upload: bool) -> None:
    st.markdown('<div class="hero"><h1>EduIA Analytics</h1><p>Sistema de Análise dos Impactos da Inteligência Artificial no Desempenho Escolar</p><p><strong>Autores:</strong> Lucas Augusto Canto Silva · Vinicius Gabriel Paulicz Drewnowski</p></div>', unsafe_allow_html=True)
    if not origem_upload:
        st.info("Você está visualizando uma base demonstrativa. Importe um arquivo .xlsx na barra lateral para analisar as respostas reais.")
    st.subheader("Visão geral da pesquisa")
    escolas = dataframe["escola"].nunique()
    with st.container():
        colunas = st.columns(4)
        for coluna, (rotulo, valor, cor) in zip(colunas, [("Participantes", f"{len(dataframe):,}".replace(",", "."), COLORS["blue"]), ("Escolas", str(escolas), COLORS["purple"]), ("Séries", str(dataframe["serie"].nunique()), COLORS["sky"]), ("Uso diário de IA", f"{percentual_uso_diario(dataframe['frequencia']):.1f}%", COLORS["green"])]):
            with coluna:
                kpi(rotulo, valor, cor)
    esquerda, direita = st.columns(2)
    with esquerda:
        serie = percentual(dataframe, "serie")
        grafico(px.bar(serie, x="resposta", y="quantidade", color="resposta", color_discrete_sequence=[COLORS["blue"], COLORS["purple"], COLORS["sky"], COLORS["lilac"]], title="Distribuição por série").update_layout(showlegend=False))
    with direita:
        escola = percentual(dataframe, "escola")
        grafico(px.bar(escola, x="quantidade", y="resposta", orientation="h", color="quantidade", color_continuous_scale=["#bfdbfe", COLORS["purple"]], title="Participantes por instituição").update_layout(coloraxis_showscale=False))


def pagina_frequencia(dataframe: pd.DataFrame) -> None:
    titulo("Frequência de Uso da IA", "Observe a regularidade de utilização e filtre o recorte por instituição.")
    escolas = st.multiselect("Filtrar escolas", opcoes(dataframe, "escola"), default=[])
    filtrado = dataframe[dataframe["escola"].isin(escolas)] if escolas else dataframe
    resumo = percentual(filtrado, "frequencia")
    col1, col2 = st.columns([1.35, 1])
    with col1:
        grafico(px.bar(resumo, x="resposta", y="quantidade", text=resumo["percentual"].map(lambda x: f"{x:.1f}%"), color="resposta", color_discrete_sequence=[COLORS["blue"], COLORS["sky"], COLORS["purple"], "#CBD5E1"], title="Frequência declarada").update_traces(textposition="outside").update_layout(showlegend=False))
    with col2:
        grafico(px.pie(resumo, names="resposta", values="quantidade", hole=.55, color_discrete_sequence=[COLORS["blue"], COLORS["sky"], COLORS["purple"], "#CBD5E1"], title="Composição percentual"))
    st.dataframe(resumo.rename(columns={"resposta": "Frequência", "quantidade": "Respostas", "percentual": "%"}), width="stretch", hide_index=True)


def pagina_finalidade(dataframe: pd.DataFrame) -> None:
    titulo("Finalidade de Utilização da IA", "Ranking das atividades em que os estudantes recorrem às ferramentas de inteligência artificial.")
    resumo = percentual(dataframe, "finalidade").sort_values("quantidade")
    grafico(px.bar(resumo, x="quantidade", y="resposta", orientation="h", text=resumo["percentual"].map(lambda x: f"{x:.1f}%"), color="quantidade", color_continuous_scale=["#c4b5fd", COLORS["purple"]], title="Ranking de finalidades").update_traces(textposition="outside").update_layout(coloraxis_showscale=False))
    st.dataframe(resumo.sort_values("quantidade", ascending=False).rename(columns={"resposta": "Finalidade", "quantidade": "Respostas", "percentual": "%"}), width="stretch", hide_index=True)


def pagina_plataformas(dataframe: pd.DataFrame) -> None:
    titulo("Plataformas Educacionais do Estado", "Uso declarado das plataformas Redação Paraná e Fluência Paraná.")
    registros = []
    for coluna, nome in [("redacao", "Redação Paraná"), ("fluencia", "Fluência Paraná")]:
        for resposta, quantidade, percent in percentual(dataframe, coluna).itertuples(index=False):
            registros.append({"Plataforma": nome, "Resposta": resposta, "Quantidade": quantidade, "Percentual": percent})
    resumo = pd.DataFrame(registros)
    col1, col2 = st.columns(2)
    with col1:
        grafico(px.bar(resumo, x="Plataforma", y="Percentual", color="Resposta", barmode="group", text="Percentual", color_discrete_sequence=[COLORS["green"], COLORS["orange"], COLORS["purple"]], title="Percentual de utilização").update_traces(texttemplate="%{text:.1f}%", textposition="outside").update_layout(yaxis_title="Percentual (%)"))
    with col2:
        respostas_positivas = resumo[resumo["Resposta"].str.lower().isin(["sim", "sempre", "frequentemente"])]
        totais = respostas_positivas.groupby("Plataforma", as_index=False)["Percentual"].sum()
        grafico(px.bar(totais, x="Plataforma", y="Percentual", text="Percentual", color="Plataforma", color_discrete_sequence=[COLORS["blue"], COLORS["purple"]], title="Adesão estimada").update_traces(texttemplate="%{text:.1f}%", textposition="outside").update_layout(showlegend=False, yaxis_title="Percentual (%)"))


def pagina_professor(dataframe: pd.DataFrame) -> None:
    titulo("IA x Professor", "Compare a percepção dos estudantes sobre a relação entre inteligência artificial e mediação docente.")
    resumo = percentual(dataframe, "professor")
    col1, col2 = st.columns(2)
    with col1:
        grafico(px.pie(resumo, names="resposta", values="quantidade", hole=.48, color_discrete_sequence=[COLORS["blue"], COLORS["purple"], COLORS["orange"]], title="A IA substitui o professor?"))
    with col2:
        grafico(px.bar(resumo, x="resposta", y="quantidade", color="resposta", text="percentual", color_discrete_sequence=[COLORS["blue"], COLORS["purple"], COLORS["orange"]], title="Comparação das respostas").update_traces(texttemplate="%{text:.1f}%", textposition="outside").update_layout(showlegend=False))


def pagina_pensamento(dataframe: pd.DataFrame) -> None:
    titulo("Impacto no Pensamento Próprio", "Indicadores da percepção dos estudantes sobre autonomia e pensamento crítico.")
    resumo = percentual(dataframe, "pensamento")
    colunas = st.columns(3)
    for coluna, resposta, cor in zip(colunas, resumo["resposta"].tolist()[:3], [COLORS["green"], COLORS["blue"], COLORS["orange"]]):
        with coluna:
            valor = resumo.loc[resumo["resposta"] == resposta, "percentual"].iloc[0]
            kpi(resposta, f"{valor:.1f}%", cor)
    grafico(px.bar(resumo, x="resposta", y="quantidade", text=resumo["percentual"].map(lambda x: f"{x:.1f}%"), color="resposta", color_discrete_sequence=[COLORS["green"], COLORS["blue"], COLORS["orange"]], title="Distribuição das respostas").update_traces(textposition="outside").update_layout(showlegend=False))
    st.caption(f"Média de respostas válidas: {resumo['quantidade'].sum()} · Categorias observadas: {len(resumo)}")


def pagina_perguntas(dataframe: pd.DataFrame) -> None:
    """Exibe análise inteligente da pergunta selecionada e suas respostas originais."""
    st.markdown('<div class="eyebrow">EduIA Analytics / análise principal</div>', unsafe_allow_html=True)
    st.title("Perguntas e Respostas")
    perguntas = {
        "pergunta_1": "1. Você utiliza ferramentas de Inteligência Artificial para estudar ou realizar tarefas escolares?",
        "pergunta_2": "2. Quais ferramentas de IA você mais utiliza na sua rotina de estudos? (Marque todas que se aplicam)",
        "pergunta_3": "3. Com qual finalidade principal você costuma usar a IA nos estudos?",
        "pergunta_4": "4. Com qual frequência você utiliza as plataformas com IA do estado (ex: Redação Paraná, Fluência Paraná)?",
        "pergunta_5": "5. Como os seus professores costumam reagir ao saber que os alunos usam IA nos trabalhos?",
        "pergunta_6": "6. Algum professor já passou uma atividade ou trabalho que EXIGIA o uso de ferramentas de IA?",
        "pergunta_7": "7. Você sente que o uso da IA diminui sua capacidade de pensar por conta própria?",
        "pergunta_8": "8. Na sua opinião, uma ferramenta de IA pode explicar a matéria melhor que um professor?",
    }
    controles = st.columns([2.2, 1])
    with controles[0]:
        escolhida = st.selectbox("Selecione uma pergunta", list(perguntas), format_func=lambda chave: perguntas[chave])
    with controles[1]:
        filtro_escola = st.selectbox("Instituição", ["Todas"] + opcoes(dataframe, "escola"))
    filtrado = dataframe if filtro_escola == "Todas" else dataframe[dataframe["escola"] == filtro_escola]
    respostas, multipla, notas = resumir_pergunta(filtrado, escolhida)
    st.markdown(f"<div class='question-banner'><div class='eyebrow'>Enunciado em análise</div><h2>{perguntas[escolhida]}</h2><p>Recorte: {filtro_escola} · {len(filtrado)} respondentes · {'múltipla escolha' if multipla else 'escolha única'}</p></div>", unsafe_allow_html=True)
    if respostas.empty or respostas["resposta"].eq("Não informado").all():
        st.warning("Não há respostas válidas para esta pergunta na base atual.")
        return
    card_metricas(respostas)
    insight, academica = gerar_insights(respostas, perguntas[escolhida])
    aba_visualizacoes, aba_insights, aba_tabela = st.tabs(["Visualizações", "Insights e interpretação", "Dados detalhados"])
    with aba_visualizacoes:
        tipos = opcoes_grafico(escolhida, respostas, multipla)
        escolhas_grafico = st.session_state.setdefault("tipos_grafico_por_pergunta", {})
        if escolhas_grafico.get(escolhida) not in tipos:
            escolhas_grafico[escolhida] = tipos[0]
        st.markdown("<div class='chart-selector-label'>Tipo de gráfico</div>", unsafe_allow_html=True)
        tipo = st.radio("Tipo de gráfico", tipos, index=tipos.index(escolhas_grafico[escolhida]), horizontal=True, label_visibility="collapsed", help="Tipos indisponíveis são ocultados quando não representam corretamente a pergunta.")
        escolhas_grafico[escolhida] = tipo
        esquerda, direita = st.columns([1.6, 1])
        with esquerda:
            grafico(grafico_principal_por_tipo(respostas, tipo))
        with direita:
            if multipla or tipo in {"▥ Barras verticais", "╱ Linha", "▰ Área", "✦ Dispersão com tendência"}:
                st.markdown("<h3>Insight automático</h3>", unsafe_allow_html=True)
                st.markdown(f"<div class='insight-box'>{insight}</div>", unsafe_allow_html=True)
            else:
                grafico(grafico_donut(respostas), margem=(260, 120, 100, 120))
                st.markdown(f"<div class='insight-box'>{insight}</div>", unsafe_allow_html=True)
            if notas:
                nota = "<br>".join(f"<b>{item['curto']}</b>: {item['completo']}" for item in notas)
                st.markdown(f"<div class='secondary-note'><b>Texto completo das respostas:</b><br>{nota}</div>", unsafe_allow_html=True)
    with aba_insights:
        st.markdown(f"<div class='insight-box'><b>Insight automático</b><br>{insight}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='academic-box'><b>Interpretação acadêmica</b><br>{academica}</div>", unsafe_allow_html=True)
        st.subheader("Ranking das respostas")
        ranking = respostas.sort_values("quantidade", ascending=False).reset_index(drop=True)
        ranking.insert(0, "Posição", ranking.index + 1)
        st.dataframe(ranking.rename(columns={"resposta": "Resposta", "quantidade": "Quantidade", "percentual": "Percentual (%)"}), width="stretch", hide_index=True)
    with aba_tabela:
        st.dataframe(respostas.rename(columns={"resposta": "Resposta original", "quantidade": "Quantidade", "percentual": "Percentual (%)"}), width="stretch", hide_index=True)
        st.markdown(f"<div class='secondary-note'>Categorias observadas: {len(respostas)} · Menções: {int(respostas['quantidade'].sum())} · Percentual total: {respostas['percentual'].sum():.1f}%</div>", unsafe_allow_html=True)


def indice_frequencia(respostas: pd.Series) -> pd.Series:
    """Converte a escala ordinal de uso da IA em índice analítico."""
    valores = respostas.fillna("").astype(str).map(normalizar_nome)
    return valores.map(lambda valor: 4 if "sempre" in valor else 3 if "frequentemente" in valor or "sim_com_frequencia" in valor else 2 if "vezes" in valor else 1 if "raramente" in valor else 0)


def tendencia_status(valores: pd.Series) -> tuple[str, str]:
    """Classifica uma série por sua variação do primeiro ao último ponto."""
    if len(valores) < 2:
        return "Estável", COLORS["muted"]
    variacao = float(valores.iloc[-1] - valores.iloc[0])
    if variacao > 0.15:
        return "Crescente", COLORS["green"]
    if variacao < -0.15:
        return "Decrescente", COLORS["orange"]
    return "Estável", COLORS["blue"]


def figura_tendencia(tabela: pd.DataFrame, titulo_grafico: str) -> go.Figure:
    """Cria linha suavizada com média móvel e linha de tendência."""
    tabela = tabela.reset_index(drop=True)
    x = list(range(len(tabela)))
    y = tabela["indice"].astype(float).tolist()
    figura = go.Figure()
    figura.add_trace(go.Scatter(x=tabela["grupo"], y=y, mode="lines+markers", name="Média observada", line=dict(color=COLORS["purple"], width=3, shape="spline"), marker=dict(size=9), hovertemplate="<b>%{x}</b><br>Índice: %{y:.2f}<extra></extra>"))
    if len(y) > 1:
        media_x = sum(x) / len(x)
        media_y = sum(y) / len(y)
        denominador = sum((item - media_x) ** 2 for item in x) or 1
        inclinacao = sum((item - media_x) * (valor - media_y) for item, valor in zip(x, y)) / denominador
        intercepto = media_y - inclinacao * media_x
        ajuste = [intercepto + inclinacao * item for item in x]
        figura.add_trace(go.Scatter(x=tabela["grupo"], y=ajuste, mode="lines", name="Regressão linear", line=dict(color=COLORS["orange"], dash="dash", width=2), hovertemplate="Tendência: %{y:.2f}<extra></extra>"))
    if len(y) >= 3:
        media_movel = pd.Series(y).rolling(3, min_periods=1).mean()
        figura.add_trace(go.Scatter(x=tabela["grupo"], y=media_movel, mode="lines", name="Média móvel", line=dict(color=COLORS["sky"], width=2), hovertemplate="Média móvel: %{y:.2f}<extra></extra>"))
    figura.update_layout(title=titulo_grafico, yaxis_title="Índice médio", xaxis_title="Grupo")
    return apply_layout(figura, 430)


def pagina_tendencias(dataframe: pd.DataFrame) -> None:
    """Analisa comportamentos por série e instituição com linhas de tendência."""
    titulo("Tendências e Comportamentos", "Leitura longitudinal dos padrões de uso, aceitação e relação com o professor.")
    base = dataframe.copy()
    base["indice_uso"] = indice_frequencia(base["frequencia"])
    base["indice_pensamento"] = base["pensamento"].astype(str).str.lower().map(lambda valor: 3 if "não" in valor and "diminui" in valor else 2 if "um pouco" in valor else 1 if "sim" in valor or "diminui" in valor else 2)
    base["indice_professor"] = base["professor"].astype(str).str.lower().map(lambda valor: 3 if "sim" in valor or "melhor" in valor else 2 if "parcial" in valor or "aceitam" in valor else 1)
    visao = st.radio("Agrupar tendências por", ["Série", "Instituição"], horizontal=True)
    campo = "serie" if visao == "Série" else "escola"
    agrupado = base.groupby(campo, as_index=False).agg(indice_uso=("indice_uso", "mean"), indice_pensamento=("indice_pensamento", "mean"), indice_professor=("indice_professor", "mean"), participantes=("escola", "size")).sort_values(campo)
    status = []
    para_graficos = [("Adoção da IA", "indice_uso"), ("Dependência / autonomia", "indice_pensamento"), ("Aceitação e IA x professor", "indice_professor")]
    colunas = st.columns(3)
    for coluna, (nome, campo_indice) in zip(colunas, para_graficos):
        direcao, cor = tendencia_status(agrupado[campo_indice])
        with coluna:
            kpi(nome, direcao, cor)
        status.append({"Indicador": nome, "Direção": direcao})
    escolha = st.selectbox("Comportamento em detalhe", [item[0] for item in para_graficos])
    campo_escolhido = dict(para_graficos)[escolha]
    tabela = agrupado.rename(columns={campo: "grupo", campo_escolhido: "indice"})[["grupo", "indice", "participantes"]]
    grafico(figura_tendencia(tabela, escolha))
    st.dataframe(pd.DataFrame(status), use_container_width=True, hide_index=True)


def pagina_escola(dataframe: pd.DataFrame) -> None:
    titulo("Análise por Escola", "Compare indicadores entre instituições e aprofunde a leitura de uma escola específica.")
    selecionadas = st.multiselect("Instituições para comparar", opcoes(dataframe, "escola"), default=opcoes(dataframe, "escola")[:3])
    filtrado = dataframe[dataframe["escola"].isin(selecionadas)] if selecionadas else dataframe
    tabela = filtrado.groupby("escola", as_index=False).agg(Participantes=("escola", "size"), Uso_diario=("frequencia", lambda valores: round(percentual_uso_diario(valores), 1)), Uso_redacao=("redacao", lambda valores: round(valores.astype(str).str.lower().eq("sim").mean() * 100, 1)), Autonomia_aumentou=("pensamento", lambda valores: round(valores.astype(str).str.lower().eq("aumentou").mean() * 100, 1))).sort_values("Participantes", ascending=False)
    st.dataframe(tabela.rename(columns={"escola": "Instituição", "Uso_diario": "Uso diário (%)", "Uso_redacao": "Redação Paraná (%)", "Autonomia_aumentou": "Pensamento próprio aumentou (%)"}), use_container_width=True, hide_index=True)
    if not tabela.empty:
        grafico(px.bar(tabela, x="escola", y=["Uso_diario", "Uso_redacao", "Autonomia_aumentou"], barmode="group", title="Indicadores comparativos", labels={"value": "Percentual (%)", "variable": "Indicador", "escola": "Instituição"}, color_discrete_sequence=[COLORS["blue"], COLORS["purple"], COLORS["green"]]))


def relatorio_texto(dataframe: pd.DataFrame) -> str:
    """Monta um relatório textual simples e portátil para download."""
    uso_diario = percentual_uso_diario(dataframe["frequencia"])
    return f"""EduIA Analytics
Relatório consolidado | Gerado em {datetime.now():%d/%m/%Y %H:%M}

Pesquisa: Impactos do Uso de Inteligência Artificial no Desempenho Escolar
Estudo de Caso nas Escolas com Maiores Índices no IDEB em Curitiba - PR

Autores: Lucas Augusto Canto Silva e Vinicius Gabriel Paulicz Drewnowski

INDICADORES GERAIS
Participantes: {len(dataframe)}
Instituições: {dataframe['escola'].nunique()}
Uso diário de IA: {uso_diario:.1f}%

FREQUÊNCIA DE USO
{percentual(dataframe, 'frequencia').to_string(index=False)}

PENSAMENTO PRÓPRIO
{percentual(dataframe, 'pensamento').to_string(index=False)}
"""


def pdf_graficos(dataframe: pd.DataFrame) -> bytes | None:
    """Gera um PDF com gráficos principais usando o mecanismo do Plotly.

    O exportador do Plotly depende do pacote opcional ``kaleido``. O retorno
    nulo permite que a aplicação continue utilizável quando esse pacote ainda
    não estiver instalado no ambiente de execução.
    """
    try:
        import plotly.graph_objects as go

        frequencia = percentual(dataframe, "frequencia")
        pensamento = percentual(dataframe, "pensamento")
        figura = go.Figure()
        figura.add_bar(name="Frequência de uso", x=frequencia["resposta"], y=frequencia["quantidade"], marker_color=COLORS["blue"])
        figura.add_bar(name="Pensamento próprio", x=pensamento["resposta"], y=pensamento["quantidade"], marker_color=COLORS["purple"])
        figura.update_layout(title="EduIA Analytics - Indicadores principais", barmode="group", width=1100, height=700, font=dict(family="DM Sans"))
        return figura.to_image(format="pdf", engine="kaleido")
    except Exception:
        return None


def pagina_exportacao(dataframe: pd.DataFrame) -> None:
    titulo("Exportação", "Baixe um relatório consolidado e uma página HTML com os dados resumidos para documentação do TCC.")
    st.subheader("Relatório consolidado")
    st.download_button("⬇ Baixar relatório TXT", relatorio_texto(dataframe), file_name="relatorio_eduia_analytics.txt", mime="text/plain")
    pdf = pdf_graficos(dataframe)
    if pdf is not None:
        st.download_button("⬇ Gerar PDF dos gráficos", pdf, file_name="eduia_graficos.pdf", mime="application/pdf")
    else:
        st.warning("Para habilitar o PDF dos gráficos, instale a dependência opcional com: `pip install kaleido`.")
    st.divider()
    st.subheader("Dados tratados")
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as escritor:
        dataframe.to_excel(escritor, index=False, sheet_name="Respostas tratadas")
        percentual(dataframe, "frequencia").to_excel(escritor, index=False, sheet_name="Frequência")
    st.download_button("⬇ Baixar Excel tratado", buffer.getvalue(), file_name="eduia_dados_tratados.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("Os gráficos Plotly também podem ser salvos em PNG/SVG pelo menu de cada gráfico. Para PDF, use a opção de impressão do navegador sobre a página ou exporte os gráficos com o botão de câmera do Plotly.")


# -----------------------------------------------------------------------------
# Entrada principal e navegação
# -----------------------------------------------------------------------------
def main() -> None:
    with st.sidebar:
        st.markdown('<div class="brand"><span class="brand-mark">✦</span><span class="brand-title">EduIA Analytics</span></div>', unsafe_allow_html=True)
        st.markdown("Pesquisa escolar · Curitiba, PR", unsafe_allow_html=True)
        st.markdown("Fonte: Total_master.xlsx", unsafe_allow_html=True)
        modo_apresentacao = st.toggle("Modo apresentação", value=False) if hasattr(st, "toggle") else False
        st.divider()
        paginas = ["Dashboard Geral", "Perguntas e Respostas", "Tendências e Comportamentos", "Frequência de Uso da IA", "Finalidade de Utilização da IA", "Plataformas Educacionais do Estado", "IA x Professor", "Impacto no Pensamento Próprio", "Análise por Escola", "Exportação"]
        icones = ["speedometer2", "chat-square-text", "graph-up-arrow", "activity", "list-check", "book", "person-video3", "lightbulb", "building", "download"]
        st.markdown("""
        <style>
        [data-testid="stSidebar"] { background: #0F172A !important; min-width: 320px !important; }
        [data-testid="stSidebar"] .nav-link,
        [data-testid="stSidebar"] .nav-link span,
        [data-testid="stSidebar"] .nav-item a { color: #FFFFFF !important; font-size: 18px !important; font-weight: 600 !important; white-space: normal !important; }
        [data-testid="stSidebar"] .nav-link i { color: #FFFFFF !important; font-size: 24px !important; }
        [data-testid="stSidebar"] .nav-link.active,
        [data-testid="stSidebar"] .nav-link.active span { background: #FFFFFF !important; color: #0F172A !important; font-weight: 700 !important; }
        [data-testid="stSidebar"] .nav-link.active i { color: #0F172A !important; }
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] span { color: #FFFFFF !important; font-size: 18px !important; }
        [data-testid="stSidebar"] .nav-link { padding: 14px 12px !important; margin: 6px 0 !important; border-radius: 8px !important; }
        </style>
        """, unsafe_allow_html=True)
        pagina = option_menu(
            menu_title=None,
            options=paginas,
            icons=icones,
            menu_icon="cast",
            default_index=1,
            styles={
                "container": {"padding": "8px!important", "background-color": "#0F172A"},
                "icon": {"color": "#FFFFFF", "font-size": "24px"},
                "nav-link": {"color": "#FFFFFF", "font-size": "18px", "font-weight": "600", "text-align": "left", "margin": "6px 0", "padding": "14px 12px", "border-radius": "8px", "--hover-color": "#1E3A8A"},
                "nav-link-selected": {"background-color": "#FFFFFF", "color": "#0F172A", "font-weight": "700"},
            },
        )
        st.divider()
        if not modo_apresentacao:
            st.markdown("TCC · Engenharia de Software", unsafe_allow_html=True)
    if modo_apresentacao:
        st.markdown("""
        <style>
        .section-note, .eyebrow, .secondary-note, [data-testid="stCaptionContainer"] { display: none !important; }
        </style>
        """, unsafe_allow_html=True)
    dataframe, origem_upload = carregar_dados()
    if pagina == "Dashboard Geral":
        pagina_dashboard(dataframe, origem_upload)
    elif pagina == "Perguntas e Respostas":
        pagina_perguntas(dataframe)
    elif pagina == "Tendências e Comportamentos":
        pagina_tendencias(dataframe)
    elif pagina == "Frequência de Uso da IA":
        pagina_frequencia(dataframe)
    elif pagina == "Finalidade de Utilização da IA":
        pagina_finalidade(dataframe)
    elif pagina == "Plataformas Educacionais do Estado":
        pagina_plataformas(dataframe)
    elif pagina == "IA x Professor":
        pagina_professor(dataframe)
    elif pagina == "Impacto no Pensamento Próprio":
        pagina_pensamento(dataframe)
    elif pagina == "Análise por Escola":
        pagina_escola(dataframe)
    else:
        pagina_exportacao(dataframe)


if __name__ == "__main__":
    main()