"""EduIA Analytics
Sistema de análise dos impactos da Inteligência Artificial no desempenho escolar.

Para executar:
    streamlit run app.py

Os dados vêm de Total_master.xlsx (respostas do Google Forms). As perguntas são
identificadas pelo número ("1.", "2.", ...). Sem a planilha na pasta, o app
avisa e não exibe resultados.
"""

from __future__ import annotations

import io
import re
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from streamlit_option_menu import option_menu

from theme import COLORS, FONT_FAMILY, PLOTLY_CONFIG, apply_layout


# -----------------------------------------------------------------------------
# Informações do trabalho (edite aqui; campos vazios não aparecem na tela)
# -----------------------------------------------------------------------------
TITULO_TCC = "Impactos do Uso de Inteligência Artificial no Desempenho Escolar"
SUBTITULO_TCC = "Estudo de caso nas escolas com maiores índices no IDEB em Curitiba – PR"
AUTORES = "Lucas Augusto Canto Silva · Vinicius Gabriel Paulicz Drewnowski"
ORIENTADOR = ""
CURSO = "Engenharia de Software"
INSTITUICAO = ""
PERIODO_COLETA = ""

# Interpretação escrita pelos autores para cada pergunta. Quando preenchida,
# aparece na aba Perguntas e Respostas ao lado da leitura automática dos dados.
INTERPRETACOES: dict[str, str] = {
    "pergunta_1": "",
    "pergunta_2": "",
    "pergunta_3": "",
    "pergunta_4": "",
    "pergunta_5": "",
    "pergunta_6": "",
    "pergunta_7": "",
    "pergunta_8": "",
}


# -----------------------------------------------------------------------------
# Configuração visual
# -----------------------------------------------------------------------------
st.set_page_config(page_title="EduIA Analytics", page_icon="📘", layout="wide", initial_sidebar_state="expanded")

st.markdown(
    f"""
    <style>
    html, body, [class*="css"], .stMarkdown, button, input, select {{ font-family: {FONT_FAMILY}; }}
    .block-container {{ padding-top: 3.6rem; padding-bottom: 2rem; max-width: 1280px; }}
    h1 {{ font-size: 2rem !important; color: #0F172A; letter-spacing: -.01em; }}
    h2, h3 {{ color: #0F172A; }}
    [data-testid="stSidebar"] {{ background: #0F172A; min-width: 250px !important; max-width: 250px !important; }}
    [data-testid="stSidebar"] p, [data-testid="stSidebar"] label, [data-testid="stSidebar"] span {{ color: #E2E8F0 !important; }}
    .brand {{ display: flex; align-items: center; gap: .6rem; padding: .2rem 0 .3rem; }}
    .brand-mark {{ width: 38px; height: 38px; border-radius: 10px; display: inline-flex; align-items: center; justify-content: center;
                  background: linear-gradient(135deg, #3B82F6, #1D4ED8); color: #fff; font-weight: 800; font-size: 1.05rem; }}
    .brand-title {{ color: #FFFFFF; font-weight: 700; font-size: 1.15rem; }}
    .brand-sub {{ color: #94A3B8 !important; font-size: .85rem; margin: 0 0 .4rem; }}
    .eyebrow {{ color: #1D4ED8; font-size: .8rem; font-weight: 700; letter-spacing: .1em; text-transform: uppercase; margin-bottom: .2rem; }}
    .page-lead {{ color: #334155; font-size: 1.05rem; margin: -.4rem 0 1.2rem; max-width: 900px; }}
    .capa {{ background: linear-gradient(120deg, #0F172A 0%, #1E3A8A 100%); border-radius: 14px; padding: 2.2rem 2.4rem; color: #fff; margin-bottom: 1.4rem; }}
    .capa .eyebrow {{ color: #93C5FD; }}
    .capa h1 {{ color: #FFFFFF !important; font-size: 2.1rem !important; margin: .2rem 0 .4rem; line-height: 1.2; }}
    .capa .sub {{ color: #CBD5E1; font-size: 1.1rem; margin: 0 0 1.1rem; }}
    .capa .meta {{ color: #E2E8F0; font-size: .98rem; margin: .15rem 0; }}
    .capa .meta b {{ color: #FFFFFF; }}
    .kpi {{ background: #FFFFFF; border: 1px solid #E2E8F0; border-top: 4px solid #1D4ED8; border-radius: 10px; padding: .9rem 1.1rem 1rem;
           min-height: 128px; box-shadow: 0 1px 2px rgba(15,23,42,.06); margin-bottom: 1.1rem; }}
    .kpi-label {{ color: #475569; font-size: .92rem; font-weight: 600; line-height: 1.3; }}
    .kpi-value {{ color: #0F172A; font-size: 2.3rem; font-weight: 800; line-height: 1.1; margin-top: .35rem; }}
    .kpi-text {{ color: #0F172A; font-size: 1.2rem; font-weight: 700; line-height: 1.25; margin-top: .45rem; }}
    .achados {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: .9rem; margin: .4rem 0 1rem; }}
    .achado {{ background: #F8FAFC; border: 1px solid #E2E8F0; border-left: 5px solid #1D4ED8; border-radius: 10px; padding: 1rem 1.1rem; }}
    .achado .num {{ font-size: 2rem; font-weight: 800; color: #1D4ED8; line-height: 1; }}
    .achado .txt {{ color: #0F172A; font-size: 1.02rem; margin-top: .4rem; line-height: 1.35; }}
    .achado .ref {{ color: #64748B; font-size: .82rem; margin-top: .45rem; }}
    .pergunta {{ background: #F8FAFC; border: 1px solid #E2E8F0; border-left: 6px solid #1D4ED8; border-radius: 10px; padding: 1rem 1.2rem; margin: .6rem 0 1rem; }}
    .pergunta h2 {{ font-size: 1.35rem !important; margin: .15rem 0 .35rem; line-height: 1.3; }}
    .pergunta p {{ margin: 0; color: #475569; font-size: .98rem; }}
    .caixa {{ background: #F8FAFC; border: 1px solid #E2E8F0; border-left: 5px solid #059669; border-radius: 10px; padding: .9rem 1.1rem; color: #0F172A; font-size: 1.02rem; line-height: 1.45; margin: .3rem 0 .8rem; }}
    .caixa.autores {{ border-left-color: #1D4ED8; }}
    .caixa b.titulo {{ display: block; font-size: .8rem; letter-spacing: .08em; text-transform: uppercase; color: #475569; margin-bottom: .3rem; }}
    .legenda-respostas {{ color: #334155; font-size: .95rem; line-height: 1.5; }}
    .grafico-titulo {{ margin: .6rem 0 -.2rem; line-height: 1.35; }}
    .grafico-titulo b {{ font-size: 1.2rem; color: #0F172A; }}
    .grafico-titulo span {{ font-size: .92rem !important; }}
    .rodape {{ color: #64748B; font-size: .85rem; border-top: 1px solid #E2E8F0; margin-top: 1.6rem; padding-top: .7rem; }}
    [data-testid="stRadio"] [role="radiogroup"] {{ gap: .4rem; flex-wrap: wrap; }}
    [data-testid="stRadio"] [role="radiogroup"] label {{ padding: .45rem .9rem !important; border: 1px solid #CBD5E1; border-radius: 999px; background: #FFFFFF; }}
    [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) {{ background: #1D4ED8 !important; border-color: #1D4ED8 !important; }}
    [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) p {{ color: #FFFFFF !important; }}
    [data-testid="stRadio"] [role="radiogroup"] label > div:first-child {{ display: none; }}
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# Questionário e rótulos
# -----------------------------------------------------------------------------
PERGUNTAS = {
    "pergunta_1": "1. Você utiliza ferramentas de Inteligência Artificial para estudar ou realizar tarefas escolares?",
    "pergunta_2": "2. Quais ferramentas de IA você mais utiliza na sua rotina de estudos? (Marque todas que se aplicam)",
    "pergunta_3": "3. Com qual finalidade principal você costuma usar a IA nos estudos?",
    "pergunta_4": "4. Com qual frequência você utiliza as plataformas com IA do estado (ex: Redação Paraná, Fluência Paraná)?",
    "pergunta_5": "5. Como os seus professores costumam reagir ao saber que os alunos usam IA nos trabalhos?",
    "pergunta_6": "6. Algum professor já passou uma atividade ou trabalho que EXIGIA o uso de ferramentas de IA?",
    "pergunta_7": "7. Você sente que o uso da IA diminui sua capacidade de pensar por conta própria?",
    "pergunta_8": "8. Na sua opinião, uma ferramenta de IA pode explicar a matéria melhor que um professor?",
}

# Somente a pergunta 2 permite marcar várias opções. As demais podem ter
# vírgulas no texto da resposta, mas são de escolha única.
PERGUNTAS_MULTIPLAS = {"pergunta_2"}

# Cada tema analisado pelas páginas corresponde a uma pergunta numerada.
CAMPOS_PERGUNTA = {
    "frequencia": "pergunta_1",
    "ferramentas": "pergunta_2",
    "finalidade": "pergunta_3",
    "plataformas": "pergunta_4",
    "reacao_professor": "pergunta_5",
    "atividade_ia": "pergunta_6",
    "pensamento": "pergunta_7",
    "professor": "pergunta_8",
}

# Rótulos curtos para os gráficos; o texto completo continua no tooltip e na legenda.
ROTULOS_CURTOS = {
    "Inteligências Artificiais generativas (ChatGPT, Gemini, Copilot)": "IA generativa (ChatGPT, Gemini…)",
    "Plataformas estaduais/educacionais (Redação Paraná, Fluência Paraná, etc.)": "Plataformas do Estado",
    "Editores e geradores visuais/áudio (Canva, NotebookLM, geradores de áudio/slides)": "Editores e geradores (Canva…)",
    "Frequentemente (3 a 4 vezes por semana)": "Frequentemente (3–4x/sem.)",
    "Às vezes (1 a 2 vezes por semana)": "Às vezes (1–2x/sem.)",
    "Raramente (1 a 2 vezes por mês)": "Raramente (1–2x/mês)",
    "Resolver exercícios, conferir respostas e praticar para avaliações": "Resolver exercícios e praticar",
    "Economizar meu tempo e livrar as tarefas logo": "Economizar tempo",
    "Obter ideias, temas e estruturar rascunhos de trabalhos ou redações": "Obter ideias e rascunhos",
    "Tirar dúvidas e resumir e revisar matérias": "Tirar dúvidas e revisar",
    "Apenas quando há necessidade pelos professores": "Só quando o professor pede",
    "Toda semana nas aulas ou tarefas": "Toda semana",
    "Raramente utilizo": "Raramente",
    "Proíbem totalmente e penalizam quem usa": "Proíbem e penalizam",
    "Incentivam o uso correto e ensinam a pesquisar melhor": "Incentivam o uso correto",
    "Aceitam, desde que o aluno não faça apenas": "Aceitam, sem copia e cola",
    "A maioria não comenta sobre o assunto": "Não comentam",
    "Sim, raramente (1 vez)": "Raro (1 vez)",
    "Sim, pontualmente (2 ou mais vezes)": "Pontual (2+ vezes)",
    "Não, ela serve apenas como um apoio aos meus estudos": "Não, é só um apoio",
    "Um pouco, mas tento não depender 100% dela": "Um pouco",
    "Sim, percebo que fico mais preguiçoso para raciocinar": "Sim, fico mais preguiçoso",
    "Não, o professor explica de forma mais clara": "Não, o professor explica melhor",
    "Parcialmente": "Parcialmente",
    "Sim, para aprender a matéria e tirar dúvidas": "Sim, a IA explica melhor",
    "Não sei opinar": "Não sei opinar",
    "Apenas IA's Generativas": "Só IA generativa",
    "Apenas Plataformas Estaduais": "Só plataformas do Estado",
    "Apenas IA's de Ferramentas de Edição": "Só editores",
    "Uso de IA's Generativas, Plataformas e Editores": "IA generativa + plataformas + editores",
    "Uso de IA's Generativas e Plataformas do Estado": "IA generativa + plataformas",
    "Uso de IA's Generativas e em Editores": "IA generativa + editores",
    "Uso de Plataformas e Editores": "Plataformas + editores",
    "Ensino Fundamental Inicial": "6º e 7º ano",
    "Ensino Fundamental Final": "8º e 9º ano",
    "Ensino Médio Inicial": "1ª e 2ª série EM",
    "Ensino Médio Final": "3ª série EM",
}

# Perguntas com escala ordenada: o gráfico segue a escala, não o ranking.
ORDEM_RESPOSTAS = {
    "pergunta_1": ["Não utilizo", "Raramente (1–2x/mês)", "Às vezes (1–2x/sem.)", "Frequentemente (3–4x/sem.)", "Sim, com frequência", "Sempre"],
    "pergunta_4": ["Toda semana", "Só quando o professor pede", "Raramente", "Não utilizo"],
    "pergunta_6": ["Sim, com frequência", "Pontual (2+ vezes)", "Raro (1 vez)", "Nunca"],
    "pergunta_7": ["Não, é só um apoio", "Um pouco", "Sim, fico mais preguiçoso", "Não utilizo"],
    "pergunta_8": ["Sim, a IA explica melhor", "Parcialmente", "Não, o professor explica melhor", "Não sei opinar"],
}

ORDEM_SERIES = ["fundamental_inicial", "fundamental_final", "medio_inicial", "medio_final"]

COR_DESTAQUE = COLORS["blue"]
COR_BASE = "#A9BCE3"
CORES_PENSAMENTO = {"Não, é só um apoio": "#059669", "Um pouco": "#F59E0B", "Sim, fico mais preguiçoso": "#DC2626"}


# -----------------------------------------------------------------------------
# Dados: leitura e limpeza da planilha
# -----------------------------------------------------------------------------
def normalizar_nome(nome: object) -> str:
    """Converte um texto para um identificador comparável (sem acentos e símbolos)."""
    nome = unicodedata.normalize("NFKD", str(nome).strip().lower())
    nome = "".join(caractere for caractere in nome if not unicodedata.combining(caractere))
    nome = re.sub(r"[^a-z0-9 ]", "", nome)
    return re.sub(r"\s+", "_", nome)


_ROTULOS_NORMALIZADOS = [(normalizar_nome(original), curto) for original, curto in ROTULOS_CURTOS.items()]


def rotulo_curto(valor: object) -> str:
    """Rótulo curto e legível para gráficos e cards."""
    texto = re.sub(r"\s+", " ", str(valor)).strip()
    normalizado = normalizar_nome(texto)
    for original, curto in _ROTULOS_NORMALIZADOS:
        if normalizado.startswith(original):
            return curto
    return texto.rstrip(".")


def nome_escola(valor: object) -> str:
    """Abrevia "Colégio Estadual" e "Escola Estadual" para caber nos gráficos."""
    texto = re.sub(r"\s+", " ", str(valor)).strip()
    texto = re.sub(r"^Col[eé]gio Estadual\b", "C. E.", texto, flags=re.IGNORECASE)
    return re.sub(r"^Escola Estadual\b", "E. E.", texto, flags=re.IGNORECASE)


FAIXAS_ETARIAS = ["11 a 12 anos", "13 a 14 anos", "15 a 16 anos", "17 anos ou mais"]


def faixa_etaria(serie: object) -> str:
    """Faixa etária esperada para a série (classificação criada na tratativa dos dados)."""
    posicao = ordem_serie(serie)
    return FAIXAS_ETARIAS[posicao] if posicao < len(FAIXAS_ETARIAS) else "Não informado"


def ordem_serie(valor: object) -> int:
    """Posição da etapa escolar, do 6º ano do Fundamental ao 3º do Médio."""
    nome = normalizar_nome(valor)
    return next((posicao for posicao, chave in enumerate(ORDEM_SERIES) if chave in nome), len(ORDEM_SERIES))


def limpar_resposta(valor: object) -> str:
    """Remove espaços extras e unifica as variações de "Não utilizo"."""
    if pd.isna(valor) or not str(valor).strip():
        return "Não informado"
    texto = re.sub(r"\s+", " ", str(valor)).strip()
    return "Não utilizo" if normalizar_nome(texto).startswith("nao_utilizo") else texto


def unificar_nomes(valores: pd.Series) -> pd.Series:
    """Junta grafias da mesma escola que diferem só em maiúsculas ou acentos.

    Ex.: "Colégio estadual Leôncio Correia" e "Colégio Estadual Leôncio Correia"
    viram uma única instituição, usando a grafia mais frequente.
    """
    chaves = valores.map(normalizar_nome)
    canonico = valores.groupby(chaves).agg(lambda grupo: grupo.value_counts().index[0])

    def capitalizar(nome: str) -> str:
        nome = re.sub(r"^col[eé]gio estadual\b", "Colégio Estadual", nome, flags=re.IGNORECASE)
        nome = re.sub(r"^escola estadual\b", "Escola Estadual", nome, flags=re.IGNORECASE)
        return " ".join(palavra[:1].upper() + palavra[1:] if len(palavra) > 3 and palavra.islower() else palavra for palavra in nome.split())

    return chaves.map(canonico).map(capitalizar)


def categoria_ferramentas(valor: object) -> str:
    """Classifica a combinação de ferramentas marcadas na pergunta 2 (7 categorias)."""
    marcadas = {rotulo_curto(opcao) for opcao in opcoes_multiplas(valor)}
    generativa = any(item.startswith("IA generativa") for item in marcadas)
    plataformas = "Plataformas do Estado" in marcadas
    editores = any(item.startswith("Editores") for item in marcadas)
    combinacoes = {
        (True, False, False): "Só IA generativa", (False, True, False): "Só plataformas do Estado", (False, False, True): "Só editores",
        (True, True, False): "IA generativa + plataformas", (True, False, True): "IA generativa + editores",
        (False, True, True): "Plataformas + editores", (True, True, True): "IA generativa + plataformas + editores",
    }
    return combinacoes.get((generativa, plataformas, editores), "Não informado")


def localizar_coluna(dataframe: pd.DataFrame, criterio) -> object | None:
    """Primeira coluna cujo nome normalizado atende ao critério."""
    return next((coluna for coluna in dataframe.columns if criterio(normalizar_nome(coluna))), None)


def padronizar_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Converte a planilha do Forms para o modelo interno usado pelas páginas.

    Escola e série são as colunas não numeradas do formulário; as perguntas
    são localizadas pelo número inicial ("1.", "2.", ...).
    """
    sem_numero = lambda nome: not nome[:1].isdigit()
    col_escola = localizar_coluna(dataframe, lambda nome: sem_numero(nome) and any(termo in nome for termo in ("instituicao", "escola", "colegio")))
    col_serie = localizar_coluna(dataframe, lambda nome: sem_numero(nome) and ("serie" in nome or "turma" in nome or "ano" in nome.split("_")))
    resultado = pd.DataFrame(index=dataframe.index)
    ausentes: list[str] = []
    for destino, coluna in (("escola", col_escola), ("serie", col_serie)):
        if coluna is None:
            ausentes.append(destino)
            resultado[destino] = "Não informado"
        else:
            resultado[destino] = unificar_nomes(dataframe[coluna].map(limpar_resposta))
    for numero in range(1, 9):
        chave = f"pergunta_{numero}"
        coluna = localizar_coluna(dataframe, lambda nome: nome.startswith(f"{numero}_"))
        if coluna is None:
            ausentes.append(chave)
            resultado[chave] = "Não informado"
        else:
            resultado[chave] = dataframe[coluna].map(limpar_resposta)
    for campo, chave in CAMPOS_PERGUNTA.items():
        resultado[campo] = resultado[chave]
    resultado["faixa_etaria"] = resultado["serie"].map(faixa_etaria)
    resultado["etapa"] = resultado["serie"].map(lambda valor: "Ensino Médio" if "medio" in normalizar_nome(valor) else "Ensino Fundamental")
    # Categoria de combinação de ferramentas (pergunta 2): usa a coluna criada na
    # aba de tratativas quando existe; senão, deriva das opções marcadas.
    col_categoria = localizar_coluna(dataframe, lambda nome: sem_numero(nome) and "categoria" in nome)
    if col_categoria is not None:
        resultado["categoria_uso"] = dataframe[col_categoria].map(limpar_resposta).map(rotulo_curto)
    else:
        resultado["categoria_uso"] = resultado["pergunta_2"].map(categoria_ferramentas)
    resultado = resultado.reset_index(drop=True)
    resultado.attrs["colunas_ausentes"] = ausentes
    return resultado


# Aumente quando mudar a limpeza dos dados, para descartar o cache antigo.
VERSAO_DADOS = 3


@st.cache_data(show_spinner=False)
def ler_planilha(caminho: str, modificado_em: float, versao: int = VERSAO_DADOS) -> pd.DataFrame:
    """Lê a aba de respostas; o horário de modificação e a versão invalidam o cache."""
    abas = pd.read_excel(caminho, sheet_name=None, engine="openpyxl")
    # Usa uma única aba com todas as respostas: "Tratativas" (base tratada pelos
    # autores) ou, na falta dela, "Principal". Somar abas repetiria respostas.
    preferidas = {normalizar_nome(nome): tabela for nome, tabela in abas.items()}
    escolhida = preferidas.get("tratativas", preferidas.get("principal"))
    tabelas = [escolhida] if escolhida is not None else list(abas.values())
    tabelas = [tabela.dropna(how="all") for tabela in tabelas if not tabela.dropna(how="all").empty]
    if not tabelas:
        raise ValueError("Nenhuma aba com respostas foi encontrada")
    return padronizar_dataframe(pd.concat(tabelas, ignore_index=True))


def carregar_dados() -> pd.DataFrame:
    """Carrega a base Total_master.xlsx. Sem ela, nenhum resultado é exibido."""
    arquivo_master = Path(__file__).with_name("Total_master.xlsx")
    if not arquivo_master.exists():
        st.error(f"A planilha **Total_master.xlsx** não foi encontrada. Coloque o arquivo na pasta `{arquivo_master.parent}` e recarregue a página.")
        st.stop()
    try:
        dados = ler_planilha(str(arquivo_master), arquivo_master.stat().st_mtime, VERSAO_DADOS)
    except Exception as erro:  # Erro amigável na interface, sem derrubar o app.
        st.error(f"Não foi possível ler a planilha Total_master.xlsx: {erro}")
        st.stop()
    if dados.empty:
        st.warning("A planilha Total_master.xlsx não contém respostas.")
        st.stop()
    return dados


def opcoes(dataframe: pd.DataFrame, coluna: str) -> list[str]:
    """Valores distintos de uma coluna, para filtros."""
    return sorted(dataframe[coluna].dropna().astype(str).unique().tolist())


# -----------------------------------------------------------------------------
# Cálculos
# -----------------------------------------------------------------------------
def opcoes_multiplas(valor: object) -> list[str]:
    """Separa as opções marcadas numa resposta de múltipla escolha.

    As opções conhecidas são procuradas pelo texto, porque elas mesmas contêm
    vírgulas dentro dos parênteses (ex.: "Redação Paraná, Fluência Paraná").
    """
    texto = re.sub(r"\s+", " ", str(valor)).strip()
    encontradas = [original for original in ROTULOS_CURTOS if original.lower() in texto.lower()]
    if encontradas:
        return encontradas
    return [parte.strip() for parte in texto.split(";" if ";" in texto else ",") if parte.strip()]


def resumir_pergunta(dataframe: pd.DataFrame, chave: str) -> tuple[pd.DataFrame, bool]:
    """Contagem e percentual de cada resposta de uma pergunta.

    Retorna colunas resposta (rótulo curto), completo, quantidade e percentual.
    O percentual usa sempre o número de respondentes como denominador; na
    pergunta de múltipla escolha a soma, portanto, passa de 100%.
    """
    multipla = chave in PERGUNTAS_MULTIPLAS
    total = max(len(dataframe), 1)
    respostas = dataframe[chave].fillna("Não informado")
    if multipla:
        respostas = respostas.map(opcoes_multiplas).explode()
    contagem = respostas.astype(str).value_counts()
    resumo = pd.DataFrame({"completo": contagem.index, "quantidade": contagem.values})
    resumo["resposta"] = resumo["completo"].map(rotulo_curto)
    resumo = resumo.groupby("resposta", as_index=False).agg(completo=("completo", "first"), quantidade=("quantidade", "sum"))
    resumo["percentual"] = (resumo["quantidade"] / total * 100).round(1)
    ordem = ORDEM_RESPOSTAS.get(chave)
    if ordem:
        resumo["_ordem"] = resumo["resposta"].map(lambda valor: ordem.index(valor) if valor in ordem else len(ordem))
        resumo = resumo.sort_values(["_ordem", "quantidade"], ascending=[True, False]).drop(columns="_ordem")
    else:
        resumo = resumo.sort_values("quantidade", ascending=False)
    return resumo.reset_index(drop=True), multipla


def indice_frequencia(respostas: pd.Series) -> pd.Series:
    """Escala de uso da IA (pergunta 1): 0 = não utiliza ... 4 = sempre."""
    valores = respostas.fillna("").astype(str).map(normalizar_nome)
    return valores.map(lambda valor: 4 if "sempre" in valor else 3 if "frequentemente" in valor or "sim_com_frequencia" in valor else 2 if "vezes" in valor else 1 if "raramente" in valor else 0)


def percentual_uso_frequente(respostas: pd.Series) -> float:
    """Percentual que usa IA ao menos 3 vezes por semana ("Sempre" ou "Frequentemente")."""
    if respostas.empty:
        return 0.0
    return float((indice_frequencia(respostas) >= 3).mean() * 100)


def percentual_que_comeca(respostas: pd.Series, *prefixos: str) -> float:
    """Percentual de respostas cujo texto normalizado começa com um dos prefixos."""
    if respostas.empty:
        return 0.0
    normalizadas = respostas.fillna("").astype(str).map(normalizar_nome)
    return float(normalizadas.map(lambda valor: valor.startswith(prefixos)).mean() * 100)


def pct(valor: float) -> str:
    """Formata percentual sem casas decimais desnecessárias (45% ou 45,5%)."""
    texto = f"{valor:.0f}" if abs(valor - round(valor)) < 0.05 else f"{valor:.1f}"
    return texto.replace(".", ",") + "%"


INDICADORES = {
    "Usam IA 3+ vezes por semana": lambda grupo: percentual_uso_frequente(grupo["frequencia"]),
    "Sentem que a IA reduz o raciocínio": lambda grupo: percentual_que_comeca(grupo["pensamento"], "sim", "um_pouco"),
    "Professores proíbem a IA": lambda grupo: percentual_que_comeca(grupo["reacao_professor"], "proibem"),
    "Acham que a IA explica melhor (sim ou parcialmente)": lambda grupo: percentual_que_comeca(grupo["professor"], "sim", "parcialmente"),
    "Usam plataformas do Estado toda semana": lambda grupo: percentual_que_comeca(grupo["plataformas"], "toda_semana"),
}


# -----------------------------------------------------------------------------
# Componentes visuais
# -----------------------------------------------------------------------------
def cabecalho(secao: str, titulo_pagina: str, descricao: str = "") -> None:
    st.markdown(f'<div class="eyebrow">{secao}</div>', unsafe_allow_html=True)
    st.title(titulo_pagina)
    if descricao:
        st.markdown(f'<p class="page-lead">{descricao}</p>', unsafe_allow_html=True)


def kpi(rotulo: str, valor: str, cor: str = COLORS["blue"], texto: bool = False) -> None:
    classe = "kpi-text" if texto else "kpi-value"
    st.markdown(f'<div class="kpi" style="border-top-color:{cor}"><div class="kpi-label">{rotulo}</div><div class="{classe}">{valor}</div></div>', unsafe_allow_html=True)


def caixa(titulo_caixa: str, conteudo: str, autores: bool = False) -> None:
    classe = "caixa autores" if autores else "caixa"
    st.markdown(f'<div class="{classe}"><b class="titulo">{titulo_caixa}</b>{conteudo}</div>', unsafe_allow_html=True)


def rodape(dataframe: pd.DataFrame) -> None:
    periodo = f" em {PERIODO_COLETA}" if PERIODO_COLETA else ""
    st.markdown(f'<div class="rodape">Fonte: questionário aplicado{periodo} a {len(dataframe)} estudantes de {dataframe["escola"].nunique()} escolas estaduais de Curitiba – PR. Elaboração dos autores.</div>', unsafe_allow_html=True)


def mostrar(figura: go.Figure) -> None:
    """Exibe o gráfico com título em HTML (mesma fonte da página, sem cortes)."""
    titulo = figura.layout.title.text
    if titulo:
        st.markdown(f"<div class='grafico-titulo'>{titulo}</div>", unsafe_allow_html=True)
        figura.update_layout(title=None, margin=dict(t=20))
    st.plotly_chart(figura, width="stretch", config=PLOTLY_CONFIG)


def quebrar(texto: object, limite: int = 22) -> str:
    """Quebra rótulos longos em linhas para o eixo dos gráficos."""
    linhas: list[str] = []
    linha = ""
    for palavra in str(texto).split():
        if linha and len(linha) + len(palavra) + 1 > limite:
            linhas.append(linha)
            linha = palavra
        else:
            linha = f"{linha} {palavra}".strip()
    if linha:
        linhas.append(linha)
    return "<br>".join(linhas)


def titulo_grafico(titulo: str, subtitulo: str) -> dict:
    return dict(text=f"<b>{titulo}</b><br><span style='font-size:14px;color:#64748B'>{subtitulo}</span>", x=0, xanchor="left")


def subtitulo_amostra(n: int, multipla: bool = False) -> str:
    texto = f"n = {n} respondentes"
    return texto + " · múltipla escolha" if multipla else texto


def cores_destaque(resumo: pd.DataFrame) -> list[str]:
    """Destaca a resposta mais escolhida; as demais ficam em tom claro."""
    maior = resumo["quantidade"].max()
    return [COR_DESTAQUE if quantidade == maior else COR_BASE for quantidade in resumo["quantidade"]]


def grafico_barras(resumo: pd.DataFrame, titulo: str, subtitulo: str, cores: list[str] | None = None) -> go.Figure:
    """Barras horizontais com percentual e quantidade escritos em cada barra."""
    figura = go.Figure(go.Bar(
        x=resumo["percentual"], y=resumo["resposta"].map(quebrar), orientation="h",
        marker=dict(color=cores or cores_destaque(resumo)),
        text=[f"<b>{pct(p)}</b>  ({q})" for p, q in zip(resumo["percentual"], resumo["quantidade"])],
        textposition="outside", cliponaxis=False,
        customdata=resumo[["completo", "quantidade"]],
        hovertemplate="<b>%{customdata[0]}</b><br>%{x:.1f}% · %{customdata[1]} respostas<extra></extra>",
    ))
    figura.update_yaxes(autorange="reversed", showgrid=False)
    figura.update_xaxes(visible=False, range=[0, max(resumo["percentual"].max() * 1.25, 10)])
    figura.update_layout(title=titulo_grafico(titulo, subtitulo), bargap=.32, showlegend=False)
    return apply_layout(figura, 150 + 62 * len(resumo))


def grafico_colunas(resumo: pd.DataFrame, titulo: str, subtitulo: str) -> go.Figure:
    """Barras verticais com rótulos quebrados."""
    figura = go.Figure(go.Bar(
        x=resumo["resposta"].map(lambda valor: quebrar(valor, 16)), y=resumo["percentual"],
        marker=dict(color=cores_destaque(resumo)),
        text=[f"<b>{pct(p)}</b><br>({q})" for p, q in zip(resumo["percentual"], resumo["quantidade"])],
        textposition="outside", cliponaxis=False, customdata=resumo["completo"],
        hovertemplate="<b>%{customdata}</b><br>%{y:.1f}%<extra></extra>",
    ))
    figura.update_yaxes(visible=False, range=[0, resumo["percentual"].max() * 1.3])
    figura.update_layout(title=titulo_grafico(titulo, subtitulo), bargap=.35, showlegend=False)
    return apply_layout(figura, 500)


def grafico_rosca(resumo: pd.DataFrame, titulo: str, subtitulo: str, n: int) -> go.Figure:
    """Gráfico de rosca para perguntas de escolha única."""
    cores = [COR_DESTAQUE, "#60A5FA", "#F59E0B", "#10B981", "#A78BFA", "#94A3B8"]
    figura = go.Figure(go.Pie(
        labels=resumo["resposta"], values=resumo["quantidade"], hole=.6, sort=False, direction="clockwise",
        marker=dict(colors=cores[:len(resumo)], line=dict(color="white", width=2)),
        textinfo="percent", textfont=dict(size=17), customdata=resumo["completo"],
        hovertemplate="<b>%{customdata}</b><br>%{percent} · %{value} respostas<extra></extra>",
    ))
    figura.update_layout(title=titulo_grafico(titulo, subtitulo), showlegend=True,
                         legend=dict(orientation="v", x=1.02, y=.5, font=dict(size=15)),
                         annotations=[dict(text=f"<b>{n}</b><br>respondentes", x=.5, y=.5, xref="paper", yref="paper", showarrow=False, font=dict(size=18, color=COLORS["title"]))])
    return apply_layout(figura, 480)


def grafico_pergunta(dataframe: pd.DataFrame, chave: str, titulo: str) -> None:
    """Atalho: resumo + barras horizontais de uma pergunta."""
    resumo, multipla = resumir_pergunta(dataframe, chave)
    mostrar(grafico_barras(resumo, titulo, subtitulo_amostra(len(dataframe), multipla)))


def titulo_destaque(resumo: pd.DataFrame, complemento: str = "") -> str:
    """Título que já anuncia o resultado principal (ex.: "45% – Não, o professor explica melhor")."""
    principal = resumo.loc[resumo["quantidade"].idxmax()]
    return f"{pct(principal['percentual'])} {complemento}“{principal['resposta']}”"


# -----------------------------------------------------------------------------
# Páginas
# -----------------------------------------------------------------------------
def pagina_inicio(dataframe: pd.DataFrame) -> None:
    metas = [f"<p class='meta'><b>Autores:</b> {AUTORES}</p>"]
    if ORIENTADOR:
        metas.append(f"<p class='meta'><b>Orientação:</b> {ORIENTADOR}</p>")
    curso = " · ".join(parte for parte in (CURSO, INSTITUICAO) if parte)
    if curso:
        metas.append(f"<p class='meta'>{curso}</p>")
    st.markdown(f"<div class='capa'><div class='eyebrow'>Trabalho de Conclusão de Curso</div><h1>{TITULO_TCC}</h1><p class='sub'>{SUBTITULO_TCC}</p>{''.join(metas)}</div>", unsafe_allow_html=True)

    indicadores = [
        ("Estudantes participantes", str(len(dataframe)), COLORS["blue"]),
        ("Escolas estaduais", str(dataframe["escola"].nunique()), COLORS["purple"]),
        ("Usam IA 3 ou mais vezes por semana", pct(percentual_uso_frequente(dataframe["frequencia"])), COLORS["green"]),
        ("Sentem que a IA reduz o raciocínio", pct(percentual_que_comeca(dataframe["pensamento"], "sim", "um_pouco")), COLORS["orange"]),
    ]
    for coluna, (rotulo, valor, cor) in zip(st.columns(4), indicadores):
        with coluna:
            kpi(rotulo, valor, cor)

    st.subheader("Principais achados")
    finalidade = resumir_pergunta(dataframe, "pergunta_3")[0].iloc[0]
    ferramentas = resumir_pergunta(dataframe, "pergunta_2")[0]
    generativa = ferramentas.loc[ferramentas["resposta"].str.startswith("IA generativa"), "percentual"].sum()
    achados = [
        (pct(generativa), "citaram IA generativa (ChatGPT, Gemini, Copilot) entre as ferramentas que mais usam.", "Pergunta 2"),
        (pct(finalidade["percentual"]), f"usam a IA principalmente para “{finalidade['resposta'].lower()}”.", "Pergunta 3"),
        (pct(percentual_que_comeca(dataframe["reacao_professor"], "proibem")), "relatam que os professores proíbem e penalizam o uso de IA.", "Pergunta 5"),
        (pct(percentual_que_comeca(dataframe["pensamento"], "sim", "um_pouco")), "sentem que a IA diminui, ao menos um pouco, a capacidade de pensar por conta própria.", "Pergunta 7"),
        (pct(percentual_que_comeca(dataframe["professor"], "sim")), "acham que a IA explica a matéria melhor que o professor.", "Pergunta 8"),
    ]
    cartoes = "".join(f"<div class='achado'><div class='num'>{numero}</div><div class='txt'>{texto}</div><div class='ref'>{referencia}</div></div>" for numero, texto, referencia in achados)
    st.markdown(f"<div class='achados'>{cartoes}</div>", unsafe_allow_html=True)
    rodape(dataframe)


def pagina_perfil(dataframe: pd.DataFrame) -> None:
    cabecalho("Amostra", "Perfil dos participantes", "Distribuição dos estudantes por etapa escolar e por instituição de ensino.")
    n = len(dataframe)
    series = dataframe["serie"].value_counts().rename_axis("completo").reset_index(name="quantidade")
    series["resposta"] = [f"{rotulo_curto(serie)} · {faixa_etaria(serie)}" for serie in series["completo"]]
    series["percentual"] = (series["quantidade"] / n * 100).round(1)
    series = series.assign(_ordem=series["completo"].map(ordem_serie)).sort_values("_ordem").drop(columns="_ordem").reset_index(drop=True)
    escolas = dataframe["escola"].value_counts().rename_axis("completo").reset_index(name="quantidade")
    escolas["resposta"] = escolas["completo"].map(nome_escola)
    escolas["percentual"] = (escolas["quantidade"] / n * 100).round(1)
    esquerda, direita = st.columns([1, 1.35])
    with esquerda:
        mostrar(grafico_barras(series, "Participantes por etapa escolar", subtitulo_amostra(n), cores=[COR_DESTAQUE] * len(series)))
    with direita:
        mostrar(grafico_barras(escolas, "Participantes por instituição", subtitulo_amostra(n), cores=[COLORS["purple"]] * len(escolas)))
    pequenas = escolas[escolas["quantidade"] < 10]
    if not pequenas.empty:
        st.caption(f"{len(pequenas)} instituição(ões) com menos de 10 respondentes: comparações envolvendo essas escolas devem ser lidas com cautela.")
    rodape(dataframe)


def pagina_perguntas(dataframe: pd.DataFrame) -> None:
    """Análise individual de cada pergunta do questionário."""
    cabecalho("Questionário", "Perguntas e Respostas")
    controles = st.columns([2.4, 1])
    with controles[0]:
        escolhida = st.selectbox("Pergunta", list(PERGUNTAS), format_func=lambda chave: PERGUNTAS[chave])
    with controles[1]:
        filtro_escola = st.selectbox("Instituição", ["Todas"] + opcoes(dataframe, "escola"), format_func=lambda valor: valor if valor == "Todas" else nome_escola(valor))
    filtrado = dataframe if filtro_escola == "Todas" else dataframe[dataframe["escola"] == filtro_escola]
    resumo, multipla = resumir_pergunta(filtrado, escolhida)
    n = len(filtrado)
    recorte = "todas as instituições" if filtro_escola == "Todas" else nome_escola(filtro_escola)
    st.markdown(f"<div class='pergunta'><div class='eyebrow'>Pergunta {escolhida[-1]}</div><h2>{PERGUNTAS[escolhida][3:]}</h2><p>{n} respondentes · {recorte} · {'múltipla escolha' if multipla else 'escolha única'}</p></div>", unsafe_allow_html=True)
    if resumo.empty:
        st.warning("Não há respostas para esta pergunta no recorte selecionado.")
        return

    principal = resumo.loc[resumo["quantidade"].idxmax()]
    colunas = st.columns([1, 2, 1])
    with colunas[0]:
        kpi("Respondentes", str(n), COLORS["blue"])
    with colunas[1]:
        kpi("Resposta mais escolhida", principal["resposta"], COLORS["purple"], texto=True)
    with colunas[2]:
        kpi("Escolheram essa resposta", pct(principal["percentual"]), COLORS["green"])

    tipos = ["Barras horizontais", "Barras verticais"] + ([] if multipla else ["Rosca"])
    tipo = st.radio("Tipo de gráfico", tipos, horizontal=True, key=f"tipo_{escolhida}")
    titulo = f"{pct(principal['percentual'])} {'citaram' if multipla else 'responderam'} “{principal['resposta']}”"
    subtitulo = subtitulo_amostra(n, multipla)
    if tipo == "Barras verticais":
        mostrar(grafico_colunas(resumo, titulo, subtitulo))
    elif tipo == "Rosca":
        mostrar(grafico_rosca(resumo, titulo, subtitulo, n))
    else:
        mostrar(grafico_barras(resumo, titulo, subtitulo))

    esquerda, direita = st.columns(2)
    with esquerda:
        caixa("Leitura dos dados", leitura_automatica(resumo, multipla))
        if INTERPRETACOES.get(escolhida):
            caixa("Interpretação dos autores", INTERPRETACOES[escolhida], autores=True)
    with direita:
        legenda = "<br>".join(f"<b>{linha.resposta}:</b> {linha.completo}" for linha in resumo.itertuples() if linha.resposta != linha.completo)
        if legenda:
            caixa("Texto completo das respostas", f"<div class='legenda-respostas'>{legenda}</div>")
    with st.expander("Tabela de dados"):
        st.dataframe(resumo[["completo", "quantidade", "percentual"]].rename(columns={"completo": "Resposta", "quantidade": "Quantidade", "percentual": "Percentual (%)"}), width="stretch", hide_index=True)
    rodape(dataframe)


def leitura_automatica(resumo: pd.DataFrame, multipla: bool) -> str:
    """Descrição objetiva do resultado, sem interpretar além dos números."""
    ranking = resumo.sort_values("quantidade", ascending=False).reset_index(drop=True)
    primeira = ranking.iloc[0]
    if multipla:
        texto = f"“{primeira['resposta']}” foi citada por <b>{pct(primeira['percentual'])}</b> dos estudantes"
        if len(ranking) > 1:
            texto += f", seguida de “{ranking.iloc[1]['resposta']}” ({pct(ranking.iloc[1]['percentual'])})"
        return texto + ". Como era possível marcar mais de uma opção, os percentuais somam mais de 100%."
    texto = f"A resposta mais frequente foi “{primeira['resposta']}”, com <b>{pct(primeira['percentual'])}</b>"
    if len(ranking) > 1:
        segunda = ranking.iloc[1]
        texto += f", seguida de “{segunda['resposta']}” ({pct(segunda['percentual'])}). Juntas, somam <b>{pct(primeira['percentual'] + segunda['percentual'])}</b> das respostas"
    texto += "."
    if primeira["percentual"] < 50:
        texto += " Nenhuma resposta alcança a maioria absoluta."
    return texto


def pagina_uso(dataframe: pd.DataFrame) -> None:
    cabecalho("Resultados", "Uso da Inteligência Artificial", "Com que frequência os estudantes usam IA, quais ferramentas e quanto usam as plataformas do Estado.")
    uso = percentual_uso_frequente(dataframe["frequencia"])
    colunas = st.columns(3)
    with colunas[0]:
        kpi("Usam IA 3 ou mais vezes por semana", pct(uso), COLORS["blue"])
    with colunas[1]:
        kpi("Não utilizam IA", pct(percentual_que_comeca(dataframe["frequencia"], "nao_utilizo")), COLORS["orange"])
    with colunas[2]:
        kpi("Usam plataformas do Estado toda semana", pct(percentual_que_comeca(dataframe["plataformas"], "toda_semana")), COLORS["green"])
    frequencia, _ = resumir_pergunta(dataframe, "pergunta_1")
    mostrar(grafico_barras(frequencia, f"{pct(uso)} usam IA ao menos 3 vezes por semana", "Pergunta 1 · " + subtitulo_amostra(len(dataframe))))
    esquerda, direita = st.columns(2)
    with esquerda:
        ferramentas, _ = resumir_pergunta(dataframe, "pergunta_2")
        mostrar(grafico_barras(ferramentas, "Ferramentas mais usadas", "Pergunta 2 · " + subtitulo_amostra(len(dataframe), True)))
    with direita:
        plataformas, _ = resumir_pergunta(dataframe, "pergunta_4")
        mostrar(grafico_barras(plataformas, "Uso das plataformas do Estado", "Pergunta 4 · " + subtitulo_amostra(len(dataframe))))
    st.caption("Na pergunta 2, Redação Paraná e Fluência Paraná aparecem numa única opção, por isso não são separadas.")
    categorias = dataframe["categoria_uso"].value_counts().rename_axis("resposta").reset_index(name="quantidade")
    categorias["completo"] = categorias["resposta"]
    categorias["percentual"] = (categorias["quantidade"] / len(dataframe) * 100).round(1)
    combinado = float(dataframe["categoria_uso"].str.contains("+", regex=False).mean() * 100)
    mostrar(grafico_barras(categorias, f"{pct(combinado)} combinam mais de um tipo de ferramenta", "Combinações de ferramentas (pergunta 2) · " + subtitulo_amostra(len(dataframe))))
    rodape(dataframe)


def pagina_finalidade(dataframe: pd.DataFrame) -> None:
    cabecalho("Resultados", "Finalidade do uso", PERGUNTAS["pergunta_3"][3:])
    resumo, _ = resumir_pergunta(dataframe, "pergunta_3")
    mostrar(grafico_barras(resumo, titulo_destaque(resumo, "usam para "), "Pergunta 3 · " + subtitulo_amostra(len(dataframe))))
    caixa("Leitura dos dados", leitura_automatica(resumo, False))
    rodape(dataframe)


def pagina_professores(dataframe: pd.DataFrame) -> None:
    cabecalho("Resultados", "IA e Professores", "Como os professores reagem ao uso de IA e como os estudantes comparam a IA ao professor.")
    n = len(dataframe)
    comparacao, _ = resumir_pergunta(dataframe, "pergunta_8")
    so_sim = percentual_que_comeca(dataframe["professor"], "sim")
    mostrar(grafico_barras(comparacao, f"Só {pct(so_sim)} acham que a IA explica melhor que o professor", "Pergunta 8 · " + subtitulo_amostra(n)))
    esquerda, direita = st.columns(2)
    with esquerda:
        reacao, _ = resumir_pergunta(dataframe, "pergunta_5")
        mostrar(grafico_barras(reacao, "Reação dos professores ao uso de IA", "Pergunta 5 · " + subtitulo_amostra(n)))
    with direita:
        atividade, _ = resumir_pergunta(dataframe, "pergunta_6")
        exigiu = 100 - percentual_que_comeca(dataframe["atividade_ia"], "nunca")
        mostrar(grafico_barras(atividade, f"{pct(exigiu)} já tiveram atividade que exigia IA", "Pergunta 6 · " + subtitulo_amostra(n)))
    rodape(dataframe)


def pagina_pensamento(dataframe: pd.DataFrame) -> None:
    cabecalho("Resultados", "Impacto no pensamento próprio", PERGUNTAS["pergunta_7"][3:])
    indicadores = [
        ("Não, a IA é só um apoio", pct(percentual_que_comeca(dataframe["pensamento"], "nao_ela")), CORES_PENSAMENTO["Não, é só um apoio"]),
        ("Um pouco", pct(percentual_que_comeca(dataframe["pensamento"], "um_pouco")), CORES_PENSAMENTO["Um pouco"]),
        ("Sim, ficam mais preguiçosos", pct(percentual_que_comeca(dataframe["pensamento"], "sim")), CORES_PENSAMENTO["Sim, fico mais preguiçoso"]),
    ]
    for coluna, (rotulo, valor, cor) in zip(st.columns(3), indicadores):
        with coluna:
            kpi(rotulo, valor, cor)
    resumo, _ = resumir_pergunta(dataframe, "pergunta_7")
    cores = [CORES_PENSAMENTO.get(resposta, "#94A3B8") for resposta in resumo["resposta"]]
    reduz = percentual_que_comeca(dataframe["pensamento"], "sim", "um_pouco")
    mostrar(grafico_barras(resumo, f"{pct(reduz)} sentem que a IA reduz, ao menos um pouco, o raciocínio", "Pergunta 7 · " + subtitulo_amostra(len(dataframe)), cores=cores))
    st.subheader("Quem usa mais IA sente mais impacto?")
    mostrar(grafico_cruzamento_uso_pensamento(dataframe))
    st.caption("Cruzamento das perguntas 1 e 7. Exclui quem declarou não utilizar IA. Os percentuais são calculados dentro de cada grupo.")
    rodape(dataframe)


def grafico_cruzamento_uso_pensamento(dataframe: pd.DataFrame) -> go.Figure:
    """Barras 100% empilhadas: percepção (P7) por intensidade de uso (P1)."""
    base = dataframe.assign(indice=indice_frequencia(dataframe["frequencia"]), percepcao=dataframe["pensamento"].map(rotulo_curto))
    base = base[(base["indice"] > 0) & base["percepcao"].isin(CORES_PENSAMENTO)]
    base["grupo"] = base["indice"].map(lambda valor: "Usam 3+ vezes por semana" if valor >= 3 else "Usam até 2 vezes por semana")
    tabela = pd.crosstab(base["grupo"], base["percepcao"], normalize="index") * 100
    contagem = base["grupo"].value_counts()
    grupos = [grupo for grupo in ("Usam até 2 vezes por semana", "Usam 3+ vezes por semana") if grupo in tabela.index]
    rotulos = [f"{grupo}<br><span style='font-size:13px;color:#64748B'>n = {contagem[grupo]}</span>" for grupo in grupos]
    figura = go.Figure()
    for percepcao, cor in CORES_PENSAMENTO.items():
        valores = [float(tabela.loc[grupo, percepcao]) if percepcao in tabela.columns else 0.0 for grupo in grupos]
        figura.add_bar(name=percepcao, y=rotulos, x=valores, orientation="h", marker_color=cor,
                       text=[pct(valor) if valor >= 9 else "" for valor in valores], textposition="inside", insidetextanchor="middle", textangle=0,
                       textfont=dict(color="white", size=16), hovertemplate=f"<b>{percepcao}</b><br>%{{x:.1f}}%<extra></extra>")
    figura.update_layout(barmode="stack", bargap=.35, legend=dict(orientation="h", y=-.12, x=0, font=dict(size=15), traceorder="normal"),
                         title=titulo_grafico("Percepção sobre o raciocínio, por intensidade de uso", "Percentual dentro de cada grupo"))
    figura.update_xaxes(visible=False, range=[0, 100])
    figura.update_yaxes(autorange="reversed")
    return apply_layout(figura, 360)


ETAPAS = ["Ensino Fundamental", "Ensino Médio"]
DESCRICAO_ETAPA = {"Ensino Fundamental": "6º ao 9º ano", "Ensino Médio": "1ª a 3ª série"}
CORES_FREQUENCIA = {"Não utilizo": "#CBD5E1", "Raramente (1–2x/mês)": "#BFDBFE", "Às vezes (1–2x/sem.)": "#93C5FD", "Frequentemente (3–4x/sem.)": "#3B82F6", "Sim, com frequência": "#2563EB", "Sempre": "#1E3A8A"}
CORES_FINALIDADE = {"Resolver exercícios e praticar": "#1D4ED8", "Obter ideias e rascunhos": "#F59E0B", "Economizar tempo": "#DC2626", "Tirar dúvidas e revisar": "#10B981", "Não utilizo": "#CBD5E1"}
CORES_ATIVIDADE = {"Nunca": "#CBD5E1", "Raro (1 vez)": "#BFDBFE", "Pontual (2+ vezes)": "#60A5FA", "Sim, com frequência": "#1D4ED8"}
CORES_CLARAS = {"#CBD5E1", "#BFDBFE", "#93C5FD", "#60A5FA", "#F59E0B"}


def grafico_empilhado(base: pd.DataFrame, linha: str, coluna: str, ordem_linhas: list[str], cores: dict[str, str], titulo: str, subtitulo: str, rotulos_linha: dict[str, str] | None = None) -> tuple[go.Figure, pd.DataFrame]:
    """Barras 100% empilhadas: distribuição de `coluna` dentro de cada grupo de `linha`."""
    contagem = pd.crosstab(base[linha], base[coluna])
    percentuais = contagem.div(contagem.sum(axis=1), axis=0) * 100
    linhas = [valor for valor in ordem_linhas if valor in contagem.index]
    rotulos_linha = rotulos_linha or {}
    eixo = [f"{quebrar(rotulos_linha.get(valor, valor), 26)}<br><span style='font-size:13px;color:#64748B'>n = {int(contagem.loc[valor].sum())}</span>" for valor in linhas]
    figura = go.Figure()
    for categoria, cor in cores.items():
        if categoria not in contagem.columns:
            continue
        valores = [float(percentuais.loc[valor, categoria]) for valor in linhas]
        quantidades = [int(contagem.loc[valor, categoria]) for valor in linhas]
        figura.add_bar(name=categoria, y=eixo, x=valores, orientation="h", marker=dict(color=cor, line=dict(color="white", width=1)),
                       text=[pct(valor) if valor >= 7 else "" for valor in valores], textposition="inside", insidetextanchor="middle", textangle=0,
                       textfont=dict(color="#0F172A" if cor in CORES_CLARAS else "white", size=15), customdata=quantidades,
                       hovertemplate=f"<b>{categoria}</b><br>%{{x:.1f}}% · %{{customdata}} respostas<extra></extra>")
    figura.update_layout(barmode="stack", bargap=.3, title=titulo_grafico(titulo, subtitulo),
                         legend=dict(orientation="h", y=-.08, yanchor="top", x=0, font=dict(size=14), traceorder="normal"))
    figura.update_xaxes(visible=False, range=[0, 100])
    figura.update_yaxes(autorange="reversed")
    tabela = contagem.loc[linhas].copy()
    for categoria in tabela.columns:
        tabela[f"{categoria} (%)"] = percentuais.loc[linhas, categoria].round(1)
    tabela.insert(0, "Total", contagem.loc[linhas].sum(axis=1))
    return apply_layout(figura, 190 + 78 * len(linhas)), tabela.reset_index().rename(columns={linha: "Grupo"})


def mostrar_tabela_cruzamento(tabela: pd.DataFrame) -> None:
    with st.expander("Tabela do cruzamento (quantidades e percentuais na linha)"):
        st.dataframe(tabela, width="stretch", hide_index=True)


def pagina_cruzamentos(dataframe: pd.DataFrame) -> None:
    cabecalho("Análise dos resultados", "Cruzamentos entre perguntas", "Associações descritivas entre as respostas. Os percentuais são calculados dentro de cada grupo e não indicam relação de causa e efeito.")
    secoes = ["4.2 Uso por etapa de ensino", "4.2.1 Frequência × finalidade", "4.2.2 Frequência × dependência", "4.2.3 Reação docente × atividades"]
    secao = st.radio("Cruzamento", secoes, horizontal=True, label_visibility="collapsed")
    base = dataframe.assign(
        freq=dataframe["frequencia"].map(rotulo_curto), fin=dataframe["finalidade"].map(rotulo_curto),
        dep=dataframe["pensamento"].map(rotulo_curto), reacao=dataframe["reacao_professor"].map(rotulo_curto),
        ativ=dataframe["atividade_ia"].map(rotulo_curto),
    )
    rotulos_etapa = {etapa: f"{etapa} ({DESCRICAO_ETAPA[etapa]})" for etapa in ETAPAS}
    por_etapa = {etapa: base[base["etapa"] == etapa] for etapa in ETAPAS}

    if secao == secoes[0]:
        uso = {etapa: percentual_uso_frequente(dados["frequencia"]) for etapa, dados in por_etapa.items()}
        colunas = st.columns(2)
        for coluna, etapa in zip(colunas, ETAPAS):
            with coluna:
                kpi(f"{etapa}: usam IA 3+ vezes por semana", pct(uso[etapa]), COLORS["blue"] if etapa == "Ensino Médio" else COLORS["purple"])
        figura, tabela = grafico_empilhado(base, "etapa", "freq", ETAPAS, CORES_FREQUENCIA, "Frequência de uso de IA por etapa de ensino", "Pergunta 1 × etapa · percentual dentro de cada etapa", rotulos_etapa)
        mostrar(figura)
        caixa("Leitura dos dados", f"Usam IA ao menos 3 vezes por semana (\"Sempre\" ou \"Frequentemente\") <b>{pct(uso['Ensino Médio'])}</b> dos estudantes do Ensino Médio e <b>{pct(uso['Ensino Fundamental'])}</b> dos do Ensino Fundamental.")
    elif secao == secoes[1]:
        ordem = [valor for valor in ORDEM_RESPOSTAS["pergunta_1"] if valor != "Não utilizo"][::-1]
        usuarios = base[base["freq"] != "Não utilizo"]
        figura, tabela = grafico_empilhado(usuarios, "freq", "fin", ordem, CORES_FINALIDADE, "Finalidade principal, por frequência de uso", f"Pergunta 1 × Pergunta 3 · exclui quem não utiliza IA (n = {len(usuarios)})")
        mostrar(figura)
        frequentes = usuarios[indice_frequencia(usuarios["frequencia"]) >= 3]
        principal = frequentes["fin"].value_counts()
        if not principal.empty:
            caixa("Leitura dos dados", f"Entre quem usa IA 3 ou mais vezes por semana (n = {len(frequentes)}), a finalidade mais citada é “{principal.index[0]}” (<b>{pct(principal.iloc[0] / len(frequentes) * 100)}</b>), seguida de “{principal.index[1]}” ({pct(principal.iloc[1] / len(frequentes) * 100)}).")
    elif secao == secoes[2]:
        dependencia = {etapa: percentual_que_comeca(dados["pensamento"], "sim", "um_pouco") for etapa, dados in por_etapa.items()}
        colunas = st.columns(2)
        for coluna, etapa in zip(colunas, ETAPAS):
            with coluna:
                kpi(f"{etapa}: sentem algum grau de dependência", pct(dependencia[etapa]), COLORS["orange"])
        figura, tabela_etapa = grafico_empilhado(base, "etapa", "dep", ETAPAS, CORES_PENSAMENTO, "Percepção de dependência por etapa de ensino", "Pergunta 7 × etapa · percentual dentro de cada etapa", rotulos_etapa)
        mostrar(figura)
        ordem = [valor for valor in ORDEM_RESPOSTAS["pergunta_1"] if valor != "Não utilizo"][::-1]
        usuarios = base[(base["freq"] != "Não utilizo") & base["dep"].isin(CORES_PENSAMENTO)]
        figura, tabela = grafico_empilhado(usuarios, "freq", "dep", ordem, CORES_PENSAMENTO, "Percepção de dependência, por frequência de uso", "Pergunta 1 × Pergunta 7 · exclui quem não utiliza IA")
        mostrar(figura)
        caixa("Leitura dos dados", f"Sentem que a IA diminui, ao menos um pouco, a capacidade de pensar por conta própria (\"Sim\" ou \"Um pouco\") <b>{pct(dependencia['Ensino Médio'])}</b> dos estudantes do Ensino Médio e <b>{pct(dependencia['Ensino Fundamental'])}</b> dos do Ensino Fundamental.")
        mostrar_tabela_cruzamento(tabela_etapa)
    else:
        ordem = ["Proíbem e penalizam", "Não comentam", "Aceitam, sem copia e cola", "Incentivam o uso correto"]
        figura, tabela = grafico_empilhado(base, "reacao", "ativ", ordem, CORES_ATIVIDADE, "Atividades que exigiam IA, por reação dos professores", "Pergunta 5 × Pergunta 6 · percentual dentro de cada reação")
        mostrar(figura)
        partes = []
        for reacao in ordem:
            grupo = base[base["reacao"] == reacao]
            if len(grupo):
                partes.append(f"“{reacao}”: <b>{pct((grupo['ativ'] != 'Nunca').mean() * 100)}</b> já tiveram atividade com IA (n = {len(grupo)})")
        caixa("Leitura dos dados", "<br>".join(partes))
    mostrar_tabela_cruzamento(tabela)
    rodape(dataframe)


# Proficiência média por escola (tabela da seção 4.3 do TCC). None = sem dado.
PROFICIENCIA = {
    "Colégio Estadual do Paraná": {"LP Fundamental": 317.5, "LP Médio": 330.41, "MAT Fundamental": 322.91, "MAT Médio": 324.59},
    "Colégio Estadual Leôncio Correia": {"LP Fundamental": 262.05, "LP Médio": 293.69, "MAT Fundamental": 272.97, "MAT Médio": 291.67},
    "Colégio Estadual Victor do Amaral": {"LP Fundamental": 269.66, "LP Médio": None, "MAT Fundamental": 260.69, "MAT Médio": None},
    "Colégio Estadual Gottlieb Mueller": {"LP Fundamental": 275.69, "LP Médio": 294.17, "MAT Fundamental": 284.67, "MAT Médio": 276.6},
    "Colégio Estadual República Oriental do Uruguai": {"LP Fundamental": 288.52, "LP Médio": 301.65, "MAT Fundamental": 293.17, "MAT Médio": 312.36},
    "Colégio Estadual Paulo Leminsky": {"LP Fundamental": 279.69, "LP Médio": 286.52, "MAT Fundamental": 275.45, "MAT Médio": 277.69},
    "Colégio Estadual Dona Branca": {"LP Fundamental": 268.76, "LP Médio": 313.09, "MAT Fundamental": 272.97, "MAT Médio": 304.63},
}


def pagina_proficiencia(dataframe: pd.DataFrame) -> None:
    cabecalho("Análise exploratória", "Plataformas do Estado e proficiência", "Uso semanal das plataformas do Estado (pergunta 4) comparado à proficiência média das escolas em Língua Portuguesa e Matemática.")
    controles = st.columns(2)
    with controles[0]:
        etapa = st.radio("Etapa", ETAPAS, horizontal=True)
    with controles[1]:
        disciplina = st.radio("Proficiência", ["Língua Portuguesa", "Matemática"], horizontal=True)
    chave = f"{'LP' if disciplina == 'Língua Portuguesa' else 'MAT'} {etapa.split()[-1]}"
    chaves_escola = {normalizar_nome(escola): escola for escola in PROFICIENCIA}
    linhas = []
    for escola, dados in dataframe[dataframe["etapa"] == etapa].groupby("escola"):
        referencia = chaves_escola.get(normalizar_nome(escola))
        valor = PROFICIENCIA[referencia][chave] if referencia else None
        if valor is None:
            continue
        linhas.append({"escola": escola, "n": len(dados), "uso": percentual_que_comeca(dados["plataformas"], "toda_semana"), "proficiencia": valor})
    tabela = pd.DataFrame(linhas)
    if tabela.empty:
        st.warning("Não há escolas com dados de proficiência e respostas nesta etapa.")
        return
    confiavel = tabela["n"] >= 10
    figura = go.Figure()
    for filtro, nome, cor in ((confiavel, "10 ou mais respondentes", COR_DESTAQUE), (~confiavel, "menos de 10 respondentes", COR_BASE)):
        parte = tabela[filtro]
        if parte.empty:
            continue
        figura.add_scatter(x=parte["uso"], y=parte["proficiencia"], mode="markers+text", name=nome,
                           marker=dict(size=parte["n"].clip(lower=4).pow(.5) * 5 + 8, color=cor, line=dict(color="white", width=1.5)),
                           text=[f"{nome_escola(e)} (n = {n})" for e, n in zip(parte["escola"], parte["n"])], textposition="top center", textfont=dict(size=13), cliponaxis=False,
                           customdata=parte[["n"]], hovertemplate="<b>%{text}</b><br>Uso semanal: %{x:.1f}%<br>Proficiência: %{y:.1f}<extra></extra>")
    margem = (tabela["proficiencia"].max() - tabela["proficiencia"].min()) * .18 + 5
    figura.update_xaxes(title="Usam plataformas do Estado toda semana (%)", range=[-14, 110], ticksuffix="%", showgrid=True, gridcolor="#EEF2F7")
    figura.update_yaxes(title=f"Proficiência em {disciplina}", range=[tabela["proficiencia"].min() - margem, tabela["proficiencia"].max() + margem])
    figura.update_layout(title=titulo_grafico(f"{disciplina} × uso das plataformas do Estado · {etapa}", "Cada ponto é uma escola; o tamanho indica o número de respondentes"),
                         legend=dict(orientation="h", y=-.18, x=0))
    mostrar(apply_layout(figura, 520))
    validos = tabela[confiavel]
    if len(validos) >= 3:
        r = validos["uso"].corr(validos["proficiencia"])
        direcao = "positiva" if r > .3 else "negativa" if r < -.3 else "fraca ou inexistente"
        caixa("Leitura dos dados", f"Considerando as {len(validos)} escolas com 10 ou mais respondentes nesta etapa, a correlação entre o uso semanal das plataformas e a proficiência é <b>{direcao}</b> (r = {f"{r:.2f}".replace(".", ",")}). Com tão poucas escolas, o resultado é apenas exploratório e não indica causa e efeito.")
    st.caption("A pergunta 4 trata das plataformas com IA do Estado em geral (ex.: Redação Paraná, Fluência Paraná), não apenas da Redação Paraná. Escolas sem dado de proficiência para a etapa não aparecem.")
    with st.expander("Tabela de proficiência por escola"):
        st.dataframe(pd.DataFrame(PROFICIENCIA).T.rename_axis("Escola").reset_index(), width="stretch", hide_index=True)
    rodape(dataframe)


def pagina_comparacao(dataframe: pd.DataFrame) -> None:
    cabecalho("Resultados", "Escolas e etapas escolares", "Compare os principais indicadores entre etapas escolares ou entre instituições.")
    controles = st.columns([1, 2])
    with controles[0]:
        agrupamento = st.radio("Agrupar por", ["Fundamental × Médio", "Etapa escolar", "Instituição"], horizontal=True)
    with controles[1]:
        indicador = st.selectbox("Indicador", list(INDICADORES))
    campo = {"Fundamental × Médio": "etapa", "Etapa escolar": "serie"}.get(agrupamento, "escola")
    linhas = []
    for grupo, dados in dataframe.groupby(campo):
        linha = {"grupo": grupo, "n": len(dados)}
        linha.update({nome: round(calculo(dados), 1) for nome, calculo in INDICADORES.items()})
        linhas.append(linha)
    tabela = pd.DataFrame(linhas)
    if campo == "etapa":
        tabela = tabela.assign(_ordem=tabela["grupo"].map(ETAPAS.index)).sort_values("_ordem").drop(columns="_ordem")
        tabela["rotulo"] = tabela["grupo"]
    elif campo == "serie":
        tabela = tabela.assign(_ordem=tabela["grupo"].map(ordem_serie)).sort_values("_ordem").drop(columns="_ordem")
        tabela["rotulo"] = tabela["grupo"].map(rotulo_curto)
    else:
        tabela = tabela.sort_values(indicador, ascending=False)
        tabela["rotulo"] = tabela["grupo"].map(nome_escola)
    tabela = tabela.reset_index(drop=True)
    geral = INDICADORES[indicador](dataframe)
    grafico = pd.DataFrame({
        "resposta": [f"{rotulo} (n = {n})" for rotulo, n in zip(tabela["rotulo"], tabela["n"])],
        "completo": tabela["grupo"], "quantidade": tabela["n"], "percentual": tabela[indicador],
    })
    cores = [COR_BASE if n < 10 else COR_DESTAQUE for n in tabela["n"]]
    figura = grafico_barras(grafico, f"{indicador}: {pct(geral)} no geral", f"Por {agrupamento.lower()} · barras claras: grupos com menos de 10 respondentes", cores=cores)
    rotulos = [f"{quebrar(rotulo, 26)}<br><span style='font-size:13px;color:#64748B'>n = {n}</span>" for rotulo, n in zip(tabela["rotulo"], tabela["n"])]
    figura.update_traces(y=rotulos, text=[f"<b>{pct(valor)}</b>" for valor in grafico["percentual"]], hovertemplate="<b>%{customdata[0]}</b><br>%{x:.1f}% · n = %{customdata[1]}<extra></extra>")
    figura.update_xaxes(range=[0, 115])
    figura.add_vline(x=geral, line=dict(color=COLORS["orange"], dash="dash", width=2))
    figura.add_annotation(x=geral, y=1.0, yref="paper", yanchor="bottom", text=f"Geral: {pct(geral)}", showarrow=False, font=dict(color=COLORS["orange"], size=14))
    mostrar(figura)
    st.caption("A linha tracejada indica o valor considerando todos os participantes. Diferenças entre grupos pequenos podem ocorrer por acaso.")
    with st.expander("Tabela com todos os indicadores (%)"):
        st.dataframe(tabela.drop(columns="grupo").rename(columns={"rotulo": agrupamento, "n": "Respondentes"}).set_index(agrupamento).reset_index(), width="stretch", hide_index=True)
    rodape(dataframe)


def relatorio_texto(dataframe: pd.DataFrame) -> str:
    """Relatório textual com todas as perguntas, para anexar ao TCC."""
    secoes = []
    for chave, enunciado in PERGUNTAS.items():
        respostas, multipla = resumir_pergunta(dataframe, chave)
        observacao = " (múltipla escolha: percentuais somam mais de 100%)" if multipla else ""
        tabela = respostas[["completo", "quantidade", "percentual"]].rename(columns={"completo": "Resposta", "quantidade": "Qtd", "percentual": "%"}).to_string(index=False)
        secoes.append(f"{enunciado}{observacao}\n{tabela}")
    corpo = "\n\n".join(secoes)
    return f"""EduIA Analytics
Relatório consolidado | Gerado em {datetime.now():%d/%m/%Y %H:%M}

{TITULO_TCC}
{SUBTITULO_TCC}

Autores: {AUTORES.replace(' · ', ' e ')}

INDICADORES GERAIS
Participantes: {len(dataframe)}
Instituições: {dataframe['escola'].nunique()}
Usam IA 3+ vezes por semana: {pct(percentual_uso_frequente(dataframe['frequencia']))}
Sentem que a IA reduz o raciocínio (sim ou um pouco): {pct(percentual_que_comeca(dataframe['pensamento'], 'sim', 'um_pouco'))}

RESPOSTAS POR PERGUNTA

{corpo}
"""


def pdf_graficos(dataframe: pd.DataFrame) -> bytes | None:
    """PDF com um gráfico por pergunta (requer o pacote kaleido e um navegador Chrome/Edge)."""
    try:
        from plotly.subplots import make_subplots

        chaves = list(PERGUNTAS)
        figura = make_subplots(rows=len(chaves), cols=1, subplot_titles=[quebrar(PERGUNTAS[chave], 95) for chave in chaves], vertical_spacing=.045)
        for linha, chave in enumerate(chaves, start=1):
            resumo, _ = resumir_pergunta(dataframe, chave)
            figura.add_bar(row=linha, col=1, x=resumo["percentual"], y=resumo["resposta"], orientation="h", marker_color=cores_destaque(resumo),
                           text=resumo["percentual"].map(pct), textposition="outside", cliponaxis=False, showlegend=False)
            figura.update_yaxes(autorange="reversed", row=linha, col=1)
            figura.update_xaxes(range=[0, 115], visible=False, row=linha, col=1)
        figura.update_annotations(font=dict(size=13, family=FONT_FAMILY))
        figura.update_layout(title=f"EduIA Analytics – Respostas do questionário (n = {len(dataframe)})", width=1000, height=3400,
                             font=dict(family=FONT_FAMILY, size=13), margin=dict(l=230, r=40, t=90, b=40), plot_bgcolor="white")
        return figura.to_image(format="pdf")
    except Exception:
        return None


def pagina_exportacao(dataframe: pd.DataFrame) -> None:
    cabecalho("Material de apoio", "Exportação", "Baixe os resultados para anexar ao TCC ou usar nos slides.")
    colunas = st.columns(3)
    with colunas[0]:
        st.download_button("Relatório (TXT)", relatorio_texto(dataframe), file_name="relatorio_eduia_analytics.txt", mime="text/plain", width="stretch")
    with colunas[1]:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as escritor:
            dataframe.drop(columns=list(CAMPOS_PERGUNTA)).rename(columns=PERGUNTAS).to_excel(escritor, index=False, sheet_name="Respostas")
            for chave in PERGUNTAS:
                resumir_pergunta(dataframe, chave)[0][["completo", "quantidade", "percentual"]].to_excel(escritor, index=False, sheet_name=f"Pergunta {chave[-1]}")
        st.download_button("Planilha com resumos (Excel)", buffer.getvalue(), file_name="eduia_resultados.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    with colunas[2]:
        if st.button("Gerar PDF dos gráficos", width="stretch"):
            with st.spinner("Gerando PDF..."):
                st.session_state["pdf_graficos"] = pdf_graficos(dataframe)
            if st.session_state["pdf_graficos"] is None:
                st.warning("Não foi possível gerar o PDF. Verifique se o pacote kaleido está instalado (`pip install kaleido`).")
        if st.session_state.get("pdf_graficos"):
            st.download_button("Baixar PDF", st.session_state["pdf_graficos"], file_name="eduia_graficos.pdf", mime="application/pdf", width="stretch")
    st.caption("Cada gráfico também pode ser salvo como imagem: passe o mouse sobre ele e use o ícone de câmera.")


# -----------------------------------------------------------------------------
# Navegação
# -----------------------------------------------------------------------------
PAGINAS = {
    "Início": ("house", pagina_inicio),
    "Perfil da amostra": ("people", pagina_perfil),
    "Perguntas e Respostas": ("chat-square-text", pagina_perguntas),
    "Uso da IA": ("activity", pagina_uso),
    "Finalidade do uso": ("bullseye", pagina_finalidade),
    "IA e Professores": ("person-video3", pagina_professores),
    "Pensamento próprio": ("lightbulb", pagina_pensamento),
    "Cruzamentos": ("diagram-3", pagina_cruzamentos),
    "Plataformas e proficiência": ("graph-up", pagina_proficiencia),
    "Escolas e etapas": ("building", pagina_comparacao),
    "Exportação": ("download", pagina_exportacao),
}


def main() -> None:
    with st.sidebar:
        st.markdown('<div class="brand"><span class="brand-mark">IA</span><span class="brand-title">EduIA Analytics</span></div><p class="brand-sub">Pesquisa escolar · Curitiba</p>', unsafe_allow_html=True)
        pagina = option_menu(
            menu_title=None,
            options=list(PAGINAS),
            icons=[icone for icone, _ in PAGINAS.values()],
            default_index=0,
            styles={
                "container": {"padding": "0!important", "background-color": "#0F172A"},
                "icon": {"color": "#93C5FD", "font-size": "18px"},
                "nav-link": {"color": "#E2E8F0", "font-size": "15px", "font-weight": "500", "text-align": "left", "margin": "2px 0", "padding": "9px 12px", "border-radius": "8px", "--hover-color": "#1E293B"},
                "nav-link-selected": {"background-color": "#1D4ED8", "color": "#FFFFFF", "font-weight": "700"},
            },
        )
        st.divider()
        modo_apresentacao = st.toggle("Modo apresentação", value=False, help="Esconde notas e rodapés e aumenta o espaço dos gráficos.")
        st.caption(f"TCC · {CURSO}")
    if modo_apresentacao:
        st.markdown("<style>.page-lead, .rodape, [data-testid='stCaptionContainer'] { display: none !important; } .block-container { padding-top: 3rem; max-width: 100%; }</style>", unsafe_allow_html=True)
    dataframe = carregar_dados()
    PAGINAS[pagina][1](dataframe)


if __name__ == "__main__":
    main()
