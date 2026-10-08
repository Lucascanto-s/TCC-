# EduIA Analytics

Dashboard em Streamlit para analisar os impactos do uso de Inteligência Artificial no desempenho escolar. O aplicativo apresenta indicadores, gráficos, filtros por instituição e análises das respostas do questionário.

## Requisitos

- Python 3.10 ou superior
- pip
- Navegador web

As dependências do projeto estão listadas em [requirements.txt](requirements.txt).

## Instalação

Na pasta do projeto, crie e ative um ambiente virtual:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Depois, instale as dependências:

```bash
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

No Windows, ative o ambiente virtual com:

```powershell
.venv\Scripts\activate
```

## Como executar

Execute a partir da raiz do projeto, onde estão `app.py` e `theme.py`:

```bash
python3 -m streamlit run app.py
```

O Streamlit exibirá um endereço semelhante a `http://localhost:8501`. Abra esse endereço no navegador.

Também é possível usar o comando tradicional quando o executável estiver no `PATH`:

```bash
streamlit run app.py
```

### Erro `streamlit: command not found`

Esse erro significa que o pacote pode estar instalado, mas a pasta dos executáveis do Python não está no `PATH`. A forma mais confiável de executar é:

```bash
python3 -m streamlit run app.py
```

No Linux ou macOS, para usar diretamente o comando `streamlit` na sessão atual, localize o diretório de scripts do usuário e adicione-o ao `PATH`:

```bash
python3 -m site --user-base
export PATH="$(python3 -m site --user-base)/bin:$PATH"
streamlit run app.py
```

Se o Streamlit foi instalado dentro de um ambiente virtual, ative o ambiente antes de executar o app.

## Dados

O aplicativo procura automaticamente o arquivo `Total_master.xlsx` na mesma pasta de `app.py`. Quando esse arquivo não existe, o dashboard usa uma base demonstrativa para permitir a execução e a apresentação visual.

O arquivo Excel pode conter várias abas. As abas não vazias são combinadas automaticamente. O sistema reconhece cabeçalhos em português ou inglês e também tenta identificar perguntas completas do formulário.

## Funcionalidades

- Visão geral dos participantes e das instituições.
- Distribuição por série, frequência e finalidade de uso da IA.
- Análise das plataformas Redação Paraná e Fluência Paraná.
- Comparação entre IA e mediação do professor.
- Indicadores de autonomia e pensamento próprio.
- Consulta individual das perguntas e respostas do questionário.
- Tendências por série ou instituição.
- Comparação entre escolas.
- Exportação de relatório textual.

## Estrutura principal

```text
.
├── app.py             # Aplicação Streamlit e telas do dashboard
├── theme.py           # Cores, fontes e configuração dos gráficos
├── requirements.txt   # Dependências Python
├── Total_master.xlsx  # Base principal opcional
└── README.md          # Instruções do projeto
```

## Solução de problemas

### O navegador não abriu automaticamente

Copie o endereço mostrado no terminal, normalmente `http://localhost:8501`, e abra-o no navegador.

### A porta 8501 já está ocupada

Inicie o app em outra porta:

```bash
python3 -m streamlit run app.py --server.port 8502
```

### O arquivo Excel não foi encontrado

Coloque `Total_master.xlsx` na mesma pasta de `app.py`. Sem esse arquivo, o app continua funcionando com dados demonstrativos.

### Como parar o aplicativo

No terminal onde o Streamlit está rodando, pressione `Ctrl+C`.

## Observação

O dashboard foi desenvolvido como parte de um Trabalho de Conclusão de Curso sobre os impactos do uso de Inteligência Artificial no desempenho escolar.