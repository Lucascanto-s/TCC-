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

O aplicativo lê o arquivo `Total_master.xlsx` (respostas do Google Forms) na mesma pasta de `app.py`. Se existir uma aba chamada `Principal`, apenas ela é usada; caso contrário, as abas não vazias são combinadas.

As perguntas são identificadas pelo número no início do cabeçalho (`1.`, `2.`, ...). Nomes de escolas que diferem só em maiúsculas ou acentos são unificados.

Sem a planilha, o app mostra um aviso e não exibe nenhum resultado (não há dados fictícios).

## Personalização

No início de `app.py` há campos para editar o título do TCC, os autores, o orientador, a instituição e o período de coleta. Também há o dicionário `INTERPRETACOES`, onde os autores podem escrever a interpretação de cada pergunta; o texto aparece na aba Perguntas e Respostas.

## Páginas

- **Início**: capa do trabalho, indicadores gerais e principais achados.
- **Perfil da amostra**: participantes por etapa escolar e por instituição.
- **Perguntas e Respostas**: análise individual de cada pergunta, com filtro por instituição.
- **Uso da IA**: frequência de uso, ferramentas e plataformas do Estado.
- **Finalidade do uso**: para que os estudantes usam a IA.
- **IA e Professores**: reação dos professores e comparação IA x professor.
- **Pensamento próprio**: percepção sobre o raciocínio e cruzamento com a intensidade de uso.
- **Escolas e etapas**: comparação de indicadores entre grupos.
- **Exportação**: relatório TXT, planilha Excel com resumos e PDF dos gráficos.

O botão **Modo apresentação** na barra lateral esconde notas e rodapés.

## Estrutura principal

```text
.
├── app.py             # Aplicação Streamlit e telas do dashboard
├── theme.py           # Cores, fontes e configuração dos gráficos
├── requirements.txt   # Dependências Python
├── Total_master.xlsx  # Respostas do questionário
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

Coloque `Total_master.xlsx` na mesma pasta de `app.py` e recarregue a página.

### Como parar o aplicativo

No terminal onde o Streamlit está rodando, pressione `Ctrl+C`.

## Observação

O dashboard foi desenvolvido como parte de um Trabalho de Conclusão de Curso sobre os impactos do uso de Inteligência Artificial no desempenho escolar.