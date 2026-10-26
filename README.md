# Plataforma Analítica Executiva

Dashboard interativo feito com **Streamlit** e **Plotly** a partir da planilha `financeiro.xlsx`.
Tem 4 abas: Cockpit de KPIs, Matriz de Portfólio, Simulador de Cenários e Performance Regional com exportação.

## O que você precisa ter instalado

1. **Python 3.10 ou mais novo** — baixe em <https://www.python.org/downloads/>.
   - No Windows, marque a opção **"Add Python to PATH"** durante a instalação.
2. **Git** — baixe em <https://git-scm.com/downloads>.

Para conferir se deu certo, abra o terminal (Prompt de Comando, PowerShell ou Terminal) e rode:

```bash
python --version
git --version
```

> Em alguns sistemas o comando é `python3` em vez de `python`.

## Passo a passo

### 1. Baixar (importar) o projeto

```bash
git clone https://github.com/moraesalfredo13-spec/pythonlit-aulaBC.git
cd pythonlit-aulaBC
```

Sem o Git, também dá para clicar em **Code → Download ZIP** na página do repositório no GitHub,
extrair a pasta e abrir o terminal dentro dela.

### 2. Criar um ambiente virtual

O ambiente virtual isola as bibliotecas do projeto, para não bagunçar o resto do computador.

```bash
python -m venv .venv
```

### 3. Ativar o ambiente virtual

| Sistema | Comando |
| --- | --- |
| Windows (PowerShell) | `.venv\Scripts\Activate.ps1` |
| Windows (Prompt de Comando) | `.venv\Scripts\activate.bat` |
| Linux / macOS / WSL | `source .venv/bin/activate` |

Quando funcionar, aparece `(.venv)` no começo da linha do terminal.

> No PowerShell, se aparecer erro de permissão, rode uma vez:
> `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`

### 4. Instalar as dependências

```bash
pip install -r requirements.txt
```

### 5. Executar o dashboard

```bash
streamlit run app.py
```

O navegador abre sozinho. Se não abrir, acesse **http://localhost:8501**.

Para parar o dashboard, volte ao terminal e pressione `Ctrl + C`.

## Como usar

- **Barra lateral:** filtre por data, ano, país, segmento e produto. Os botões **Todos** e **Limpar** marcam ou desmarcam tudo.
- **Abas:** navegue pelas 4 abas no topo da página.
- **Simulador:** mexa nos controles deslizantes para ver o lucro projetado.
- **Exportação:** na última aba, baixe os dados filtrados em CSV ou Excel.

## Problemas comuns

| Problema | Solução |
| --- | --- |
| `python não é reconhecido` | Reinstale o Python marcando "Add Python to PATH" ou use `python3`. |
| `streamlit não é reconhecido` | O ambiente virtual não está ativo. Refaça o passo 3. |
| `FileNotFoundError: financeiro.xlsx` | Rode o comando de dentro da pasta do projeto, onde está o `app.py`. |
| Porta 8501 ocupada | Use outra porta: `streamlit run app.py --server.port 8502` |

## Estrutura do projeto

```
├── app.py               # código do dashboard
├── financeiro.xlsx      # base de dados
├── requirements.txt     # bibliotecas necessárias
└── .streamlit/
    └── config.toml      # tema (cores) do dashboard
```
