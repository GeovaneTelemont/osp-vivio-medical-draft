# 🧠 Automação OSP Vivo (Playwright + Python)

[![Python](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/)
[![Playwright](https://img.shields.io/badge/playwright-1.55+-green.svg)](https://playwright.dev/docs/intro)
[![Pandas](https://img.shields.io/badge/pandas-2.3.3+-red.svg)](https://pandas.pydata.org/)
[![Poetry](https://img.shields.io/badge/Poetry-1.8+-purple.svg)](https://python-poetry.org/)

Este projeto automatiza a coleta de dados do sistema **OSP Control (Vivo)** utilizando **Python + Playwright**.  
A ferramenta faz login (com suporte a CAPTCHA manual), acessa abas específicas de **Draft** ou **Medição**, extrai tabelas e salva os resultados em arquivos **CSV**.

---

## 🚀 Funcionalidades

- 🔐 **Login automático assistido** (preenche usuário/senha e aguarda CAPTCHA manual).  
- 💾 **Reaproveitamento de sessão (`auth.json`)** — evita logins repetidos.  
- 📂 **Leitura dinâmica de IDs via CSV** (`lista_csv/lista.csv`).  
- 🌐 **Navegação automática** até as abas **Draft** ou **Medição**.  
- 📊 **Extração de tabelas estruturadas** (Serviços, Custos e Materiais).  
- 🧾 **Geração automática de relatórios CSV** (`osp_vivo_draft.csv` ou `osp_vivo_medicao.csv`).  

---

## 🧩 Estrutura do Projeto

```
📦 projeto_osp_vivo/
│
├── main.py                 # Script principal
├── auth.json               # Sessão de login (gerada automaticamente)
├── osp_vivo_draft.csv      # Dados extraídos do modo Draft
├── osp_vivo_medicao.csv    # Dados extraídos do modo Medição
├── lista_csv/
│   └── lista.csv           # Lista de IDs OSP a processar
├── pyproject.toml          # Configuração do Poetry (dependências)
└── README.md               # Este arquivo
```

---

## 📦 Requisitos

### 🐍 Python
- Versão **3.9 ou superior**

### 📚 Dependências principais
- `playwright`
- `pandas`
- `tkinter`
- `unicodedata`
- `shutil`
- `time`
- `pathlib`

### 🧰 Instalação via Poetry (recomendado)

1. Instale o **Poetry**:
   ```bash
   pip install poetry
   ```

2. Instale as dependências do projeto:
   ```bash
   poetry install
   ```

3. Instale o navegador Chromium (usado pelo Playwright):
   ```bash
   poetry run playwright install chromium
   ```

---

## ⚙️ Configuração do Ambiente

1. **Defina suas credenciais no script**  
   Edite o início do arquivo `main.py` e insira:
   ```python
   USERNAME = "seu_usuario"
   PASSWORD = "sua_senha"
   ```

2. **Prepare o CSV de entrada**  
   Crie (ou selecione) um arquivo CSV com a seguinte estrutura e nomeie como `lista.csv`:
   ```csv
   ID;OSP MEDIDO
   123456;OSP123
   789012;OSP456
   ```
   > O script abrirá uma janela para selecionar este arquivo, e ele será copiado para a pasta `lista_csv`.

3. **Execute o script**
   ```bash
   poetry run python webscrap.py
   ```

4. Escolha a opção desejada:
   ```
   Digite (1) para Draft ou Digite (2) para Medição ou (0) para sair:
   ```

5. Durante o login:
   - O script preencherá automaticamente usuário e senha.
   - Você deve resolver o **CAPTCHA manualmente** e clicar em *Login*.
   - Assim que o login for detectado, a sessão será salva em `auth.json`.

---

## 💾 Saída Gerada

Os dados são salvos em arquivos `.csv` dependendo do modo selecionado:

| Modo | Arquivo gerado | Descrição |
|------|----------------|------------|
| Draft | `osp_vivo_draft.csv` | Dados da aba **Draft → Serviços** |
| Medição | `osp_vivo_medicao.csv` | Dados da aba **Medição → Serviços** |

### Estrutura das colunas:

| ID | TIPO DE REGISTRO | CÓDIGO | DESCRIÇÃO | QUANTIDADE | PREÇO UNITÁRIO | UNIDADE | PREÇO TOTAL | CATEGORIA |
|----|------------------|---------|------------|-------------|----------------|----------|--------------|------------|

> Cada execução adiciona novas linhas sem sobrescrever os dados existentes.

---

## 🧠 Principais Funções

| Função | Descrição |
|--------|------------|
| `read_csv_id()` | Lê o arquivo `lista.csv` com os IDs a processar |
| `login(page)` | Realiza o login automático e aguarda CAPTCHA manual |
| `is_logged(page)` | Verifica se há sessão ativa (usuário logado) |
| `pesquisar_id_draft(page, id)` | Extrai dados da aba **Draft** |
| `pesquisar_id_medicao(page, id)` | Extrai dados da aba **Medição** |
| `webscraping(page, df, func, name_file)` | Controla a extração e gravação dos dados |
| `main(number)` | Gerencia o fluxo principal (login, extração, salvamento) |
| `open_dialog_csv()` | Abre seletor de arquivo para importar a lista CSV |
| `run()` | Ponto de entrada do script |

---

## ⚠️ Observações Importantes

- O **login pode exigir CAPTCHA manual** — o script aguardará até que o usuário conclua.  
- A sessão salva (`auth.json`) evita novos logins, desde que válida.  
- Se houver erro de timeout ou bloqueio, feche e reabra o navegador.  
- É necessário acesso à **rede interna da Vivo** para usar o sistema OSP Control.  
- Use intervalos (`sleep`) com cautela para não sobrecarregar o servidor.  

---

## 🧾 Licença

Este projeto é de uso **exclusivo da empresa Telemont**.  
Reprodução, redistribuição ou uso fora do ambiente corporativo não são autorizados.

---

## 👨‍💻 Autor

**GeovaneTelemont**  
Automação de processos corporativos — Python + Playwright + Pandas + Poetry  
📧 *Desenvolvido para otimizar rotinas OSP Vivo.*

---

> 🧩 *“Automatizar é transformar tempo repetido em resultado produtivo.”*
