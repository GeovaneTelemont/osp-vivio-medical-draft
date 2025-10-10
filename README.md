# 🧠 Automação OSP Vivo (Playwright + Python)

[![Python](https://img.shields.io/badge/python-3.13+-blue.svg)](https://www.python.org/)
[![Playwright](https://img.shields.io/badge/playwright-1.55+-green.svg)](https://playwright.dev/docs/intro)
[![Pandas](https://img.shields.io/badge/pandas-2.3.3+-red.svg)](https://pandas.pydata.org/docs/)
[![Poetry](https://img.shields.io/badge/Poetry-1.8+-purple.svg)](https://python-poetry.org/)

Este projeto automatiza a coleta de dados do sistema **OSP Control** da Vivo, utilizando **Playwright** em modo síncrono com **Python** e gerenciamento de dependências via **Poetry**.  
O script realiza login (com suporte a CAPTCHA manual), acessa páginas específicas, extrai tabelas de medições e salva os resultados em um arquivo **CSV**.

---

## 🚀 Funcionalidades

- Login automatizado com suporte para login manual em caso de CAPTCHA.  
- Reaproveitamento de sessão autenticada (`auth.json`) para evitar logins repetidos.  
- Leitura de uma lista de IDs a partir de um arquivo CSV (`lista_csv/lista.csv`).  
- Navegação automática até a aba **Medição → Serviços** de cada requisição.  
- Extração estruturada de tabelas (colunas e valores).  
- Salvamento dos dados coletados no arquivo `osp_vivo.csv`.  

---

## 🧩 Estrutura do Projeto

```
📂 projeto/
│
├── main.py                 # Script principal
├── auth.json               # Sessão de login (gerada automaticamente)
├── osp_vivo.csv            # Saída dos dados coletados
├── lista_csv/
│   └── lista.csv           # Lista de IDs ou URLs a processar
├── pyproject.toml          # Configuração do Poetry
└── README.md               # Este arquivo

```

---

## 📦 Requisitos

### Python
- Versão **3.9+**

### Poetry
Instale o **Poetry** (se ainda não tiver):
```bash
pip install poetry
```

---

## ⚙️ Configuração do Ambiente

1. **Instalar dependências**  
   Na raiz do projeto, execute:
   ```bash
   poetry install
   ```

2. **Inicializar o Playwright**  
   Após instalar, baixe o navegador Chromium:
   ```bash
   poetry run playwright install chromium
   ```

3. **Definir credenciais**  
   No início do arquivo `main.py`, preencha as variáveis:
   ```python
   USERNAME = "seu_usuario"
   PASSWORD = "sua_senha"
   ```

   > 💡 O script solicitará que você complete o CAPTCHA manualmente no primeiro login.

4. **Executar o script**
   ```bash
   poetry run python main.py
   ```

5. **Após o login manual**, o script:
   - Detecta a autenticação,
   - Salva a sessão em `auth.json`,
   - E continua automaticamente o processamento dos IDs da lista.

---

## 🗂️ Saída Gerada

Os dados extraídos são salvos em `osp_vivo.csv`, com as seguintes colunas:

| ID | CÓDIGO | DESCRIÇÃO DOS SERVIÇOS | QUANTIDADE | PREÇO UNITÁRIO | UNIDADE | PREÇO TOTAL |
|----|---------|------------------------|-------------|----------------|----------|--------------|

Cada execução adiciona novas linhas ao arquivo sem sobrescrever os dados anteriores.

---

## 🧰 Principais Funções

| Função | Descrição |
|--------|------------|
| `read_csv_id()` | Lê o arquivo `lista.csv` com IDs/URLs a processar |
| `login(page)` | Realiza o login manual assistido (usuário e senha preenchidos automaticamente) |
| `is_logged(page)` | Verifica se há sessão ativa |
| `pesquisar_id(page, id)` | Navega até o ID informado e coleta os dados da tabela de serviços |
| `main()` | Controla o fluxo principal da automação |

---

## ⚠️ Observações Importantes

- O login pode exigir **CAPTCHA manual** — o script pausa até o usuário completar.
- Se já houver um `auth.json` válido, o login manual será pulado automaticamente.
- Utilize **intervalos (sleep)** e **timeouts** apropriados para evitar bloqueios do sistema.
- É necessário acesso à rede interna da Vivo para o sistema `OSP Control`.

---

## 🧾 Licença

Este projeto é de uso exclusivo para empresa Telemont.

---

## 👨‍💻 Autor

**GeovaneTelemont**  
Automação de processos com Python + Playwright + Poetry
