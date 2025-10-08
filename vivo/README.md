# 🧠 Automação OSP Vivo (Playwright + Python)

Este projeto automatiza a coleta de dados do sistema **OSP Control** da Vivo, utilizando **Playwright** em modo síncrono com **Python**.  
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
└── requirements.txt        # Dependências do projeto
```

---

## 📦 Requisitos

### Python
- Versão **3.9+**

### Bibliotecas
Instale as dependências com:

```bash
pip install playwright pandas
```

E inicialize o navegador Chromium (necessário apenas uma vez):

```bash
playwright install chromium
```

---

## ⚙️ Configuração

1. **Defina as credenciais**  
   No início do arquivo `main.py`, preencha as variáveis:
   ```python
   USERNAME = "seu_usuario"
   PASSWORD = "sua_senha"
   ```

   > 💡 Por segurança, o script solicita que o usuário complete o CAPTCHA manualmente durante o primeiro login.

2. **Prepare o arquivo de entrada**  
   No diretório `lista_csv`, crie um arquivo chamado `lista.csv` com o seguinte formato:
   ```csv
   ID;OSP MEDIDO;url
   12345;PB Classe A;https://devopsredes.vivo.com.br/ospcontrol/requisicoes-eps?id=12345
   67890;PB Classe B;https://devopsredes.vivo.com.br/ospcontrol/requisicoes-eps?id=67890
   ```

3. **Execute o script**
   ```bash
   python main.py
   ```

4. **Após o login manual**, o script:
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

Este projeto é de uso interno e educativo, sem fins comerciais.

---

## 👨‍💻 Autor

**GeovaneTelemont**  
Automação de processos com Python + Playwright
