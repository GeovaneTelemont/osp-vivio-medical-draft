import subprocess
import sys

# --- Verifica se Playwright está instalado, se não instala automaticamente ---
try:
    from playwright.sync_api import sync_playwright, TimeoutError
except ImportError:
    print("📦 Playwright não encontrado. Instalando automaticamente...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "playwright"])
    print("✅ Playwright instalado. Instalando navegadores...")
    subprocess.check_call([sys.executable, "-m", "playwright", "install"])
    print("✅ Navegadores Playwright instalados.")
    from playwright.sync_api import sync_playwright, TimeoutError



from playwright.sync_api import sync_playwright, TimeoutError
import os, unicodedata, shutil
import pandas as pd
from time import sleep
from tkinter import filedialog, Tk
from pathlib import Path

AUTH_FILE = "auth.json"
LOGIN_URL = "https://devopsredes.vivo.com.br/ospcontrol/home"
LOGGED_SELECTOR = 'xpath=//*[@id="ott-username"]' 
USERNAME = ""  # 🔒 Preencha ou use input()
PASSWORD = ""  # 🔒 Preencha ou use input()
DOWNLOAD_PATH = Path.home() / "Downloads"


# ===========================================================
# 📁 FUNÇÕES AUXILIARES
# ===========================================================

def read_csv_id():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_path = os.path.join(script_dir, "./lista_csv/lista.csv")
    df = pd.read_csv(csv_path, sep=";", encoding="utf-8")
    return df


def _normalize_text(s: str) -> str:
    if not s:
        return ""
    s = " ".join(s.split())
    s_norm = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("ASCII").lower()
    return s_norm


def salvar_incremental(resultados, colunas, arquivo):
    """Salva incrementalmente e garante integridade do Excel"""
    try:
        df_parcial = pd.DataFrame(resultados, columns=colunas)
        df_parcial.to_excel(arquivo, index=False)
        print(f"💾 Progresso salvo ({len(resultados)} registros)")
    except Exception as e:
        print(f"⚠️ Erro ao salvar arquivo: {e}")


# ===========================================================
# 🔐 LOGIN E SESSÃO
# ===========================================================

def login(page):
    """Preenche usuário e senha, espera login manual (CAPTCHA)."""
    print("Efetuando login...")
    try:
        page.wait_for_selector('//*[@id="username"]', state="visible", timeout=10000)
        page.fill('//*[@id="username"]', USERNAME)
        page.fill('//*[@id="password"]', PASSWORD)
        print("✅ Usuário e senha preenchidos. Complete o CAPTCHA e clique em login manualmente...")

        page.wait_for_selector(LOGGED_SELECTOR, timeout=60000, state="visible")
        print("✅ Login detectado! Continuando automação...")

    except TimeoutError:
        print("❌ Campos de login não encontrados — talvez já esteja logado.")
    except Exception as e:
        print("Erro ao tentar logar:", e)


def is_logged(page):
    """Retorna True se o seletor que indica login estiver presente."""
    try:
        page.wait_for_selector(LOGGED_SELECTOR, timeout=5000)
        return True
    except:
        return False

"""
    Funções para fazer pesquisa na pagina
"""


def pesquisar_id(page, id):
    """Pesquisa um ID no OSP Control, abre Medição e retorna [ID, Contrato, OSP]."""
    print(f"🔍 Pesquisando ID: {id}")
    try:
        # Converte o ID para inteiro (remove .0)
        id_int = int(float(id))
        
        sleep(2)
        page.click('//*[@id="ott-sidebar-collapse"]', timeout=10000)
        sleep(2)
        page.click('//*[@id="ott-sidebar"]/div[3]/ul/li[3]/a', timeout=10000)
        sleep(2)
        page.fill('xpath=//*[@id="filtroId"]', str(id_int))
        sleep(2)
        page.locator("a.btn.btn-primary.btn-sm.btn-block:has-text('Buscar')").click()
        sleep(2)
        page.locator("span.badge.bg-primary:has-text('Editar')").click()
        sleep(2)

        # Navega até aba "Medição"
        links = page.locator("a.nav-link")
        total = int(links.count())
        for i in range(total):
            texto = links.nth(i).text_content().strip()
            if texto == "Medição":
                links.nth(i).click()
                break

        sleep(3)
        
        # Verifica botão Serviços
        servicos_btn = page.locator('a[title="Serviços"]')
        if servicos_btn.count() == 0:
            print(f"❌ Botão 'Serviços' não encontrado para ID {id_int}")
            return [id_int, "ERRO", "Botão serviço não encontrado"]
        
        servicos_btn.click()
        sleep(3)
        
        # Extrai Contrato
        contrato = page.locator(
            'xpath=/html/body/app-root/app-requisicoes-servicos/div/div/div/div/div[2]/div[2]/div/div/div[2]/div[2]/span'
        ).text_content().strip()
        sleep(2)
        
        # Extrai OSP (com verificação)
        osp_locator = page.locator(
            'xpath=/html/body/app-root/app-requisicoes-servicos/div/div/div/div/div[2]/div[3]/div/div[2]/div/strong'
        )
        osp = osp_locator.text_content().strip() if osp_locator.count() > 0 else ""
        
        print(f"✅ ID {id_int}: Contrato='{contrato}', OSP='{osp} OSP'")
        return [id_int, contrato, osp + " OSP"]

    except Exception as e:
        print(f"❌ Erro geral ao pesquisar ID {id}: {e}")
        id_int = int(float(id)) if id else 0
        return [id_int, "ERRO", f"Erro geral: {str(e)}"]

def pesquisar_id_draft(page, id):
    try:
        sleep(2)
        page.click('//*[@id="ott-sidebar-collapse"]', timeout=10000)
        sleep(2)
        page.click('//*[@id="ott-sidebar"]/div[3]/ul/li[3]/a', timeout=10000)
        sleep(2)
        page.fill('xpath=//*[@id="filtroId"]', str(id))
        sleep(2)
        page.locator("a.btn.btn-primary.btn-sm.btn-block:has-text('Buscar')").click()
        sleep(2)
        page.locator("span.badge.bg-primary:has-text('Editar')").click()
        sleep(2)
        
        links = page.locator("a.nav-link")
        total = int(links.count())
        for i in range(total):
            texto = links.nth(i).text_content().strip()
            if texto == "Draft":
                links.nth(i).click()
                break

        sleep(3)

       

        
        page.locator('a[title="Serviços"]').click()
        sleep(2)
        # --- Extrai todas as tabelas ---
        tabelas = page.query_selector_all("//table")
        todos_dados = []

        for idx, tabela in enumerate(tabelas, start=1):
            # procura o texto útil anterior à tabela (subindo irmãos e ancestrais)
            categoria = tabela.evaluate("""
                el => {
                    function findPrevText(e){
                        // procura irmãos anteriores com texto
                        let node = e.previousElementSibling;
                        while(node){
                            const txt = node.innerText ? node.innerText.trim() : '';
                            if(txt) return txt.replace(/\\s+/g,' ');
                            node = node.previousElementSibling;
                        }
                        // sobe na arvore e procura irmãos anteriores dos ancestrais
                        let parent = e.parentElement;
                        while(parent){
                            let prev = parent.previousElementSibling;
                            while(prev){
                                const txt = prev.innerText ? prev.innerText.trim() : '';
                                if(txt) return txt.replace(/\\s+/g,' ');
                                prev = prev.previousElementSibling;
                            }
                            parent = parent.parentElement;
                        }
                        return '';
                    }
                    return findPrevText(el);
                }
            """)
            if not isinstance(categoria, str):
                categoria = "" if categoria is None else str(categoria)
            categoria = " ".join(categoria.split())

            # pega linhas do corpo
            linhas = tabela.query_selector_all("tbody tr")
            for linha in linhas:
                tds = linha.query_selector_all("td")
                valores = [td.inner_text().strip() for td in tds]

                # ignora linhas incompletas (ex: Total Geral que tem colspan)
                if len(valores) < 6:
                    continue

                # detectar tipo de registro com base na categoria (mais robusto)
                cat_norm = _normalize_text(categoria)
                if any(x in cat_norm for x in ("material", "materiais", "telefonica", "telefonica")):
                    tipo_registro = "Material"
                elif any(x in cat_norm for x in ("custo", "custos")):
                    tipo_registro = "Custo"
                elif any(x in cat_norm for x in ("servico", "servicos", "classe", "valor")):
                    tipo_registro = "Serviço"
                else:
                    # fallback: olhar unidade/descrição para adivinhar
                    unidade = valores[4].lower() if len(valores) > 4 and valores[4] else ""
                    descricao = valores[1].lower() if len(valores) > 1 and valores[1] else ""
                    if any(k in unidade for k in ("m", "u", "cj")) or any(k in descricao for k in ("cfo", "chassi", "conj", "subduto", "material")):
                        tipo_registro = "Material"
                    else:
                        tipo_registro = "Serviço"

                # monta a linha na ordem desejada
                dados_linha = [
                    id,                 # ID
                    tipo_registro,      # TIPO DE REGISTRO
                    valores[0],         # CODIGO
                    valores[1],         # DESCRIÇÃO DOS SERVIÇOS
                    valores[2],         # QUANTIDADE
                    valores[3],         # PREÇO UNITÁRIO
                    valores[4],         # UNIDADE
                    valores[5],         # PREÇO TOTAL
                    categoria           # CATEGORIA (texto anterior à tabela)
                ]

                todos_dados.append(dados_linha)

        # volta no menu (se for preciso)
        sleep(2)
        page.click('//*[@id="ott-sidebar-collapse"]')
        sleep(2)
        page.click('//*[@id="ott-sidebar"]/div[3]/ul/li[1]/a')
        sleep(2)
        print(f"✅ ID: {id} - extração finalizada. Linhas: {len(todos_dados)}")


        print(todos_dados)
        return todos_dados

    except Exception as e:
        print(f"Erro ao pesquisar ID {id}: {e}")
        return None




def pesquisar_id_medicao(page, id):
    try:
        sleep(2)
        page.click('//*[@id="ott-sidebar-collapse"]', timeout=10000)
        sleep(2)
        page.click('//*[@id="ott-sidebar"]/div[3]/ul/li[3]/a', timeout=10000)
        sleep(2)
        page.fill('xpath=//*[@id="filtroId"]', str(id))
        sleep(2)
        page.locator("a.btn.btn-primary.btn-sm.btn-block:has-text('Buscar')").click()
        sleep(2)
        page.locator("span.badge.bg-primary:has-text('Editar')").click()
        sleep(2)

        links = page.locator("a.nav-link")
        total = int(links.count())
        for i in range(total):
            texto = links.nth(i).text_content().strip()
            if texto == "Medição":
                links.nth(i).click()
                break

        sleep(3)
        page.locator('a[title="Serviços"]').click()
        sleep(2)

        # --- Extrai todas as tabelas ---
        tabelas = page.query_selector_all("//table")
        todos_dados = []

        for idx, tabela in enumerate(tabelas, start=1):
            # procura o texto útil anterior à tabela (subindo irmãos e ancestrais)
            categoria = tabela.evaluate("""
                el => {
                    function findPrevText(e){
                        // procura irmãos anteriores com texto
                        let node = e.previousElementSibling;
                        while(node){
                            const txt = node.innerText ? node.innerText.trim() : '';
                            if(txt) return txt.replace(/\\s+/g,' ');
                            node = node.previousElementSibling;
                        }
                        // sobe na arvore e procura irmãos anteriores dos ancestrais
                        let parent = e.parentElement;
                        while(parent){
                            let prev = parent.previousElementSibling;
                            while(prev){
                                const txt = prev.innerText ? prev.innerText.trim() : '';
                                if(txt) return txt.replace(/\\s+/g,' ');
                                prev = prev.previousElementSibling;
                            }
                            parent = parent.parentElement;
                        }
                        return '';
                    }
                    return findPrevText(el);
                }
            """)
            if not isinstance(categoria, str):
                categoria = "" if categoria is None else str(categoria)
            categoria = " ".join(categoria.split())

            # pega linhas do corpo
            linhas = tabela.query_selector_all("tbody tr")
            for linha in linhas:
                tds = linha.query_selector_all("td")
                valores = [td.inner_text().strip() for td in tds]

                # ignora linhas incompletas (ex: Total Geral que tem colspan)
                if len(valores) < 6:
                    continue

                # detectar tipo de registro com base na categoria (mais robusto)
                cat_norm = _normalize_text(categoria)
                if any(x in cat_norm for x in ("material", "materiais", "telefonica", "telefonica")):
                    tipo_registro = "Material"
                elif any(x in cat_norm for x in ("custo", "custos")):
                    tipo_registro = "Custo"
                elif any(x in cat_norm for x in ("servico", "servicos", "classe", "valor")):
                    tipo_registro = "Serviço"
                else:
                    # fallback: olhar unidade/descrição para adivinhar
                    unidade = valores[4].lower() if len(valores) > 4 and valores[4] else ""
                    descricao = valores[1].lower() if len(valores) > 1 and valores[1] else ""
                    if any(k in unidade for k in ("m", "u", "cj")) or any(k in descricao for k in ("cfo", "chassi", "conj", "subduto", "material")):
                        tipo_registro = "Material"
                    else:
                        tipo_registro = "Serviço"

                # monta a linha na ordem desejada
                dados_linha = [
                    id,                 # ID
                    tipo_registro,      # TIPO DE REGISTRO
                    valores[0],         # CODIGO
                    valores[1],         # DESCRIÇÃO DOS SERVIÇOS
                    valores[2],         # QUANTIDADE
                    valores[3],         # PREÇO UNITÁRIO
                    valores[4],         # UNIDADE
                    valores[5],         # PREÇO TOTAL
                    categoria           # CATEGORIA (texto anterior à tabela)
                ]

                todos_dados.append(dados_linha)

        # volta no menu (se for preciso)
        sleep(2)
        page.click('//*[@id="ott-sidebar-collapse"]')
        sleep(2)
        page.click('//*[@id="ott-sidebar"]/div[3]/ul/li[1]/a')
        sleep(2)
        print(f"✅ ID: {id} - extração finalizada. Linhas: {len(todos_dados)}")

        return todos_dados

    except Exception as e:
        print(f"Erro ao pesquisar ID {id}: {e}")
        return None

# ===========================================================
# 🔍 FUNÇÕES DE PESQUISA (mantidas iguais)
# ===========================================================

# pesquisar_id, pesquisar_id_draft, pesquisar_id_medicao
# (mantidas conforme o seu código original)
# ... [essas não mudam nada além de chamadas internas]


# ===========================================================
# 🧭 FUNÇÕES DE WEBSCRAPING (melhoradas)
# ===========================================================

def webscraping_id(page, df, func_pesquisa, name_file_xlsx):
    """Webscraping para ID Cancelados"""
    print("🔄 Iniciando webscraping para pesquisa de ID...")

    arquivo = DOWNLOAD_PATH / name_file_xlsx
    colunas = ["ID", "CONTRATO", "OSP"]
    resultados = []

    for index, row in df.iterrows():
        id_value = int(row["ID"])
        print(f"📋 Processando ID: {id_value}")

        try:
            dados = func_pesquisa(page, id_value)
            if dados:
                resultados.append(dados)
            else:
                resultados.append([id_value, "ERRO", "Nenhum dado retornado"])
        except Exception as e:
            resultados.append([id_value, "ERRO", str(e)])

        # 💾 salvamento incremental robusto
        salvar_incremental(resultados, colunas, arquivo)

    print(f"✅ Arquivo final salvo em: {arquivo}")

    # 🚀 abre a pasta com o arquivo selecionado
    try:
        os.startfile(arquivo)
        print("📂 Pasta Downloads aberta com o arquivo selecionado.")
    except Exception as e:
        print(f"⚠️ Não foi possível abrir automaticamente: {e}")


def webscraping(page, df, func_pesquisa, name_file_xlsx):
    """Webscraping para Draft e Medição"""
    print("🔄 Iniciando webscraping geral (Draft/Medição)...")

    arquivo = DOWNLOAD_PATH / name_file_xlsx
    colunas = [
        "ID", "TIPO DE REGISTRO", "CÓDIGO", "DESCRIÇÃO",
        "QUANTIDADE", "PREÇO UNITÁRIO", "UNIDADE", "PREÇO TOTAL", "CATEGORIA"
    ]

    resultados = []

    for index, row in df.iterrows():
        id_value = int(row["ID"])
        print(f"📋 Processando ID: {id_value}")

        try:
            dados = func_pesquisa(page, id_value)
            if dados:
                resultados.extend(dados)
                print(f"✅ ID {id_value} extraído ({len(dados)} linhas).")
            else:
                print(f"⚠️ Nenhum dado retornado para ID {id_value}")
        except Exception as e:
            print(f"❌ Erro ao processar ID {id_value}: {e}")

        # 💾 salvamento incremental
        salvar_incremental(resultados, colunas, arquivo)

    print(f"✅ Arquivo final salvo em:\n📁 {arquivo}")

    # 🚀 abre a pasta com o arquivo selecionado
    try:
        os.startfile(arquivo)
        print("📂 Pasta Downloads aberta com o arquivo selecionado.")
    except Exception as e:
        print(f"⚠️ Não foi possível abrir automaticamente: {e}")


# ===========================================================
# 🧠 MAIN E EXECUÇÃO (mantidos, com prints aprimorados)
# ===========================================================

def main(number):
    with sync_playwright() as p:
        print("Iniciando navegador...")
        browser = p.chromium.launch(channel="chrome", headless=False, args=["--ignore-certificate-errors"])

        if os.path.exists(AUTH_FILE):
            print("Carregando sessão existente de", AUTH_FILE)
            context = browser.new_context(storage_state=AUTH_FILE)
        else:
            print("Nenhuma sessão encontrada. Criando uma nova.")
            context = browser.new_context()

        page = context.new_page()
        page.goto(LOGIN_URL, wait_until="networkidle")
        print("✅ Página carregada.")

        if not is_logged(page):
            print("🔐 Sessão expirada ou inexistente. Tentando logar novamente.")
            login(page)
            page.wait_for_load_state("networkidle", timeout=30000)

            if is_logged(page):
                print("✅ Login bem-sucedido. Salvando sessão...")
                context.storage_state(path=AUTH_FILE)
            else:
                print("⚠️ Falha no login. Verifique as credenciais ou o CAPTCHA.")
        else:
            print("✅ Já está logado!")

        print("Lendo CSV...")
        df = read_csv_id()

        if number == 1:
            print("🧾 Salvando os dados de Draft...")
            webscraping(page, df, pesquisar_id_draft, "osp_vivo_draft.xlsx")

        elif number == 2:
            print("📊 Salvando os dados de Medição...")
            webscraping(page, df, pesquisar_id_medicao, "osp_vivo_medicao.xlsx")

        elif number == 3:
            print("🔎 Salvando os dados ID OSP...")
            webscraping_id(page, df, pesquisar_id, "osp_id_cancelado.xlsx")


def open_dialog_csv():
    base_dir = Path(__file__).resolve().parent
    pasta_csv = base_dir / "lista_csv"
    pasta_csv.mkdir(parents=True, exist_ok=True)

    root = Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo CSV",
        filetypes=(("Arquivos de planilha", "*.csv"),)
    )
    root.destroy()

    if not file_path:
        print("❌ Operação cancelada!")
        return False

    destino = pasta_csv / "lista.csv"
    shutil.copy(file_path, destino)
    print(f"✅ Upload realizado com sucesso para {destino}")
    return destino


def run():
    resultado = open_dialog_csv()

    if resultado:
        while True:
            number = input(
                "Digite (1) para Draft | (2) para Medição | (3) para ID Cancelados | (0) para sair: "
            )
            try:
                number = int(number)
                if number == 0:
                    print("❌ Programa finalizado!")
                    break

                elif number in (1, 2, 3):
                    print(f"-------------- ⏱ Iniciando busca tipo {number} -------------")
                    main(number)
                    break

                else:
                    print("⚠️ Opção inválida.")
                    continue
            except ValueError:
                print("⚠️ Entrada inválida, digite um número.")
                continue
    else:
        print("❌ Programa finalizado!")


if __name__ == "__main__":
    run()
