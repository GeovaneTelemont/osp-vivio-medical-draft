from playwright.sync_api import sync_playwright, TimeoutError
import os, csv
import pandas as pd
from time import sleep


AUTH_FILE = "auth.json"
LOGIN_URL = "https://devopsredes.vivo.com.br/ospcontrol/home"
LOGGED_SELECTOR = 'xpath=//*[@id="ott-username"]' 
USERNAME = "80969154"  # 🔒 Preencha ou use input()
PASSWORD = "Ca0109le@"  # 🔒 Preencha ou use input()


def read_csv_id():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_path = os.path.join(script_dir, "./lista_csv/lista.csv")
    df = pd.read_csv(csv_path, sep=";", encoding="utf-8")
    return df


def login(page):
    """Preenche usuário e senha, espera login manual (CAPTCHA)."""
    print("Efetuando login...")
    try:
        page.wait_for_selector('//*[@id="username"]', state="visible", timeout=10000)
        page.fill('//*[@id="username"]', USERNAME)
        page.fill('//*[@id="password"]', PASSWORD)
        print("✅ Usuário e senha preenchidos. Complete o CAPTCHA e clique em login manualmente...")
        
        # Aguarda até que o seletor de login bem-sucedido apareça
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


def pesquisar_id(page, id):
    try:
        #page.goto("https://devopsredes.vivo.com.br/ospcontrol/requisicoes-eps")
        sleep(2)
        page.click('//*[@id="ott-sidebar-collapse"]', timeout=10000)
        
        sleep(2)
        page.click('//*[@id="ott-sidebar"]/div[3]/ul/li[3]/a', timeout=10000)
        sleep(2)
        page.fill('xpath=//*[@id="filtroId"]', str(id))

        sleep(2)
        page.locator("a.btn.btn-primary.btn-sm.btn-block:has-text('Buscar')").click()
     
        sleep(2)
        #Clica no botão editar
        page.locator("span.badge.bg-primary:has-text('Editar')").click()
       
        sleep(2)
        
        links = page.locator("a.nav-link")
        
        total = int(links.count())

        for i in range(total):
            texto = links.nth(i).text_content().strip()
            if texto == "Medição":
                links.nth(i).click()

        print("Serviço")
        #page.locator('a[title="Serviços"]').wait_for(state="visible", timeout=40000)
        sleep(3)
        page.locator('a[title="Serviços"]').click()
        

        sleep(2)
        page.wait_for_selector('xpath=(//table)[2]', state="visible", timeout=10000)

        
        sleep(2)
        # Seleciona a segunda tabela
        tabela2 = page.query_selector('xpath=(//table)[2]')
        sleep(2)
        # --- Pegar cabeçalhos ---
        ths = tabela2.query_selector_all("thead th")
        colunas = [th.inner_text().strip() for th in ths]

        # --- Pegar todas as linhas ---
        linhas = tabela2.query_selector_all("tbody tr")
        dados = []

        for linha in linhas:
            tds = linha.query_selector_all("td")
            
            # Ignora linhas com menos colunas (linha do total)
            if len(tds) != len(colunas):
                continue
            valores = []
            valores.append(id)

            for td in tds:
                td = td.inner_text().strip()
                valores.append(td)
           
            dados.append(valores)

        print(dados)
        sleep(2)
        page.click('//*[@id="ott-sidebar-collapse"]')
        sleep(2)

        page.click('//*[@id="ott-sidebar"]/div[3]/ul/li[1]/a')
        sleep(2)
        print(f"✅ ID: {id} tabela salva com sucesso!")

        return dados

    except Exception as e:
        print(f"Erro ao pesquisar ID {False}: {e}")
        return None


def main():
    with sync_playwright() as p:
        # --- Browser ---
        print("Iniciando navegador...")
        browser = p.chromium.launch(channel="chrome", headless=False, args=["--ignore-certificate-errors"])

        # --- Cria ou carrega sessão ---
        if os.path.exists(AUTH_FILE):
            print("Carregando sessão existente de", AUTH_FILE)
            context = browser.new_context(storage_state=AUTH_FILE)
        else:
            print("Nenhuma sessão encontrada. Criando uma nova.")
            context = browser.new_context()

        page = context.new_page()
        page.goto(LOGIN_URL, wait_until="networkidle")
        print("✅ Página carregada.")

        # --- Verifica se está logado ---
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

        # --- Ações após login ---
        print("Lendo CSV...")
        df = read_csv_id()
        #print(df.head())

        # Exemplo: acessar uma URL para cada linha (ajuste conforme sua necessidade)
        for i, row in df.iterrows():
            url = str(row.get("url", ""))
            if not url:
                continue
            print(f"Abrindo: {url}")
            try:
                page.goto(url, wait_until="networkidle", timeout=60000)
                print("Título:", page.title())
            except Exception as e:
                print("Erro ao acessar", url, "->", e)

        for index, row in df.iterrows():
            id_value = row["ID"]
            osp_value = row["OSP MEDIDO"]
            dados = pesquisar_id(page, id_value)
            
            colunas = ["ID","CODIGO", "DESCRIÇÃO DOS SERVIÇOS", "QUANTIDADE", "PREÇO UNITÁRIO", "UNIDADE", "PREÇO TOTAL"]
            df = pd.DataFrame(dados, columns=colunas)
            df.to_csv("osp_vivo.csv", mode="a", index=False, header=False, encoding="utf-8")
            
        # input("Pressione Enter para fechar...")
        #context.close()
        #browser.close()


if __name__ == "__main__":
    main()