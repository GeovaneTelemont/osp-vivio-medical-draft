from playwright.sync_api import sync_playwright, TimeoutError
import os, unicodedata, shutil
import pandas as pd
from time import sleep
from tkinter import filedialog
from tkinter import Tk
from pathlib import Path

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

def _normalize_text(s: str) -> str:
    if not s:
        return ""
    s = " ".join(s.split())  # colapsar espaços
    # remover acentos e lower
    s_norm = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("ASCII").lower()
    return s_norm

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


def webscraping(page, df, func_pesquisa, name_file_csv):
    for index, row in df.iterrows():
        id_value = row["ID"]
        osp_value = row["OSP MEDIDO"]
        dados = func_pesquisa(page, id_value)
        colunas = ["ID", "TIPO DE REGISTRO", "CÓDIGO", "DESCRIÇÃO", "QUANTIDADE", "PREÇO UNITÁRIO", "UNIDADE","PREÇO TOTAL", "CATEGORIA"]
        arquivo = name_file_csv
        df = pd.DataFrame(dados, columns=colunas)

        if not os.path.isfile(arquivo):
            # se o arquivo não existir cria cabeçalho
            df.to_csv(name_file_csv, mode="w", index=False, header=True, encoding="utf-8")
        else:
            # se o arquivo existir continuar cadastrando os dados sem cabeçalho
            df.to_csv(name_file_csv, mode="a", index=False, header=False, encoding="utf-8")

        
def main(number):
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
   
        if number == 1:
            print("Salvando os dados de Draft")
            webscraping(page, df, pesquisar_id_draft, "osp_vivo_draft.csv")

        if number == 2:
            print("Salvando os dados de Medição")
            webscraping(page, df, pesquisar_id_medicao, "osp_vivo_medicao.csv")


def open_dialog_csv():
    base_dir = Path(__file__).resolve().parent
    pasta_csv = base_dir / "lista_csv"
    pasta_csv.mkdir(parents=True, exist_ok=True)

    root = Tk()
    root.withdraw()

    # Usuário escolhe um arquivo existente (para copiar)
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo CSV",
        filetypes=(("Arquivos de planilha", "*.csv"),)
    )

    root.destroy()

    if not file_path:
        print("❌ Operação cancelada!")
        return False

    destino = pasta_csv / "lista.csv"

    # Copia o arquivo selecionado para dentro da pasta lista_csv
    shutil.copy(file_path, destino)
    print(f"✅ Upload realizado com")

    return destino

def run():
    resultado = open_dialog_csv()
    
    if resultado:

        while True:
            number = input("Digite (1) para Draft ou Ditite (2) para Medição ou (0) para sair: ")
            try:
                if int(number) == 0:
                    print("❌ Programa finalizado!")
                    break

                elif int(number) == 1:
                    print("------------ ⏱ Iniciando busca de dados Draft! ------------")
                    main(int(number))
                    break
                
                elif int(number) == 2:
                    print("-------------- ⏱ Iniciando busca de dados Medição! -------------")
                    main(int(number))
                    break
                else:
                    continue
            except:
                continue
    else:
        print("❌ Programa finalizado!")


if __name__ == "__main__":
    run()
        
        