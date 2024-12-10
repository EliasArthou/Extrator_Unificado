import time
import json
import io
import os
import pandas as pd
from PIL import Image
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from undetected_chromedriver import Chrome, ChromeOptions
from fuzzywuzzy import fuzz
from bs4 import BeautifulSoup
from anticaptchaofficial.imagecaptcha import imagecaptcha
from anticaptchaofficial.recaptchav2proxyless import recaptchaV2Proxyless
from a_selenium2df import get_df
from a_selenium_iframes_crawler import Iframes
import auxiliares as aux
import messagebox
import sensiveis as senhas
import requests
import re
from weasyprint import HTML


class TratarSite:
    """
    Classe que armazena todas as rotinas de execução de ações no controle remoto de site através do Python
    """

    def __init__(self, url, nomeperfil, caminhodownload=aux.caminhoprojeto('Downloads')):
        self.url = url
        self.perfil = nomeperfil
        self.navegador = None
        self.dataframepagina = None
        self.options = None
        self.delay = 10
        self.caminho = ''
        self.caminhodownload = caminhodownload

    def abrirnavegador(self, habilitarpdf=False):
        """
        :return: navegador configurado com o site desejado aberto
        """
        if self.navegador is not None:
            self.fecharsite()

        self.navegador = self.configuraprofilechrome(openpdf=habilitarpdf)
        if self.navegador is not None:
            self.navegador.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            self.navegador.get(self.url)
            time.sleep(1)
            resultadolimpo = ''
            corposite = BeautifulSoup(self.navegador.page_source, 'html.parser')
            for string in corposite.strings:
                resultadolimpo = resultadolimpo + ' ' + string

            if len(resultadolimpo) == 0:
                messagebox.msgbox(f'Site com problema de carregamento!', messagebox.MB_OK, 'Site fora do ar')
                self.navegador = -1
            time.sleep(1)
            self.trataralerta()
            time.sleep(1)
            return self.navegador

    def configuraprofilechrome(self, ableprintpreview=True, openpdf=True, navegarincognito=False):
        """
        Configura usuário e opções no navegador aberto para execução
        return: o navegador configurado para iniciar a execução das rotinas
        """

        prefs = {
            'profile.name': self.perfil,
            'download.default_directory': self.caminhodownload,  # Change default directory for downloads
            'download.directory_upgrade': True,
            'download.prompt_for_download': False,  # To auto download the file
            'plugins.always_open_pdf_externally': not openpdf,  # It will not show PDF directly in chrome
            "credentials_enable_service": False,  # service to save password
            "profile.password_manager_enabled": False,  # turn off password manager
            'printing.print_preview_sticky_settings.appState': json.dumps({
                'recentDestinations': [{
                    'id': 'Save as PDF',
                    'origin': 'local',
                    'account': '',
                }],
                'selectedDestinationId': 'Save as PDF',
                'version': 2,
                'defaultDestinationId': 'Save as PDF',
                'destinationSelection': [{
                    'id': 'Save as PDF',
                    'rules': {
                        'fileExtension': '.pdf',
                        'isDefault': True,
                        'saveToFolder': f"{self.caminhodownload}"
                    }
                }]
            }),
            'savefile.default_directory': self.caminhodownload
        }

        self.options = ChromeOptions()

        if aux.caminhoprojeto('Profile') != '':
            self.options.add_argument("--start-maximized")
            self.options.add_experimental_option('prefs', prefs)
            self.options.add_argument("--disable-popup-blocking")
            self.options.add_argument("--disable-infobars")
            self.options.add_argument("--disable-notifications")
            self.options.add_argument('--disable-extensions')
            self.options.add_argument('--disable-plugins')


            if ableprintpreview:
                self.options.add_argument('--kiosk-printing')
            else:
                self.options.add_argument("---printing")
                self.options.add_argument("--disable-print-preview")

            self.options.add_argument("--silent")

        if navegarincognito:
            driver = Chrome(options=self.options)
            driver.maximize_window()
        else:

            # Habilitar logs de desempenho
            capabilities = webdriver.DesiredCapabilities.CHROME.copy()
            capabilities['goog:loggingPrefs'] = {'performance': 'ALL'}

            self.options.add_argument("--disable-features=ChromeWhatsNewUI")
            self.options.add_argument("--disable-blink-features=AutomationControlled")
            self.options.add_experimental_option('excludeSwitches', ["enable-automation"])
            self.options.add_experimental_option('useAutomationExtension', False)
            chrome_service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(
                options=self.options,
                service=chrome_service,
                desired_capabilities=capabilities
            )

        return driver

    def verificar_objeto_clicavel(self, identificador, endereco, timeout=10, retry=3):
        """
        Verifica se um objeto é clicável e espera até que seja clicável.
        :param identificador: Tipo de identificador (ID, NAME, CLASS_NAME, etc.).
        :param endereco: Valor do identificador.
        :param timeout: Tempo de espera até o elemento ser clicável.
        :param retry: Número de tentativas de clicar caso o clique seja interceptado.
        :return: Elemento Web se for clicável, None caso contrário.
        """
        if self.navegador is not None:
            attempts = 0
            while attempts < retry:
                try:
                    elemento = WebDriverWait(self.navegador, timeout).until(
                        EC.element_to_be_clickable((getattr(By, identificador), endereco))
                    )
                    return elemento
                except ElementClickInterceptedException:
                    attempts += 1
                    time.sleep(5)
                except TimeoutException:
                    return None
            return None

    def verificarobjetoexiste(self, identificador, endereco, valorselecao='', itemunico=True, iraoobjeto=False,
                              sotestar=False, buscar_em_iframes=False, esperar_clicavel=False, elemento_pai=None):
        if self.navegador is not None:
            original_window = self.navegador.current_window_handle
            try:
                elemento = self.buscar_elemento(identificador, endereco, valorselecao, itemunico, iraoobjeto, sotestar,
                                                elemento_pai)
                if elemento:
                    if esperar_clicavel:
                        if itemunico:
                            elemento = self.verificar_objeto_clicavel(identificador, endereco)
                        else:
                            elementos = [self.verificar_objeto_clicavel(identificador, endereco) for el in elemento]
                            return [el for el in elementos if el is not None]
                    return elemento if not sotestar else True
                else:
                    return [] if itemunico is False else (False if sotestar else None)
            except (NoSuchElementException, TimeoutException):
                if buscar_em_iframes:
                    iframes = Iframes(driver=self.navegador, By=By, WebDriverWait=WebDriverWait, expected_conditions=EC)
                    iframes = iframes.iframes.items()
                    elementos_encontrados = []
                    for key, iframe_path in iframes:
                        for iframe in iframe_path:
                            self.navegador.switch_to.frame(iframe)
                            elementos = self.buscar_elemento(identificador, endereco, valorselecao, itemunico,
                                                             iraoobjeto, sotestar)
                            if elementos:
                                if esperar_clicavel:
                                    elementos = [(iframe, self.verificar_objeto_clicavel(identificador, endereco)) for
                                                 elemento in elementos]
                                    elementos = [el for el in elementos if el[1] is not None]
                                else:
                                    elementos = [(iframe, elemento) for elemento in elementos]
                                elementos_encontrados.extend(elementos)
                            self.navegador.switch_to.default_content()
                    if elementos_encontrados:
                        return elementos_encontrados if not sotestar else True
                return [] if itemunico is False else (False if sotestar else None)

    def buscar_elemento(self, identificador, endereco, valorselecao='', itemunico=True, iraoobjeto=False,
                        sotestar=False, elemento_pai=None):
        try:
            by_type = getattr(By, identificador)
            elemento = None

            if itemunico:
                if not elemento_pai:
                    elemento = WebDriverWait(self.navegador, self.delay).until(
                        EC.visibility_of_element_located((by_type, endereco)))
                else:
                    elemento = elemento_pai.find_element(by_type, endereco)
                if valorselecao:
                    select = Select(elemento)
                    select.select_by_visible_text(valorselecao)
            else:
                if not elemento_pai:
                    elementos = self.navegador.find_elements(by_type, endereco)
                else:
                    elementos = elemento_pai.find_elements(by_type, endereco)
                if not elementos:
                    elemento = None
                elif valorselecao:
                    select = Select(elementos)
                    select.select_by_visible_text(valorselecao)
                else:
                    elemento = elementos

            if iraoobjeto and elemento:
                self.navegador.execute_script("arguments[0].click()", elemento)
                time.sleep(1)
                self.trataralerta()
                time.sleep(1)

            if sotestar:
                return elemento is not None

            return elemento
        except TimeoutException:
            return False if sotestar else None
        except Exception as e:
            return False if sotestar else None

    def descerrolagem(self):
        self.navegador.execute_script("window.scrollTo(0,document.body.scrollHeight)")
        time.sleep(1)

    def mexerzoom(self, valor):
        self.navegador.execute_script("document.body.style.transform='scale(" + str(valor) + ")';")
        time.sleep(1)

    def baixarimagem(self, identificador, endereco, caminho):
        achouimagem = False
        salvouimagem = False
        if self.navegador is not None:
            if os.path.isfile(caminho):
                os.remove(caminho)
            elemento = self.verificarobjetoexiste(identificador, endereco)
            if elemento is not None:
                achouimagem = True
                image = elemento.screenshot_as_png
                imagestream = io.BytesIO(image)
                im = Image.open(imagestream)
                im.save(caminho)
                if os.path.isfile(caminho):
                    if aux.quantidade_cores(caminho) > 1:
                        salvouimagem = True
                        self.caminho = caminho
                    else:
                        os.remove(caminho)
                        salvouimagem = False
                else:
                    salvouimagem = False

        return achouimagem, salvouimagem

    @staticmethod
    def esperadownloads(caminho, timeout, nfiles=None):
        time.sleep(1)
        seconds = 0
        dl_wait = True
        while dl_wait and seconds < timeout:
            time.sleep(1)
            dl_wait = False
            files = os.listdir(caminho)
            if nfiles and len(files) != nfiles:
                dl_wait = True

            for fname in files:
                if fname.endswith('.crdownload'):
                    dl_wait = True

            seconds += 1

        time.sleep(1)
        return seconds

    def num_abas(self):
        if self.navegador is not None:
            return len(self.navegador.window_handles)

    def irparaaba(self, indice=-1, titulo=''):
        achouaba = False

        if indice >= 0 or len(titulo) > 0:
            if indice >= 0:
                if indice <= self.num_abas():
                    self.navegador.switch_to.window(self.navegador.window_handles[indice - 1])
            else:
                janelaatual = self.navegador.current_window_handle
                handles = self.navegador.window_handles
                for hdl in handles:
                    self.navegador.switch_to.window(hdl)
                    if titulo.upper() in self.navegador.title.upper() + self.navegador.current_url.upper():
                        achouaba = True
                        break
                if not achouaba:
                    self.navegador.switch_to.window(janelaatual)

                return achouaba

    def fecharaba(self, indice=0):
        if indice == 0:
            self.navegador.execute_script('window.open("","_self").close()')
        else:
            self.irparaaba(indice)
            self.navegador.execute_script('window.open("","_self").close()')

    def irparaframe(self, frame):
        self.navegador.switch_to.frame(frame)

    def sairdoframe(self):
        self.navegador.switch_to.default_content()

    def trataralerta(self):
        from selenium.webdriver.common.alert import Alert
        from selenium.common.exceptions import NoAlertPresentException

        try:
            alerta = Alert(self.navegador)
            if alerta is not None:
                textoalerta = alerta.text
                alerta.accept()
                time.sleep(1)
            else:
                textoalerta = ''

        except NoAlertPresentException:
            textoalerta = ''
            pass

        return textoalerta

    def fecharsite(self):
        if self.navegador is not None and hasattr(self.navegador, 'quit'):
            self.navegador.quit()

    def resolvercaptcha(self, identificacaocaixa, caixacaptcha, identicacaobotao, botao):
        resposta = False
        textoerro = ''
        solver = imagecaptcha()
        solver.set_verbose(1)
        solver.set_key(senhas.chaveanticaptcha)

        captcha_text = solver.solve_and_return_solution(self.caminho)

        if len(str(captcha_text)) != 0:
            captcha = self.verificarobjetoexiste(identificacaocaixa, caixacaptcha, iraoobjeto=True)
            captcha.send_keys(captcha_text)
            time.sleep(1)
            self.verificarobjetoexiste(identicacaobotao, botao, iraoobjeto=True)

            try:
                mensagemerro = self.trataralerta()
                if mensagemerro == '':
                    pagina = BeautifulSoup(self.navegador.page_source, 'html.parser')
                    objeto_encontrado = pagina.find(
                        lambda tag: tag.name == 'font' and tag.get('size') == '2' and tag.get('color') == 'red')
                    if objeto_encontrado:
                        if objeto_encontrado.text.strip():
                            mensagemerro = objeto_encontrado.text.strip()

                if mensagemerro == 'Código digitado não confere! Favor refazer a consulta!':
                    solver.report_incorrect_image_captcha()
                    botao_nova_consulta = self.verificarobjetoexiste('XPATH',
                                                                     "//input[@type='button' and @name='bt' and @value='Nova Consulta']")
                    if botao_nova_consulta is not None:
                        botao_nova_consulta.click()
                    textoerro = mensagemerro
                    resposta = False
                else:
                    textoerro = mensagemerro
                    resposta = True

                if os.path.isfile(self.caminho):
                    os.remove(self.caminho)
                    self.caminho = ''

            except TimeoutException:
                resposta = True
                if os.path.isfile(self.caminho):
                    os.remove(self.caminho)
                    self.caminho = ''

        else:
            print("Erro solução captcha:" + solver.error_code)
            resposta = False

        return resposta, textoerro

    def resolvecaptchatipo2(self, chavecaptcha):
        solver = recaptchaV2Proxyless()
        solver.set_verbose(1)
        solver.set_key(senhas.chaveanticaptcha)
        solver.set_website_url(self.navegador.current_url)
        solver.set_website_key(chavecaptcha)
        solver.set_soft_id(0)

        g_response = solver.solve_and_return_solution()
        if g_response != 0:
            return g_response
        else:
            print("task finished with error " + solver.error_code)
            return None

    def retornartabela(self, tipolista):
        import re
        import unicodedata

        linha = []
        cabecalhos = []
        tabela = []
        intermediario = ''

        pagina = BeautifulSoup(self.navegador.page_source, 'html.parser')

        match tipolista:
            case 1:
                table = pagina.find('table', {'border': 0, 'width': 500})
                for element in table.find_all('td'):
                    if len(element.text.strip()) > 0:
                        if aux.right(element.text, 1) == ':':
                            intermediario = aux.to_raw(element.text)
                            cabecalhos.append(intermediario)
                        else:
                            elementotratado = BeautifulSoup(str(element), 'html.parser')
                            for br in elementotratado.find_all("br"):
                                br.replace_with("\n")
                            intermediario = unicodedata.normalize("NFKD", aux.to_raw(elementotratado.text))
                            linha.append(intermediario)

            case 2:
                for element in pagina.find_all('input', {'name': re.compile('^taxa')}):
                    if element['name'] not in cabecalhos:
                        if len(element['name']) > 0 and len(element['value']) > 0:
                            linha.append(element['value'])
                            cabecalhos.append(element['name'])
                    else:
                        tabela.append(dict(zip(cabecalhos, linha)))
                        cabecalhos = [element['name']]
                        linha = [element['value']]

            case 3:
                table = pagina.find('table')
                itenstabela = table.find_all('td')
                if len(itenstabela) == 1:
                    for element in itenstabela:
                        if len(element.text.strip()) > 0:
                            elementotratado = BeautifulSoup(str(element), 'html.parser')
                            botao = elementotratado.find('button')
                            if botao is not None:
                                botao.decompose()
                            else:
                                botao = elementotratado.find('a', {'href': 'javascript: history.back()'})
                                if botao is not None:
                                    botao.decompose()

                            for br in elementotratado.find_all("br"):
                                br.replace_with("\n")
                            intermediario = unicodedata.normalize("NFKD", aux.to_raw(elementotratado.text))
                            intermediario = re.sub('\n+', '\n', intermediario)
                            intermediario = intermediario.replace(
                                'Emissão de 2a Via da Taxa e Consulta a Débitos Anteriores', '')
                            intermediario = intermediario.strip()

        if tipolista != 3:
            if len(linha) > 0 and len(cabecalhos) > 0:
                tabela.append(dict(zip(cabecalhos, linha)))
        else:
            tabela = intermediario

        return tabela

    @staticmethod
    def obter_arquivos_atuais(diretorio):
        return set(os.listdir(diretorio))

    @staticmethod
    def esperar_novo_download(diretorio, arquivos_originais, timeout=300):
        tempo_inicio = time.time()
        while True:
            arquivos_atuais = set(os.listdir(diretorio))
            arquivos_novos = arquivos_atuais - arquivos_originais
            for arquivo in arquivos_novos:
                if not arquivo.endswith('.crdownload'):
                    print(f"Download completo: {arquivo}")
                    return arquivo
            if (time.time() - tempo_inicio) > timeout:
                print("Tempo de espera excedido.")
                return None
            time.sleep(1)

    def pegaarquivobaixado(self, timeout, quantabas=0, caminhobaixado=''):
        """
        Monitora o download de arquivos usando a aba chrome://downloads.
        """
        if quantabas > 0:
            while self.num_abas() > quantabas:
                time.sleep(1)
        quantabas = self.num_abas()

        # Abre a aba de downloads
        self.navegador.execute_script('window.open()')
        time.sleep(2)

        if quantabas < self.num_abas():
            self.navegador.switch_to.window(self.navegador.window_handles[-1])
            self.navegador.get('chrome://downloads')
            time.sleep(5)

            end_time = time.time() + timeout
            download_completed = False

            while time.time() < end_time:
                try:
                    # Obtém os itens de download
                    download_items = self.navegador.execute_script("""
                        let downloadsManager = document.querySelector('downloads-manager');
                        if (!downloadsManager) return null;

                        let shadowRoot = downloadsManager.shadowRoot;
                        if (!shadowRoot) return null;

                        let items = shadowRoot.querySelectorAll('downloads-item');
                        if (!items) return [];

                        return Array.from(items).map(item => {
                            let itemShadowRoot = item.shadowRoot;
                            let fileName = itemShadowRoot.querySelector('#file-link').textContent.trim();
                            let state = itemShadowRoot.querySelector('#tag').textContent.trim();
                            return { fileName: fileName, state: state };
                        });
                    """)

                    # Verifica os downloads concluídos
                    for item in download_items:
                        if item['state'] == 'Concluído':  # Verifica o estado do download
                            file_name = item['fileName']
                            file_path = os.path.join(self.caminhodownload, file_name)

                            if os.path.exists(file_path):
                                print(f"Arquivo baixado identificado: {file_path}")
                                self.fecharaba()  # Fecha a aba de downloads
                                return file_path

                except Exception as e:
                    print(f"Erro ao monitorar downloads: {e}")

                time.sleep(1)

            # Fecha a aba de downloads após o timeout
            if self.num_abas() > 1:
                self.fecharaba()

            print("Tempo limite para download excedido.")
            return None

    def retornarpaginaemdf(self):
        self.dataframepagina = get_df(self.navegador, By, WebDriverWait, EC, queryselector='*', with_methods=True)

    def buscarobjetoemdf(self, filtros: dict, atributo: str = None):
        if self.dataframepagina is None:
            raise ValueError("Dataframe não iniciado!")

        df = self.dataframepagina
        query_conditions = []

        for key, value in filtros.items():
            identificador = f"aa_{key}" if "_" not in key else key
            match value:
                case '^NaN':
                    condition = f"({identificador}.notna())"
                case value if pd.isna(value):
                    condition = f"({identificador}.isna())"
                case _:
                    if df[identificador].dtype == 'string' or df[identificador].dtype == 'object':
                        condition = f"({identificador} == '{value}')"
                    else:
                        condition = f"({identificador} == {value})"

            query_conditions.append(condition)

        if query_conditions:
            query_string = " & ".join(query_conditions)
            dffiltrado = df.query(query_string)

            if atributo:
                if "_" not in atributo:
                    atributo = f"aa_{atributo}"
                if atributo in dffiltrado.columns:
                    if len(dffiltrado) == 1:
                        return dffiltrado[atributo].iloc[0]
                    else:
                        return dffiltrado[atributo].tolist()
                else:
                    raise ValueError(f"Atributo '{atributo}' não encontrado no DataFrame.")
            else:
                return dffiltrado
        else:
            raise ValueError("Nenhum filtro fornecido!")

    def encontrar_input_oculto(self, identificador, endereco, buscar_em_iframes=False):
        try:
            hidden_input = self.navegador.find_element(getattr(By, identificador), endereco)
            print(hidden_input.get_attribute('value'))
            return hidden_input
        except NoSuchElementException:
            if buscar_em_iframes:
                iframes = self.navegador.find_elements(By.TAG_NAME, 'iframe')
                original_window = self.navegador.current_window_handle
                for iframe in iframes:
                    try:
                        self.navegador.switch_to.frame(iframe)
                        hidden_input = self.navegador.find_element(getattr(By, identificador), endereco)
                        print(hidden_input.get_attribute('value'))
                        return hidden_input
                    except NoSuchElementException:
                        continue
                    finally:
                        self.navegador.switch_to.window(original_window)
        return None

    def buscar_por_proximidade(self, identificador, seletor, texto_de_busca, corte_perc=70, elemento_do_item=False):
        opcao_encontrada = None
        elemento_encontrado = None
        if not elemento_do_item:
            elementos = self.navegador.find_elements(getattr(By, identificador), seletor)
        else:
            elementos = self.navegador.find_element(getattr(By, identificador), seletor)
            if elementos.tag_name.lower() != "select":
                elementos = self.navegador.find_elements(getattr(By, identificador), seletor)

        texto_mais_parecido = ''
        percentual_de_compatibilidade = 0
        elemento_encontrado = None

        if isinstance(elementos, list):
            for elemento in elementos:
                texto_elemento = elemento.text.lower()
                texto_de_busca_lower = texto_de_busca.lower()
                compatibilidade = fuzz.ratio(texto_de_busca_lower, texto_elemento)

                if compatibilidade > percentual_de_compatibilidade:
                    texto_mais_parecido = texto_elemento
                    percentual_de_compatibilidade = compatibilidade
                    elemento_encontrado = elemento
        else:
            if elementos.tag_name.lower() == "select":
                select = Select(elementos)
                for option in select.options:
                    texto_option = option.text.lower()
                    texto_de_busca_lower = texto_de_busca.lower()
                    compatibilidade = fuzz.ratio(texto_de_busca_lower, texto_option)

                    if compatibilidade > percentual_de_compatibilidade:
                        texto_mais_parecido = texto_option
                        percentual_de_compatibilidade = compatibilidade
                        opcao_encontrada = option
                        texto_option_selecionado = option.text.lower()
                        elemento_encontrado = option
                        option.click()

        if elemento_encontrado and percentual_de_compatibilidade >= corte_perc:
            resposta = (f"Elemento encontrado: {texto_mais_parecido} com {percentual_de_compatibilidade}% "
                        f"de compatibilidade com o item {texto_de_busca}.")
            if opcao_encontrada:
                resposta = resposta + f" Item selecionado: {texto_option_selecionado}"
            print(resposta)
            return elemento_encontrado
        else:
            print(f"Nenhum elemento encontrado com pelo menos {corte_perc}% de compatibilidade com o item {texto_de_busca}.")
            if elemento_encontrado:
                resposta = (f"O elemento mais próximo encontrado foi: {texto_mais_parecido} "
                            f"com {percentual_de_compatibilidade}% de compatibilidade.")

                if opcao_encontrada:
                    resposta = resposta + f" Item selecionado: {texto_option_selecionado}"

                print(resposta)
            return None


    # def monitorar_downloads_sem_href(self, link_elemento, timeout=300, clickscript=False, via_request=False,
    #                                  url_override=None):
    #     """
    #     Monitora ou realiza o download diretamente de um link_elemento.
    #
    #     Parâmetros:
    #     - link_elemento: WebElement que inicia o download.
    #     - timeout: Tempo máximo (em segundos) para aguardar o download.
    #     - clickscript: Se True, usa `execute_script` para clicar no elemento.
    #     - via_request: Se True, realiza o download via `requests` ao invés de clicar no link.
    #     - url_override: URL manual para realizar o download, usado em redirecionamentos.
    #
    #     Retorna:
    #     - O caminho completo do arquivo baixado ou None em caso de falha.
    #     """
    #     print("Iniciando monitoramento/download...")
    #
    #     if not via_request:
    #         # Modo padrão: clique no elemento e monitoramento
    #         if not clickscript:
    #             link_elemento.click()
    #         else:
    #             self.navegador.execute_script("arguments[0].click()", link_elemento)
    #
    #         time.sleep(2)
    #         print("Clique disparado. Monitorando logs e pasta de downloads...")
    #
    #         tempo_inicio = time.time()
    #         download_guid = None
    #         suggested_filename = None
    #         arquivo_final = None
    #
    #         while time.time() - tempo_inicio < timeout:
    #             logs = self.navegador.get_log("performance")
    #             for log in logs:
    #                 message = json.loads(log["message"])
    #                 method = message["message"]["method"]
    #                 print(message)
    #
    #                 # Detecta abertura de uma nova janela (window.open)
    #                 if method == "Page.windowOpen":
    #                     params = message["message"]["params"]
    #                     url_opened = params.get("url")
    #                     print(f"Nova janela detectada: URL={url_opened}")
    #                     # Confirma se o URL é de download e redireciona para request
    #                     if "boleto" in url_opened.lower():
    #                         print(f"Detectado comportamento de redirecionamento para download: {url_opened}")
    #                         return self.monitorar_downloads_sem_href(
    #                             link_elemento=None, timeout=timeout, clickscript=False, via_request=True,
    #                             url_override=url_opened
    #                         )
    #
    #                 # Detecta início do download
    #                 if method == "Page.downloadWillBegin":
    #                     params = message["message"]["params"]
    #                     download_guid = params.get("guid")
    #                     suggested_filename = params.get("suggestedFilename")
    #                     print(f"Download iniciado: GUID={download_guid}, Nome sugerido={suggested_filename}")
    #
    #                 # Detecta progresso ou finalização
    #                 if method == "Page.downloadProgress":
    #                     params = message["message"]["params"]
    #                     current_guid = params.get("guid")
    #                     state = params.get("state")
    #
    #                     if current_guid == download_guid:
    #                         if state == "completed":
    #                             print(f"Download concluído: GUID={current_guid}")
    #
    #                             # Identifica o arquivo baixado na pasta
    #                             arquivos_no_diretorio = os.listdir(self.caminhodownload)
    #                             candidatos = [
    #                                 os.path.join(self.caminhodownload, arquivo)
    #                                 for arquivo in arquivos_no_diretorio
    #                                 if suggested_filename and arquivo.startswith(suggested_filename.split('.')[0])
    #                             ]
    #
    #                             if candidatos:
    #                                 candidatos.sort(key=os.path.getmtime, reverse=True)
    #                                 arquivo_final = candidatos[0]
    #                                 print(f"Arquivo baixado identificado: {arquivo_final}")
    #                                 return arquivo_final
    #                             else:
    #                                 print("Arquivo não encontrado na pasta de downloads.")
    #                                 return None
    #
    #                 time.sleep(1)
    #
    #         print("Tempo limite para monitoramento excedido.")
    #         return None
    #
    #     else:
    #         # Exceção: Realiza o download via requests
    #         href = url_override or (link_elemento.get_attribute("href") if link_elemento else None)
    #         if not href:
    #             print("Elemento ou URL manual não fornecido. Não é possível fazer o download via request.")
    #             return None
    #
    #         # Define o nome do arquivo a partir da URL, removendo caracteres inválidos
    #         suggested_filename = 'baixado.pdf' #re.sub(r'[<>:"/\\|?*]', '_', href.split("/")[-1])
    #         arquivo_final = os.path.join(self.caminhodownload, suggested_filename)
    #
    #         try:
    #             # Captura os cookies ativos do navegador
    #             cookies = self.navegador.get_cookies()
    #             cookies_dict = {cookie['name']: cookie['value'] for cookie in cookies}
    #
    #             # Configura o User-Agent do navegador
    #             headers = {
    #                 'User-Agent': self.navegador.execute_script("return navigator.userAgent;")
    #             }
    #
    #             # Realiza o download via requests
    #             print(f"Baixando arquivo diretamente de: {href}")
    #             response = requests.get(href, headers=headers, cookies=cookies_dict, stream=True, timeout=timeout)
    #             response.raise_for_status()
    #
    #             with open(arquivo_final, "wb") as file:
    #                 for chunk in response.iter_content(chunk_size=8192):
    #                     file.write(chunk)
    #
    #             print(f"Download concluído via request: {arquivo_final}")
    #             return arquivo_final
    #
    #         except requests.RequestException as e:
    #             if os.path.isfile(arquivo_final):
    #                 os.remove(arquivo_final)
    #             print(f"Erro ao baixar arquivo via requests: {e}")
    #             return None

    def monitorar_downloads_sem_href(self, link_elemento, timeout=300, clickscript=False, via_request=False,
                                     url_override=None):
        """
        Monitora ou realiza o download diretamente de um link_elemento.

        Parâmetros:
        - link_elemento: WebElement que inicia o download.
        - timeout: Tempo máximo (em segundos) para aguardar o download.
        - clickscript: Se True, usa `execute_script` para clicar no elemento.
        - via_request: Se True, realiza o download via `requests` ao invés de clicar no link.
        - url_override: URL manual para realizar o download, usado em redirecionamentos.

        Retorna:
        - O caminho completo do arquivo baixado ou None em caso de falha.
        """
        print("Iniciando monitoramento/download...")

        if not via_request:
            original_abas = len(self.navegador.window_handles)
            original_url = self.navegador.current_url

            # Clique no elemento ou execute o script
            if not clickscript:
                link_elemento.click()
            else:
                self.navegador.execute_script("arguments[0].click()", link_elemento)

            time.sleep(2)
            print("Clique disparado. Monitorando logs e pasta de downloads...")

            tempo_inicio = time.time()
            download_guid = None
            suggested_filename = None
            arquivo_final = None

            while time.time() - tempo_inicio < timeout:
                logs = self.navegador.get_log("performance")
                for log in logs:
                    message = json.loads(log["message"])
                    method = message["message"]["method"]

                    # Detecta início do download
                    if method == "Page.downloadWillBegin":
                        params = message["message"]["params"]
                        download_guid = params.get("guid")
                        suggested_filename = params.get("suggestedFilename")
                        print(f"Download iniciado: GUID={download_guid}, Nome sugerido={suggested_filename}")

                    # Detecta progresso e finalização
                    if method == "Page.downloadProgress":
                        params = message["message"]["params"]
                        current_guid = params.get("guid")
                        state = params.get("state")

                        if current_guid == download_guid:
                            if state == "completed":
                                print(f"Download concluído: GUID={current_guid}")

                                # Identifica o arquivo na pasta de downloads
                                arquivos_no_diretorio = os.listdir(self.caminhodownload)
                                candidatos = [
                                    os.path.join(self.caminhodownload, arquivo)
                                    for arquivo in arquivos_no_diretorio
                                    if suggested_filename and arquivo.startswith(suggested_filename.split('.')[0])
                                ]
                                if candidatos:
                                    candidatos.sort(key=os.path.getmtime, reverse=True)
                                    arquivo_final = candidatos[0]
                                    print(f"Arquivo baixado identificado: {arquivo_final}")
                                    return arquivo_final
                                else:
                                    print("Arquivo não encontrado na pasta de downloads.")
                                    return None

                    # Detecta nova aba aberta
                    if method == "Page.windowOpen":
                        params = message["message"]["params"]
                        url_opened = params.get("url")
                        print(f"Nova aba detectada: URL={url_opened}")

                        # Troca para a nova aba se uma foi aberta
                        if len(self.navegador.window_handles) > original_abas:
                            self.navegador.switch_to.window(self.navegador.window_handles[-1])
                            print("Nova aba detectada. Mudando para a nova aba.")
                            time.sleep(2)

                            # Procura links de download na aba
                            download_link = self.navegador.execute_script("""
                                const links = document.querySelectorAll('a, iframe, embed');
                                for (let link of links) {
                                    if (link.href && link.href.endsWith('.pdf')) return link.href;
                                    if (link.src && link.src.endsWith('.pdf')) return link.src;
                                }
                                return null;
                            """)

                            if download_link:
                                print(f"Link de download encontrado na nova aba: {download_link}")
                                self.navegador.close()
                                self.navegador.switch_to.window(self.navegador.window_handles[0])
                                return self.monitorar_downloads_sem_href(
                                    link_elemento=None, timeout=timeout, clickscript=False, via_request=True,
                                    url_override=download_link
                                )

                            print("Nenhum link de download detectado na nova aba.")
                            self.navegador.close()
                            self.navegador.switch_to.window(self.navegador.window_handles[0])

                time.sleep(1)

            print("Tempo limite para monitoramento excedido.")
            return None

        else:
            # Download via requests
            href = url_override or (link_elemento.get_attribute("href") if link_elemento else None)
            if not href:
                print("Elemento ou URL manual não fornecido. Não é possível fazer o download via request.")
                return None

            # Ajusta URL caso necessário
            if 'visualizar-arquivo' in href:
                href = href.replace('visualizar-arquivo', 'visualizar-conteudo-do-arquivo')

            # Nome do arquivo
            suggested_filename = re.sub(r'[<>:"/\\|?*]', '_', href.split("/")[-1]) or 'arquivo_baixado.pdf'
            arquivo_final = os.path.join(self.caminhodownload, suggested_filename)

            try:
                # Captura cookies e configura headers
                cookies = self.navegador.get_cookies()
                cookies_dict = {cookie['name']: cookie['value'] for cookie in cookies}
                headers = {
                    'User-Agent': self.navegador.execute_script("return navigator.userAgent;"),
                    'Referer': href
                }

                # Download via requests
                print(f"Baixando arquivo diretamente de: {href}")
                response = requests.get(href, headers=headers, cookies=cookies_dict, stream=True, timeout=timeout)
                response.raise_for_status()

                # Verifica se é um PDF ou HTML
                content_type = response.headers.get('Content-Type', '').lower()
                if 'pdf' in content_type:
                    with open(arquivo_final, "wb") as file:
                        for chunk in response.iter_content(chunk_size=8192):
                            file.write(chunk)
                    print(f"Download concluído via request: {arquivo_final}")
                    return arquivo_final
                else:
                    # Converte HTML para PDF
                    print("O conteúdo não é PDF, convertendo para PDF...")
                    return self.converte_html_para_pdf(response.text, arquivo_final)

            except requests.RequestException as e:
                if os.path.isfile(arquivo_final):
                    os.remove(arquivo_final)
                print(f"Erro ao baixar arquivo via requests: {e}")
                return None

    def converte_html_para_pdf(self, html_content, output_path):
        """
        Converte conteúdo HTML em PDF.

        Parâmetros:
        - html_content: String contendo o código HTML.
        - output_path: Caminho onde o PDF será salvo.

        Retorna:
        - Caminho do arquivo PDF gerado ou None em caso de falha.
        """
        try:
            HTML(string=html_content).write_pdf(output_path)
            print(f"PDF gerado com sucesso: {output_path}")
            return output_path
        except Exception as e:
            print(f"Erro ao converter HTML para PDF: {e}")
            return None

    def filtrar_itens_por_fuzzy(self, lista, identificador, valor_identificador, texto_condominio, texto_apartamento, corte_condominio=70, corte_apartamento=70):
        """
        Filtra itens de uma lista com base em uma correspondência fuzzy com os textos de título e secundário.
        Retorna duas listas: uma com itens baseados no título e outra com itens baseados no texto secundário.

        Parâmetros:
        - lista: Lista de elementos para verificar.
        - texto_condominio: Texto de referência para o título.
        - texto_apartamento: Texto de referência para o secundário.
        - corte_condominio: Percentual de corte para o condomínio (default 70%).
        - corte_apartamento: Percentual de corte para o apartamento (default 70%).

        Retorna:
        - Lista de candidatos baseados no título (índice e elemento).
        - Lista de candidatos baseados no texto secundário (índice e elemento).
        """
        candidatos_titulo = []
        candidatos_secundario = []
        correspondenciaexata = False

        for idx, elemento in enumerate(lista):
            try:
                # Obtém o título e o texto secundário
                titulo = self.verificarobjetoexiste(identificador, valor_identificador, itemunico=True, elemento_pai=elemento)
                titulo_texto = titulo.text.strip() if titulo else ""
                secundario_texto = elemento.text.split('\n')
                if len(secundario_texto) >= 2:
                    secundario_texto = secundario_texto[1]
                    # Calcula a similaridade para o texto secundário
                    if secundario_texto.lower() == texto_apartamento.lower():
                        return True, [], (idx, elemento)
                    similaridade_secundario = fuzz.ratio(secundario_texto.lower(), texto_apartamento.lower())
                    if similaridade_secundario >= corte_apartamento:
                        candidatos_secundario.append((idx, elemento))
                else:
                    secundario_texto = ''

                # Calcula a similaridade para o título
                similaridade_titulo = fuzz.ratio(titulo_texto.lower(), texto_condominio.lower())
                if similaridade_titulo >= corte_condominio:
                    candidatos_titulo.append((idx, elemento))

            except Exception as e:
                print(f"Erro ao processar elemento {idx}: {e}")

        return False, candidatos_titulo, candidatos_secundario
