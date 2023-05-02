from selenium import webdriver
import io
from PIL import Image
import os
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from anticaptchaofficial.imagecaptcha import *
import auxiliares as aux
from bs4 import BeautifulSoup
import messagebox
import sensiveis as senhas
from subprocess import CREATE_NO_WINDOW
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
import time


class TratarSite:
    """
    Classe que armazena todas as rotinas de execução de ações no controle remoto de site através do Python
    """

    def __init__(self, url, nomeperfil, caminhodownload=aux.caminhoprojeto('Downloads')):
        self.url = url
        self.perfil = nomeperfil
        self.navegador = None
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
            self.navegador.get(self.url)
            time.sleep(1)
            # Testa se a página carregou (ainda tem que fazer um teste e condição quando ele apresenta um texto de erro de carregamento)
            # ==========================================================================================================================
            resultadolimpo = ''
            corposite = BeautifulSoup(self.navegador.page_source, 'html.parser')
            for string in corposite.strings:
                resultadolimpo = resultadolimpo + ' ' + string

            if len(resultadolimpo) == 0:
                messagebox.msgbox(f'Site com problema de carregamento!', messagebox.MB_OK, 'Site fora do ar')
                self.navegador = -1
            # =========================================================================================================================
            time.sleep(1)
            self.trataralerta()
            time.sleep(1)
            return self.navegador

    def configuraprofilechrome(self, ableprintpreview=True, openpdf=True):
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

        self.options = webdriver.ChromeOptions()

        if aux.caminhoprojeto('Profile') != '':
            # self.options.add_argument(f"user-data-dir={aux.caminhoprojeto('Downloads')+'\\'}" )
            self.options.add_argument("--start-maximized")
            self.options.add_argument("--disable-features=ChromeWhatsNewUI")
            # self.options.add_argument("--print-to-pdf="+aux.caminhoprojeto('Downloads'))
            self.options.add_experimental_option('prefs', prefs)
            if ableprintpreview:
                self.options.add_argument('--kiosk-printing')
            else:
                self.options.add_argument("---printing")
                self.options.add_argument("--disable-print-preview")

            self.options.add_argument("--silent")

            # Forma invisível
            # self.options.add_argument("--headless")

        chrome_service = Service(ChromeDriverManager().install())
        chrome_service.creationflags = CREATE_NO_WINDOW

        return webdriver.Chrome(options=self.options, service=chrome_service)

    def verificarobjetoexiste(self, identificador, endereco, valorselecao='', itemunico=True, iraoobjeto=False, sotestar=False):
        """
        :param sotestar: retornar se o objeto existe ou não (retorno se torna um booleano).
        :param iraoobjeto: se simula o mouse em cima do objeto ou não.
        :param identificador: como será identificado, por nome, por nome de classe, etc.
        :param endereco: nome do objeto no site (lembrando que o nome é segundo o parâmetro anterior, se for definido ID no parâmetro anterior,
                         nesse tem que vir o ID do objeto do site, por exemplo.
        :param valorselecao: caso se um combobox passar o valor de seleção desejado que ele mesmo seleciona o dado nesse parâmetro.
        :param itemunico: caso seja uma coleção de objetos que queira selecionar (todos o do nome de classe x), colocar False, padrão será item único.
        :return: vai retornar o objeto do site para ser trabalhado já verificando se o mesmo existe, caso não encontre retorna None
        """

        from selenium.webdriver.support.ui import Select

        if self.navegador is not None:
            try:
                if itemunico:
                    if len(valorselecao) == 0:
                        elemento = WebDriverWait(self.navegador, self.delay).until(EC.visibility_of_element_located((getattr(By, identificador), endereco)))
                        # elemento = self.navegador.find_element(getattr(By, identificador), endereco)

                    else:
                        elemento = WebDriverWait(self.navegador, self.delay).until(EC.visibility_of_element_located((getattr(By, identificador), endereco)))
                        # elemento = Select(self.navegador.find_element(getattr(By, identificador), endereco))
                        select = Select(elemento)
                        select.select_by_visible_text(valorselecao)
                        # elemento.select_by_value(valorselecao)
                else:
                    if len(valorselecao) == 0:
                        elemento = self.navegador.find_elements(getattr(By, identificador), endereco)
                    else:
                        elemento = Select(self.navegador.find_elements(getattr(By, identificador), endereco))
                        select = Select(elemento)
                        select.select_by_visible_text(valorselecao)
                        # elemento.select_by_value(valorselecao)

                if iraoobjeto and elemento is not None:
                    self.navegador.execute_script("arguments[0].click()", elemento)
                    time.sleep(1)
                    self.trataralerta()
                    time.sleep(1)

                if not sotestar:
                    return elemento
                else:
                    return True

            except NoSuchElementException:
                if not sotestar:
                    return None
                else:
                    return False

            except TimeoutException:
                # messagebox.msgbox('Erro de carregamento objeto!', messagebox.MB_OK, 'Erro Carregamento')
                if not sotestar:
                    return None
                else:
                    return False

    def descerrolagem(self):
        """
        Desce a barra de rolagem do navegador
        """
        self.navegador.execute_script("window.scrollTo(0,document.body.scrollHeight)")
        time.sleep(1)

    def mexerzoom(self, valor):
        """
        Mexe no zoom da página
        """
        self.navegador.execute_script("document.body.style.transform='scale(" + str(valor) + ")';")
        time.sleep(1)

    def baixarimagem(self, identificador, endereco, caminho):
        """
        :param identificador: como será identificado, por nome, por nome de classe, etc.
        :param endereco: nome do objeto no site (lembrando que o nome é segundo o parâmetro anterior, se for definido ID no parâmetro anterior,
                         nesse tem que vir o ID do objeto do site, por exemplo.
        :param caminho: caminho onde a imagem será salva.
        :return: informa se achou a imagem no site e se conseguiu salvar a mesma.
        """
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
                    # Verifica se a imagem veio toda "preta" (quando o site não carrega o CAPTCHA)
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
        """
        Wait for downloads to finish with a specified timeout.
        Args
        ----
        caminho : str
            The path to the folder where the files will be downloaded.
        timeout : int
            How many seconds to wait until timing out.
        nfiles : int, defaults to None
            If provided, also wait for the expected number of files.
        """
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
        """
        :return: a quantidade de abas abertas no navegador
        """
        if self.navegador is not None:
            return len(self.navegador.window_handles)

    def irparaaba(self, indice=-1, titulo=''):
        """

        :param titulo: título da aba
        :param indice: número absoluto da aba desejada
        :return:
        """
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
        """
        fecha o browser carregado no objeto
        """
        if self.navegador is not None and hasattr(self.navegador, 'quit'):
            self.navegador.quit()



    def resolvercaptcha(self, identificacaocaixa, caixacaptcha, identicacaobotao, botao):
        """
        :param identificacaocaixa: opção de como a caixa de texto do captcha será identificada (ID, NAME, CLASS, ETC.)
        :param caixacaptcha: idenficação da caixa do captcha segundo a variável anterior (se for ID, colocar o nome ID, por exemplo)
        :param identicacaobotao: opção de como o botão de ação do captcha será identificado (ID, NAME, CLASS, ETC.)
        :param botao: idenficação do botão de ação do captcha segundo a variável anterior (se for ID, colocar o nome ID, por exemplo)
        :return: booleana dizendo se conseguiu ou não resolver o captcha
        """
        from selenium.common.exceptions import TimeoutException

        resposta = False
        textoerro = ''
        solver = imagecaptcha()
        solver.set_verbose(1)
        solver.set_key(senhas.chaveanticaptcha)

        captcha_text = solver.solve_and_return_solution(self.caminho)

        if len(str(captcha_text)) != 0:
            captcha = self.verificarobjetoexiste(identificacaocaixa, caixacaptcha)
            captcha.send_keys(captcha_text)
            self.verificarobjetoexiste(identicacaobotao, botao, iraoobjeto=True)

            try:
                mensagemerro = self.retornartabela(3)
                if mensagemerro == 'Código de Segurança inválido. Favor retornar.':
                    solver.report_incorrect_image_captcha()
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
                # 4170
                table = pagina.find('table', {'border': 0, 'cellpadding': 12, 'cellspacing': 1})
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
                            intermediario = intermediario.replace('Emissão de 2a Via da Taxa e Consulta a Débitos Anteriores', '')
                            intermediario = intermediario.strip()

        if tipolista != 3:
            if len(linha) > 0 and len(cabecalhos) > 0:
                tabela.append(dict(zip(cabecalhos, linha)))
        else:
            tabela = intermediario

        return tabela

    def pegaarquivobaixado(self, timeout, quantabas=0, caminhobaixado=''):
        if quantabas > 0:
            while self.num_abas() > quantabas:
                time.sleep(1)
        quantabas = self.num_abas()
        self.navegador.execute_script('window.open()')
        time.sleep(1)
        if quantabas < self.num_abas():
            # switch to new tab
            self.navegador.switch_to.window(self.navegador.window_handles[-1])
            # navigate to chrome downloads
            self.navegador.get('chrome://downloads')
            time.sleep(1)
            # define the endTime
            endTime = time.time() + timeout
            while True:
                try:
                    # get downloaded percentage
                    downloadPercentage = self.navegador.execute_script(
                        "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('#progress').value")
                    # check if downloadPercentage is 100 (otherwise the script will keep waiting)
                    if downloadPercentage == 100:
                        time.sleep(1)
                        # return the file name once the download is completed
                        return self.navegador.execute_script(
                            "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('div#content  #file-link').text")
                    time.sleep(1)

                except BaseException as err:
                    try:
                        time.sleep(1)
                        arquivo = self.navegador.execute_script(
                            "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('div#content  #file-link').text")
                        if len(arquivo) > 0:
                            return arquivo
                        else:
                            pass

                    except Exception as e:
                        # Obtendo o caminho da pasta de downloads padrão do Chrome
                        if len(caminhobaixado) > 0:
                            # Espera o download finalizar
                            self.esperadownloads(caminhobaixado, timeout, 1)
                            # Pega o último arquivo baixado da pasta
                            arquivo = aux.ultimoarquivo(caminhobaixado, 'pdf')
                            return os.path.basename(arquivo)

                finally:
                    time.sleep(1)
                    if self.num_abas() > 1:
                        if self.irparaaba(titulo='Downloads'):
                            if self.navegador.current_url == 'chrome://downloads/':
                                self.fecharaba()

                time.sleep(1)
                if time.time() > endTime:
                    break
