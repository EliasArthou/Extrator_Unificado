# from selenium import webdriver
# from selenium.webdriver.chrome.service import Service
# from webdriver_manager.chrome import ChromeDriverManager
# from fuzzywuzzy import fuzz
# from a_selenium2df import get_df
# from undetected_chromedriver import Chrome, ChromeOptions
# import io
# from PIL import Image
# import os
# from selenium.common.exceptions import NoSuchElementException
# from selenium.webdriver.common.by import By
# from anticaptchaofficial.imagecaptcha import *
# from anticaptchaofficial.recaptchav2proxyless import *
# import auxiliares as aux
# from bs4 import BeautifulSoup
# import messagebox
# import sensiveis as senhas
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.common.exceptions import TimeoutException
# from selenium.webdriver.support.ui import Select
# import time
# import pandas as pd
# from a_selenium_iframes_crawler import Iframes
#
#
# class TratarSite:
#     """
#     Classe que armazena todas as rotinas de execução de ações no controle remoto de site através do Python
#     """
#
#     def __init__(self, url, nomeperfil, caminhodownload=aux.caminhoprojeto('Downloads')):
#         self.url = url
#         self.perfil = nomeperfil
#         self.navegador = None
#         self.dataframepagina = None
#         self.options = None
#         self.delay = 10
#         self.caminho = ''
#         self.caminhodownload = caminhodownload
#
#     def abrirnavegador(self, habilitarpdf=False):
#         """
#         :return: navegador configurado com o site desejado aberto
#         """
#         if self.navegador is not None:
#             self.fecharsite()
#
#         self.navegador = self.configuraprofilechrome(openpdf=habilitarpdf)
#         if self.navegador is not None:
#             # Changing the property of the navigator value for webdriver to undefined
#             self.navegador.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
#             self.navegador.get(self.url)
#             time.sleep(1)
#             # Testa se a página carregou (ainda tem que fazer um teste e condição quando ele apresenta um texto de erro de carregamento)
#             # ==========================================================================================================================
#             resultadolimpo = ''
#             corposite = BeautifulSoup(self.navegador.page_source, 'html.parser')
#             for string in corposite.strings:
#                 resultadolimpo = resultadolimpo + ' ' + string
#
#             if len(resultadolimpo) == 0:
#                 messagebox.msgbox(f'Site com problema de carregamento!', messagebox.MB_OK, 'Site fora do ar')
#                 self.navegador = -1
#             # =========================================================================================================================
#             time.sleep(1)
#             self.trataralerta()
#             time.sleep(1)
#             return self.navegador
#
#     def configuraprofilechrome(self, ableprintpreview=True, openpdf=True, navegarincognito=False):
#         """
#         Configura usuário e opções no navegador aberto para execução
#         return: o navegador configurado para iniciar a execução das rotinas
#         """
#
#         prefs = {
#             'profile.name': self.perfil,
#             'download.default_directory': self.caminhodownload,  # Change default directory for downloads
#             'download.directory_upgrade': True,
#             'download.prompt_for_download': False,  # To auto download the file
#             'plugins.always_open_pdf_externally': not openpdf,  # It will not show PDF directly in chrome
#             "credentials_enable_service": False,  # service to save password
#             "profile.password_manager_enabled": False,  # turn off password manager
#             'printing.print_preview_sticky_settings.appState': json.dumps({
#                 'recentDestinations': [{
#                     'id': 'Save as PDF',
#                     'origin': 'local',
#                     'account': '',
#                 }],
#                 'selectedDestinationId': 'Save as PDF',
#                 'version': 2,
#                 'defaultDestinationId': 'Save as PDF',
#                 'destinationSelection': [{
#                     'id': 'Save as PDF',
#                     'rules': {
#                         'fileExtension': '.pdf',
#                         'isDefault': True,
#                         'saveToFolder': f"{self.caminhodownload}"
#                     }
#                 }]
#             }
#             ),
#             'savefile.default_directory': self.caminhodownload
#         }
#
#         self.options = ChromeOptions()
#
#         if aux.caminhoprojeto('Profile') != '':
#             self.options.add_argument("--start-maximized")
#             self.options.add_experimental_option('prefs', prefs)
#             self.options.add_argument("--disable-popup-blocking")
#             self.options.add_argument("--disable-infobars")
#             self.options.add_argument("--disable-notifications")
#             self.options.add_argument('--disable-extensions')
#             self.options.add_argument('--disable-plugins')
#
#             if ableprintpreview:
#                 self.options.add_argument('--kiosk-printing')
#             else:
#                 self.options.add_argument("---printing")
#                 self.options.add_argument("--disable-print-preview")
#
#             self.options.add_argument("--silent")
#
#             # Forma invisível
#             # self.options.add_argument("--headless")
#
#         if navegarincognito:
#             driver = Chrome(options=self.options)
#             driver.maximize_window()
#         else:
#             self.options.add_argument("--disable-features=ChromeWhatsNewUI")
#             # Adding argument to disable the AutomationControlled flag
#             self.options.add_argument("--disable-blink-features=AutomationControlled")
#             self.options.add_experimental_option('excludeSwitches', ["enable-automation"])
#             self.options.add_experimental_option('useAutomationExtension', False)
#             chrome_service = Service(ChromeDriverManager().install())
#             driver = webdriver.Chrome(options=self.options, service=chrome_service)
#
#         return driver
#
#     def verificarobjetoexiste(self, identificador, endereco, valorselecao='', itemunico=True, iraoobjeto=False,
#                               sotestar=False, buscar_em_iframes=False):
#         """
#          :param sotestar: retornar se o objeto existe ou não (retorno se torna um booleano).
#          :param iraoobjeto: se simula o mouse em cima do objeto ou não.
#          :param identificador: como será identificado, por nome, por nome de classe, etc.
#          :param endereco: nome do objeto no site (lembrando que o nome é segundo o parâmetro anterior, se for definido ID no parâmetro anterior,
#                           nesse tem que vir o ID do objeto do site, por exemplo.
#          :param valorselecao: caso se um combobox passar o valor de seleção desejado que ele mesmo seleciona o dado nesse parâmetro.
#          :param itemunico: caso seja uma coleção de objetos que queira selecionar (todos o do nome de classe x), colocar False, padrão será item único.
#          :param buscar_em_iframes: buscar o elemento em todos os iframes se necessário.
#          :return: vai retornar o objeto do site para ser trabalhado já verificando se o mesmo existe, caso não encontre retorna None
#         """
#
#         if self.navegador is not None:
#             original_window = self.navegador.current_window_handle
#             try:
#                 return self.buscar_elemento(identificador, endereco, valorselecao, itemunico, iraoobjeto, sotestar)
#             except (NoSuchElementException, TimeoutException):
#                 if buscar_em_iframes:
#
#                     # iframes = self.verificarobjetoexiste('TAG_NAME', 'iframe', itemunico=False)
#                     iframes = Iframes(driver=self.navegador, By=By, WebDriverWait=WebDriverWait, expected_conditions=EC)
#                     iframes = iframes.iframes.items()
#                     for iframe in iframes:
#                         try:
#                             self.navegador.switch_to.frame(iframe)
#                             elemento = self.buscar_elemento(identificador, endereco, valorselecao, itemunico,
#                                                             iraoobjeto, sotestar)
#                             if elemento:
#                                 return elemento
#                         except (NoSuchElementException, TimeoutException):
#                             pass
#                         self.navegador.switch_to.window(original_window)
#                 return False if sotestar else None
#
#     def buscar_elemento(self, identificador, endereco, valorselecao, itemunico, iraoobjeto, sotestar):
#         """
#         Função auxiliar para buscar o elemento.
#         """
#         if itemunico:
#             if len(valorselecao) == 0:
#                 elemento = WebDriverWait(self.navegador, self.delay).until(
#                     EC.visibility_of_element_located((getattr(By, identificador), endereco)))
#             else:
#                 elemento = WebDriverWait(self.navegador, self.delay).until(
#                     EC.visibility_of_element_located((getattr(By, identificador), endereco)))
#                 select = Select(elemento)
#                 select.select_by_visible_text(valorselecao)
#         else:
#             if len(valorselecao) == 0:
#                 elemento = self.navegador.find_elements(getattr(By, identificador), endereco)
#             else:
#                 elemento = Select(self.navegador.find_elements(getattr(By, identificador), endereco))
#                 select = Select(elemento)
#                 select.select_by_visible_text(valorselecao)
#
#         if iraoobjeto and elemento is not None:
#             self.navegador.execute_script("arguments[0].click()", elemento)
#             time.sleep(1)
#             self.trataralerta()
#             time.sleep(1)
#
#         return True if sotestar else elemento
#
#     def descerrolagem(self):
#         """
#         Desce a barra de rolagem do navegador
#         """
#         self.navegador.execute_script("window.scrollTo(0,document.body.scrollHeight)")
#         time.sleep(1)
#
#     def mexerzoom(self, valor):
#         """
#         Mexe no zoom da página
#         """
#         self.navegador.execute_script("document.body.style.transform='scale(" + str(valor) + ")';")
#         time.sleep(1)
#
#     def baixarimagem(self, identificador, endereco, caminho):
#         """
#         :param identificador: como será identificado, por nome, por nome de classe, etc.
#         :param endereco: nome do objeto no site (lembrando que o nome é segundo o parâmetro anterior, se for definido ID no parâmetro anterior,
#                          nesse tem que vir o ID do objeto do site, por exemplo.
#         :param caminho: caminho onde a imagem será salva.
#         :return: informa se achou a imagem no site e se conseguiu salvar a mesma.
#         """
#         achouimagem = False
#         salvouimagem = False
#         if self.navegador is not None:
#             if os.path.isfile(caminho):
#                 os.remove(caminho)
#             elemento = self.verificarobjetoexiste(identificador, endereco)
#             if elemento is not None:
#                 achouimagem = True
#                 image = elemento.screenshot_as_png
#                 imagestream = io.BytesIO(image)
#                 im = Image.open(imagestream)
#                 im.save(caminho)
#                 if os.path.isfile(caminho):
#                     # Verifica se a imagem veio toda "preta" (quando o site não carrega o CAPTCHA)
#                     if aux.quantidade_cores(caminho) > 1:
#                         salvouimagem = True
#                         self.caminho = caminho
#                     else:
#                         os.remove(caminho)
#                         salvouimagem = False
#                 else:
#                     salvouimagem = False
#
#         return achouimagem, salvouimagem
#
#     @staticmethod
#     def esperadownloads(caminho, timeout, nfiles=None):
#         """
#         Wait for downloads to finish with a specified timeout.
#         Args
#         ----
#         caminho : str
#             The path to the folder where the files will be downloaded.
#         timeout : int
#             How many seconds to wait until timing out.
#         nfiles : int, defaults to None
#             If provided, also wait for the expected number of files.
#         """
#         time.sleep(1)
#         seconds = 0
#         dl_wait = True
#         while dl_wait and seconds < timeout:
#             time.sleep(1)
#             dl_wait = False
#             files = os.listdir(caminho)
#             if nfiles and len(files) != nfiles:
#                 dl_wait = True
#
#             for fname in files:
#                 if fname.endswith('.crdownload'):
#                     dl_wait = True
#
#             seconds += 1
#
#         time.sleep(1)
#         return seconds
#
#     def num_abas(self):
#         """
#         :return: a quantidade de abas abertas no navegador
#         """
#         if self.navegador is not None:
#             return len(self.navegador.window_handles)
#
#     def irparaaba(self, indice=-1, titulo=''):
#         """
#
#         :param titulo: título da aba
#         :param indice: número absoluto da aba desejada
#         :return:
#         """
#         achouaba = False
#
#         if indice >= 0 or len(titulo) > 0:
#             if indice >= 0:
#                 if indice <= self.num_abas():
#                     self.navegador.switch_to.window(self.navegador.window_handles[indice - 1])
#             else:
#                 janelaatual = self.navegador.current_window_handle
#                 handles = self.navegador.window_handles
#                 for hdl in handles:
#                     self.navegador.switch_to.window(hdl)
#                     if titulo.upper() in self.navegador.title.upper() + self.navegador.current_url.upper():
#                         achouaba = True
#                         break
#                 if not achouaba:
#                     self.navegador.switch_to.window(janelaatual)
#
#                 return achouaba
#
#     def fecharaba(self, indice=0):
#         if indice == 0:
#             self.navegador.execute_script('window.open("","_self").close()')
#         else:
#             self.irparaaba(indice)
#             self.navegador.execute_script('window.open("","_self").close()')
#
#     def irparaframe(self, frame):
#         self.navegador.switch_to.frame(frame)
#
#     def sairdoframe(self):
#         self.navegador.switch_to.default_content()
#
#     def trataralerta(self):
#         from selenium.webdriver.common.alert import Alert
#         from selenium.common.exceptions import NoAlertPresentException
#
#         try:
#             alerta = Alert(self.navegador)
#             if alerta is not None:
#                 textoalerta = alerta.text
#                 alerta.accept()
#                 time.sleep(1)
#             else:
#                 textoalerta = ''
#
#         except NoAlertPresentException:
#             textoalerta = ''
#             pass
#
#         return textoalerta
#
#     def fecharsite(self):
#         """
#         fecha o browser carregado no objeto
#         """
#         if self.navegador is not None and hasattr(self.navegador, 'quit'):
#             self.navegador.quit()
#
#     def resolvercaptcha(self, identificacaocaixa, caixacaptcha, identicacaobotao, botao):
#         """
#         :param identificacaocaixa: opção de como a caixa de texto do captcha será identificada (ID, NAME, CLASS, ETC.)
#         :param caixacaptcha: idenficação da caixa do captcha segundo a variável anterior (se for ID, colocar o nome ID, por exemplo)
#         :param identicacaobotao: opção de como o botão de ação do captcha será identificado (ID, NAME, CLASS, ETC.)
#         :param botao: idenficação do botão de ação do captcha segundo a variável anterior (se for ID, colocar o nome ID, por exemplo)
#         :return: booleana dizendo se conseguiu ou não resolver o captcha
#         """
#         from selenium.common.exceptions import TimeoutException
#
#         resposta = False
#         textoerro = ''
#         solver = imagecaptcha()
#         solver.set_verbose(1)
#         solver.set_key(senhas.chaveanticaptcha)
#
#         captcha_text = solver.solve_and_return_solution(self.caminho)
#
#         if len(str(captcha_text)) != 0:
#             captcha = self.verificarobjetoexiste(identificacaocaixa, caixacaptcha, iraoobjeto=True)
#             captcha.send_keys(captcha_text)
#             time.sleep(1)
#             self.verificarobjetoexiste(identicacaobotao, botao, iraoobjeto=True)
#
#             try:
#                 mensagemerro = self.trataralerta()
#                 if mensagemerro == '':
#                     # Use o método find() para encontrar o primeiro objeto com a formatação desejada
#                     pagina = BeautifulSoup(self.navegador.page_source, 'html.parser')
#                     objeto_encontrado = pagina.find(
#                         lambda tag: tag.name == 'font' and tag.get('size') == '2' and tag.get('color') == 'red')
#                     if objeto_encontrado:
#                         # Verifica se o objeto contém texto
#                         if objeto_encontrado.text.strip():
#                             mensagemerro = objeto_encontrado.text.strip()
#
#                 if mensagemerro == 'Código digitado não confere! Favor refazer a consulta!':
#                     solver.report_incorrect_image_captcha()
#                     botao_nova_consulta = self.verificarobjetoexiste('XPATH',
#                                                                      "//input[@type='button' and @name='bt' and @value='Nova Consulta']")
#                     if botao_nova_consulta is not None:
#                         botao_nova_consulta.click()
#                     textoerro = mensagemerro
#                     resposta = False
#                 else:
#                     textoerro = mensagemerro
#                     resposta = True
#
#                 if os.path.isfile(self.caminho):
#                     os.remove(self.caminho)
#                     self.caminho = ''
#
#             except TimeoutException:
#                 resposta = True
#                 if os.path.isfile(self.caminho):
#                     os.remove(self.caminho)
#                     self.caminho = ''
#
#         else:
#             print("Erro solução captcha:" + solver.error_code)
#             resposta = False
#
#         return resposta, textoerro
#
#     def resolvecaptchatipo2(self, chavecaptcha):
#         solver = recaptchaV2Proxyless()
#         solver.set_verbose(1)
#         solver.set_key(senhas.chaveanticaptcha)
#         solver.set_website_url(self.navegador.current_url)
#         solver.set_website_key(chavecaptcha)
#         # set optional custom parameter which Google made for their search page Recaptcha v2
#         # solver.set_data_s('"data-s" token from Google Search results "protection"')
#
#         # Specify softId to earn 10% commission with your app.
#         # Get your softId here: https://anti-captcha.com/clients/tools/devcenter
#         solver.set_soft_id(0)
#
#         g_response = solver.solve_and_return_solution()
#         if g_response != 0:
#             return g_response
#         else:
#             print("task finished with error " + solver.error_code)
#             return None
#
#     def retornartabela(self, tipolista):
#         import re
#         import unicodedata
#
#         linha = []
#         cabecalhos = []
#         tabela = []
#         intermediario = ''
#
#         pagina = BeautifulSoup(self.navegador.page_source, 'html.parser')
#
#         match tipolista:
#             case 1:
#                 table = pagina.find('table', {'border': 0, 'width': 500})
#                 for element in table.find_all('td'):
#                     if len(element.text.strip()) > 0:
#                         if aux.right(element.text, 1) == ':':
#                             intermediario = aux.to_raw(element.text)
#                             cabecalhos.append(intermediario)
#                         else:
#                             elementotratado = BeautifulSoup(str(element), 'html.parser')
#                             for br in elementotratado.find_all("br"):
#                                 br.replace_with("\n")
#                             intermediario = unicodedata.normalize("NFKD", aux.to_raw(elementotratado.text))
#                             linha.append(intermediario)
#
#             case 2:
#                 for element in pagina.find_all('input', {'name': re.compile('^taxa')}):
#                     if element['name'] not in cabecalhos:
#                         if len(element['name']) > 0 and len(element['value']) > 0:
#                             linha.append(element['value'])
#                             cabecalhos.append(element['name'])
#                     else:
#                         tabela.append(dict(zip(cabecalhos, linha)))
#                         cabecalhos = [element['name']]
#                         linha = [element['value']]
#
#             case 3:
#                 # 4170
#                 table = pagina.find('table')
#                 itenstabela = table.find_all('td')
#                 if len(itenstabela) == 1:
#                     for element in itenstabela:
#                         if len(element.text.strip()) > 0:
#                             elementotratado = BeautifulSoup(str(element), 'html.parser')
#                             botao = elementotratado.find('button')
#                             if botao is not None:
#                                 botao.decompose()
#                             else:
#                                 botao = elementotratado.find('a', {'href': 'javascript: history.back()'})
#                                 if botao is not None:
#                                     botao.decompose()
#
#                             for br in elementotratado.find_all("br"):
#                                 br.replace_with("\n")
#                             intermediario = unicodedata.normalize("NFKD", aux.to_raw(elementotratado.text))
#                             intermediario = re.sub('\n+', '\n', intermediario)
#                             intermediario = intermediario.replace(
#                                 'Emissão de 2a Via da Taxa e Consulta a Débitos Anteriores', '')
#                             intermediario = intermediario.strip()
#
#         if tipolista != 3:
#             if len(linha) > 0 and len(cabecalhos) > 0:
#                 tabela.append(dict(zip(cabecalhos, linha)))
#         else:
#             tabela = intermediario
#
#         return tabela
#
#     @staticmethod
#     def obter_arquivos_atuais(diretorio):
#         """Obtém uma lista dos arquivos no diretório especificado."""
#         return set(os.listdir(diretorio))
#
#     @staticmethod
#     def esperar_novo_download(diretorio, arquivos_originais, timeout=300):
#         """Espera até que um novo download seja concluído no diretório especificado."""
#         tempo_inicio = time.time()
#         while True:
#             arquivos_atuais = set(os.listdir(diretorio))
#             arquivos_novos = arquivos_atuais - arquivos_originais
#             for arquivo in arquivos_novos:
#                 if not arquivo.endswith('.crdownload'):
#                     print(f"Download completo: {arquivo}")
#                     return arquivo
#             if (time.time() - tempo_inicio) > timeout:
#                 print("Tempo de espera excedido.")
#                 return None
#             time.sleep(1)
#
#     def pegaarquivobaixado(self, timeout, quantabas=0, caminhobaixado=''):
#         if quantabas > 0:
#             while self.num_abas() > quantabas:
#                 time.sleep(1)
#         quantabas = self.num_abas()
#         self.navegador.execute_script('window.open()')
#         time.sleep(2)
#         if quantabas < self.num_abas():
#             # switch to new tab
#             self.navegador.switch_to.window(self.navegador.window_handles[-1])
#             # navigate to chrome downloads
#             self.navegador.get('chrome://downloads')
#             time.sleep(1)
#             # define the endTime
#             endTime = time.time() + timeout
#
#             # Aguarda até que algum download seja iniciado
#             max_attempts = 30  # Número máximo de tentativas
#             attempts = 0
#             download_started = False
#
#             while attempts < max_attempts and not download_started:
#                 # Verifica se há algum download listado
#                 download_items = self.navegador.execute_script(
#                     "return document.querySelector('downloads-manager').shadowRoot.querySelector('iron-list').items;")
#                 if download_items:
#                     download_started = True
#                 else:
#                     time.sleep(1)  # Espera um segundo antes de verificar novamente
#                     attempts += 1
#
#             if download_started:
#                 while True:
#                     try:
#                         # get downloaded percentage
#                         downloadPercentage = self.navegador.execute_script(
#                             "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('#progress').value")
#                         # check if downloadPercentage is 100 (otherwise the script will keep waiting)
#                         if downloadPercentage == 100:
#                             time.sleep(1)
#                             # return the file name once the download is completed
#                             return self.navegador.execute_script(
#                                 "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('div#content  #file-link').text")
#                         time.sleep(1)
#
#                     except BaseException as err:
#                         try:
#                             time.sleep(1)
#                             arquivo = self.navegador.execute_script(
#                                 "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('div#content  #file-link').text")
#                             if len(arquivo) > 0:
#                                 return arquivo
#                             else:
#                                 pass
#
#                         except Exception as e:
#                             # Obtendo o caminho da pasta de downloads padrão do Chrome
#                             if len(caminhobaixado) > 0:
#                                 # Espera o download finalizar
#                                 self.esperadownloads(caminhobaixado, timeout, 1)
#                                 # Pega o último arquivo baixado da pasta
#                                 arquivo = aux.ultimoarquivo(caminhobaixado, 'pdf')
#                                 return os.path.basename(arquivo)
#
#                     finally:
#                         time.sleep(1)
#                         if self.num_abas() > 1:
#                             if self.irparaaba(titulo='Downloads'):
#                                 if self.navegador.current_url == 'chrome://downloads/':
#                                     self.fecharaba()
#
#                     time.sleep(1)
#                     if time.time() > endTime:
#                         break
#
#     def retornarpaginaemdf(self):
#         self.dataframepagina = get_df(self.navegador, By, WebDriverWait, EC, queryselector='*', with_methods=True)
#
#     def buscarobjetoemdf(self, filtros: dict, atributo: str = None):
#         """
#         Busca objetos em um dataframe com base em filtros e retorna o resultado.
#         Args:
#             filtros (dict): Dicionário de filtros com chave-valor.
#             atributo (str, optional): Atributo específico a ser retornado. Defaults to None.
#         Returns:
#             pd.DataFrame | list | str | None: Resultado da busca.
#         """
#         if self.dataframepagina is None:
#             raise ValueError("Dataframe não iniciado!")
#
#         df = self.dataframepagina
#         query_conditions = []
#
#         for key, value in filtros.items():
#             identificador = f"aa_{key}" if "_" not in key else key
#             match value:
#                 case '^NaN':
#                     condition = f"({identificador}.notna())"
#                 case value if pd.isna(value):
#                     condition = f"({identificador}.isna())"
#                 case _:
#                     if df[identificador].dtype == 'string' or df[identificador].dtype == 'object':
#                         # query_conditions.append(f"({identificador} == '{value}')")
#                         condition = f"({identificador} == '{value}')"
#                     else:
#                         # query_conditions.append(f"({identificador} == {value})")
#                         condition = f"({identificador} == {value})"
#
#             query_conditions.append(condition)
#
#         if query_conditions:
#             query_string = " & ".join(query_conditions)
#             dffiltrado = df.query(query_string)
#
#             if atributo:
#                 if "_" not in atributo:
#                     atributo = f"aa_{atributo}"
#                 if atributo in dffiltrado.columns:
#                     if len(dffiltrado) == 1:
#                         return dffiltrado[atributo].iloc[0]
#                     else:
#                         return dffiltrado[atributo].tolist()
#                 else:
#                     raise ValueError(f"Atributo '{atributo}' não encontrado no DataFrame.")
#             else:
#                 return dffiltrado
#         else:
#             raise ValueError("Nenhum filtro fornecido!")
#
#     def encontrar_input_oculto(self, identificador, endereco, buscar_em_iframes=False):
#         """
#         Encontra um elemento input oculto e retorna o objeto WebElement.
#         :param identificador: str - Método de identificação do elemento (ID, NAME, CLASS_NAME, TAG_NAME, etc.).
#         :param endereco: str - Valor do identificador (por exemplo, o ID real do elemento).
#         :param buscar_em_iframes: bool - Se deve buscar o elemento em iframes.
#         :return: WebElement ou None - O objeto WebElement do elemento input oculto, ou None se não for encontrado.
#         """
#         try:
#             # Tenta encontrar o elemento na página principal primeiro
#             hidden_input = self.navegador.find_element(getattr(By, identificador), endereco)
#             print(hidden_input.get_attribute('value'))
#             return hidden_input
#         except NoSuchElementException:
#             # Se não encontrar na página principal e buscar_em_iframes for True, busca nos iframes
#             if buscar_em_iframes:
#                 iframes = self.navegador.find_elements(By.TAG_NAME, 'iframe')
#                 original_window = self.navegador.current_window_handle
#                 for iframe in iframes:
#                     try:
#                         self.navegador.switch_to.frame(iframe)
#                         hidden_input = self.navegador.find_element(getattr(By, identificador), endereco)
#                         print(hidden_input.get_attribute('value'))
#                         return hidden_input
#                     except NoSuchElementException:
#                         continue
#                     finally:
#                         self.navegador.switch_to.window(
#                             original_window)  # Garante que volta para o conteúdo principal após cada tentativa
#         return None  # Retorna None se o elemento não for encontrado em lugar nenhum
#
#     @staticmethod
#     def buscar_por_proximidade(lista, texto_de_busca, prop=None, corte_perc=1):
#         texto_mais_parecido = ''
#         percentual_de_compatibilidade = 0
#         item_lista = None
#
#         for elemento in lista:
#             if prop is None:
#                 texto_elemento = elemento
#             else:
#                 texto_elemento = getattr(elemento, prop)
#
#             compatibilidade = fuzz.ratio(texto_de_busca, texto_elemento)
#
#             if compatibilidade > percentual_de_compatibilidade and compatibilidade >= corte_perc:
#                 texto_mais_parecido = texto_elemento
#                 percentual_de_compatibilidade = compatibilidade
#                 item_lista = elemento
#
#         return texto_mais_parecido, percentual_de_compatibilidade, item_lista

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
from selenium.common.exceptions import NoSuchElementException, TimeoutException
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
            self.options.add_argument("--disable-features=ChromeWhatsNewUI")
            self.options.add_argument("--disable-blink-features=AutomationControlled")
            self.options.add_experimental_option('excludeSwitches', ["enable-automation"])
            self.options.add_experimental_option('useAutomationExtension', False)
            chrome_service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(options=self.options, service=chrome_service)

        return driver

    def verificarobjetoexiste(self, identificador, endereco, valorselecao='', itemunico=True, iraoobjeto=False,
                              sotestar=False, buscar_em_iframes=False):
        if self.navegador is not None:
            original_window = self.navegador.current_window_handle
            try:
                # print(f"Procurando por elemento com {identificador} = {endereco}")
                elemento = self.buscar_elemento(identificador, endereco, valorselecao, itemunico, iraoobjeto, sotestar)
                if elemento:
                    # print("Elemento encontrado.")
                    return elemento if not sotestar else True
                else:
                    # print("Elemento não encontrado.")
                    return [] if itemunico is False else (False if sotestar else None)
            except (NoSuchElementException, TimeoutException) as e:
                # print(f"Erro ao procurar o elemento: {endereco} com o identificador {identificador}")
                if buscar_em_iframes:
                    iframes = Iframes(driver=self.navegador, By=By, WebDriverWait=WebDriverWait, expected_conditions=EC)
                    iframes = iframes.iframes.items()
                    for key, iframe_path in iframes:
                        try:
                            for iframe in iframe_path:
                                self.navegador.switch_to.frame(iframe)
                            elemento = self.buscar_elemento(identificador, endereco, valorselecao, itemunico,
                                                            iraoobjeto, sotestar)
                            if elemento:
                                # print("Elemento encontrado dentro de um iframe.")
                                return elemento if not sotestar else True
                        except (NoSuchElementException, TimeoutException) as e:
                            w = 1
                            # print(f"Erro ao procurar o elemento: {endereco} com o identificador {identificador}")
                            # print(f"Erro ao procurar o elemento dentro do iframe: {e}")
                        finally:
                            self.navegador.switch_to.window(original_window)
                return [] if itemunico is False else (False if sotestar else None)

    def buscar_elemento(self, identificador, endereco, valorselecao, itemunico, iraoobjeto, sotestar):
        """
        Função auxiliar para buscar o elemento.
        """
        if itemunico:
            if len(valorselecao) == 0:
                elemento = WebDriverWait(self.navegador, self.delay).until(
                    EC.visibility_of_element_located((getattr(By, identificador), endereco)))
            else:
                elemento = WebDriverWait(self.navegador, self.delay).until(
                    EC.visibility_of_element_located((getattr(By, identificador), endereco)))
                select = Select(elemento)
                select.select_by_visible_text(valorselecao)
        else:
            if len(valorselecao) == 0:
                elemento = self.navegador.find_elements(getattr(By, identificador), endereco)
            else:
                elemento = Select(self.navegador.find_elements(getattr(By, identificador), endereco))
                select = Select(elemento)
                select.select_by_visible_text(valorselecao)

        if iraoobjeto and elemento is not None:
            self.navegador.execute_script("arguments[0].click()", elemento)
            time.sleep(1)
            self.trataralerta()
            time.sleep(1)

        return True if sotestar else elemento

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
        if quantabas > 0:
            while self.num_abas() > quantabas:
                time.sleep(1)
        quantabas = self.num_abas()
        self.navegador.execute_script('window.open()')
        time.sleep(2)
        if quantabas < self.num_abas():
            self.navegador.switch_to.window(self.navegador.window_handles[-1])
            self.navegador.get('chrome://downloads')
            time.sleep(1)
            endTime = time.time() + timeout

            max_attempts = 30
            attempts = 0
            download_started = False

            while attempts < max_attempts and not download_started:
                download_items = self.navegador.execute_script(
                    "return document.querySelector('downloads-manager').shadowRoot.querySelector('iron-list').items;")
                if download_items:
                    download_started = True
                else:
                    time.sleep(1)
                    attempts += 1

            if download_started:
                while True:
                    try:
                        downloadPercentage = self.navegador.execute_script(
                            "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('#progress').value")
                        if downloadPercentage == 100:
                            time.sleep(1)
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
                            if len(caminhobaixado) > 0:
                                self.esperadownloads(caminhobaixado, timeout, 1)
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

    def buscar_por_proximidade(self, identificador, seletor, texto_de_busca, corte_perc=70):
        elementos = self.navegador.find_elements(getattr(By, identificador), seletor)
        texto_mais_parecido = ''
        percentual_de_compatibilidade = 0
        elemento_encontrado = None

        for elemento in elementos:
            texto_elemento = elemento.text.lower()
            texto_de_busca_lower = texto_de_busca.lower()
            compatibilidade = fuzz.ratio(texto_de_busca_lower, texto_elemento)

            if compatibilidade > percentual_de_compatibilidade:
                texto_mais_parecido = texto_elemento
                percentual_de_compatibilidade = compatibilidade
                elemento_encontrado = elemento

        if elemento_encontrado and percentual_de_compatibilidade >= corte_perc:
            print(
                f"Elemento encontrado: {texto_mais_parecido} com {percentual_de_compatibilidade}% de compatibilidade com o item {texto_de_busca}.")
            return elemento_encontrado
        else:
            print(
                f"Nenhum elemento encontrado com pelo menos {corte_perc}% de compatibilidade com o item {texto_de_busca}.")
            if elemento_encontrado:
                print(
                    f"O elemento mais próximo encontrado foi: {texto_mais_parecido} com {percentual_de_compatibilidade}% de compatibilidade.")
            return None