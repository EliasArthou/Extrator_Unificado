import time
import auxiliares as aux
import sensiveis as info
import web
import re
import os
import traceback
import inspect
from dotenv import load_dotenv

# Carrega as variáveis do arquivo .env
load_dotenv()

identificador = 0
Usuario = 1
Senha = 2
Administradora = 3
Condominio = 4
Apartamento = 5
Loginmultiplo  = 6
Resposta = 7
CheckArquivo = 8
CheckErro = 9
Nomefuncao = 10
ProblemaLogin = 11

timeout = 10
pastadownload = aux.caminhoprojeto('Downloads')
tempoesperadownload = 180
mensagemerropadrao = 'Deu erro! Tentar novamente!'
mensagemsemcondominio = "Condomínio não encontrado!"


def apsa(objeto, linha):
    # XPATH ícone loading : '/html/body/app-root/app-loading-screen/div/div/div[2]'
    # Tela de Loading que fica "invisível" (XPATH): '/html/body/app-root/app-loading-screen'
    site = None
    telaloading = '/html/body/app-root/app-loading-screen'
    listaboletos = []

    try:
        linha[ProblemaLogin] = False
        linha[Nomefuncao] = inspect.currentframe().f_code.co_name
        # Variável que vai retornar a quantidade de boletos
        numboleto = 0
        # Prepara o objeto
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
                              os.getenv('NOMEPROFILECOND'))
        # Abre o browser
        site.abrirnavegador(False)

        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
                                  os.getenv('NOMEPROFILECOND'))
            # Abre o browser
            site.abrirnavegador()
        if site is not None and site.navegador != -1:
            # Pega o delay configurado no objeto "Site"
            delay = site.delay
            # Diminui o delay para achar a pergunta
            site.delay = 2
            # Responde à pergunta de inscrição no site
            seinscrever = site.verificarobjetoexiste('ID', 'onesignal-slidedown-cancel-button')
            if seinscrever is not None:
                site.navegador.execute_script("arguments[0].click()", seinscrever)

            site.delay = delay

            # Campo de usuário
            cmpusuario = site.verificarobjetoexiste('ID', 'login')
            # Verifica se achou o campo de usuário
            if cmpusuario is not None:
                # Limpa o campo de usuário
                cmpusuario.clear()
                # Carrega o campo de usuário
                cmpusuario.send_keys(linha[Usuario])
                # Campo de senha
                cmpsenha = site.verificarobjetoexiste('XPATH',
                                                      '/html/body/app-root/ion-app/ion-router-outlet/app-default-layout/ion-app/ion-router-outlet/app-index/ion-content/div/div[2]/app-form-login/form/div[2]/div[1]/input')
                # Verifica se achou o campo de senha
                if cmpsenha is not None:
                    # Limpa o campo de senha
                    cmpsenha.clear()
                    # Carrega o campo de senha
                    cmpsenha.send_keys(linha[Senha])
                    # Botão de Login
                    botao = site.verificarobjetoexiste('XPATH',
                                                       '/html/body/app-root/ion-app/ion-router-outlet/app-default-layout/ion-app/ion-router-outlet/app-index/ion-content/div/div[2]/app-form-login/form/div[4]/div[2]/button/div')
                    # Verifica se achou o botão de ‘login’
                    if botao is not None:
                        # Clica no botão de ‘login’
                        site.navegador.execute_script("arguments[0].click()", botao)
                        # Verifica site carregando
                        # ====================================================================
                        objtelaloading = site.verificarobjetoexiste('XPATH', telaloading)
                        if objtelaloading is not None:
                            if hasattr(objtelaloading, 'visible'):
                                while objtelaloading.visible:
                                    time.sleep(1)
                        time.sleep(1)
                        # ====================================================================

                        # Mensagem de erro
                        mensagemerro = site.verificarobjetoexiste('XPATH',
                                                                  '/html/body/app-ligthboxes-default[1]/div/div/div[2]/div')
                        if mensagemerro is None:
                            condominios = site.verificarobjetoexiste('CLASS_NAME', 'Item_Label_Name', itemunico=False)
                            if condominios is not None:
                                for condominio in condominios:
                                    if linha[Condominio] in condominio.text:
                                        site.navegador.execute_script("arguments[0].click()", condominio)
                                        # Verifica site carregando
                                        # ====================================================================
                                        objtelaloading = site.verificarobjetoexiste('XPATH', telaloading)
                                        if objtelaloading is not None:
                                            if hasattr(objtelaloading, 'visible'):
                                                while objtelaloading.visible:
                                                    time.sleep(1)
                                        time.sleep(1)
                                        # ====================================================================

                                        # Clica no botão segunda via de boleto
                                        botaoboletos = site.verificarobjetoexiste('XPATH',
                                                                                  '/html/body/app-root/ion-app/ion-router-outlet/app-layout/ion-app/div[2]/section/ion-router-outlet/app-index/div/home-desktop/div/div[2]/div[1]/app-cotas-com-vencimento-no-mes/a/ds-button/button/div')
                                        # Verifica botão de boleto existe
                                        if botaoboletos is not None:
                                            site.navegador.execute_script("arguments[0].click()", botaoboletos)
                                            # Verifica site carregando
                                            # ====================================================================
                                            objtelaloading = site.verificarobjetoexiste('XPATH', telaloading)
                                            if objtelaloading is not None:
                                                if hasattr(objtelaloading, 'visible'):
                                                    while objtelaloading.visible:
                                                        time.sleep(1)
                                            time.sleep(1)
                                            # ====================================================================
                                        # Mensagem objeto sem boleto
                                        semboletos = site.verificarobjetoexiste('CLASS_NAME', 'ListEmpty_Text')
                                        if semboletos is None:
                                            # Áreas de clique da tabela de boleto
                                            iconesboletos = site.verificarobjetoexiste('CLASS_NAME',
                                                                                       'Actions_Option_Label',
                                                                                       itemunico=False)

                                            # "Varre" os itens clicáveis da lista de boletos
                                            for i, boleto in enumerate(iconesboletos, start=1):
                                                numboleto = 0
                                                # Verifica se tem texto
                                                if hasattr(boleto, 'text'):
                                                    # Verifica se é a área de clique da segunda via dos boletos
                                                    if boleto.text == 'Emitir 2ª via':
                                                        # Clica na segunda via
                                                        site.navegador.execute_script("arguments[0].click()", boleto)
                                                        time.sleep(1)
                                                        # Verifica se abriu uma segunda tela com a opção de extração de boleto
                                                        if site.num_abas() > 1:
                                                            # Vai para tela de extração de boleto
                                                            site.irparaaba(2)
                                                            time.sleep(1)
                                                            # Botão de lembrar depois o cadastro de boleto somente online
                                                            lembrardepois = site.verificarobjetoexiste('ID',
                                                                                                       'ctl00_ContentPlaceHolder1_btnLembrarDepois')
                                                            # Verifica se achou o botão
                                                            if lembrardepois is not None:
                                                                # Clica no botão de lembrar depois
                                                                site.navegador.execute_script("arguments[0].click()",
                                                                                              lembrardepois)
                                                                time.sleep(1)

                                                                # Define o novo nome e caminho do arquivo baixado (baseado no nr do boleto)
                                                                caminho_arquivo = site.monitorar_downloads_sem_href(
                                                                    link_elemento=boleto, timeout=120)
                                                                if caminho_arquivo:
                                                                    novonomearquivo = os.path.join(
                                                                        objeto.pastadownload,
                                                                        aux.left(linha[identificador], 4) + (
                                                                            f"_{numboleto}.pdf" if numboleto > 1 else ".pdf")
                                                                    )
                                                                    listaboletos.append(
                                                                        aux.adicionarcabecalhopdf(caminho_arquivo,
                                                                                                  novonomearquivo,
                                                                                                  aux.left(linha[
                                                                                                               identificador],
                                                                                                           4)))
                                                                    numboleto += 1
                                                                    time.sleep(1)
                                                                else:
                                                                    print(f"Falha ao capturar o arquivo baixado para o boleto {i}.\n"
                                                                          f"Função: {inspect.currentframe().f_code.co_name} \n"
                                                                          f"Cliente: {linha[identificador]}")
                                                            time.sleep(1)
                                                            site.irparaaba(2)
                                                            site.fecharaba()
                                                            time.sleep(1)
                                                            site.irparaaba(1)
                                                            time.sleep(1)
                                                            pesquisasatisfacao = site.verificarobjetoexiste('XPATH',
                                                                                                            '/html/body/yes-or-no-with-link[1]/div/div/div[4]/button')
                                                            if pesquisasatisfacao is not None:
                                                                site.navegador.execute_script("arguments[0].click()",
                                                                                              pesquisasatisfacao)
                                                                # Verifica site carregando
                                                                # ====================================================================
                                                                objtelaloading = site.verificarobjetoexiste('XPATH',
                                                                                                            telaloading)
                                                                if objtelaloading is not None:
                                                                    if hasattr(objtelaloading, 'visible'):
                                                                        while objtelaloading.visible:
                                                                            time.sleep(1)
                                                                time.sleep(1)
                                                                # ====================================================================
                                        # Retorna resposta na linha
                                        linha[Resposta] = respostapadrao(numboleto)
                                        break

                        else:
                            if hasattr(mensagemerro, 'text'):
                                linha[Resposta] = mensagemerro.text
                                linha[ProblemaLogin] = True

                        # Fecha o browser
                        site.fecharsite()

        if listaboletos is not None:
            return listaboletos

    except Exception as e:
        linha[Resposta] = str(e)
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = inspect.currentframe().f_code.co_name
        if site is not None:
            site.fecharsite()


def bap(objeto, linha):
    site = None
    listaboletos = []
    linha[Nomefuncao] = inspect.currentframe().f_code.co_name
    try:
        linha[ProblemaLogin] = False
        numboleto = 0
        # Prepara o objeto
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
                              os.getenv('NOMEPROFILECOND'))
        site.delay = 10
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
                                  os.getenv('NOMEPROFILECOND'))
            # Abre o browser
            site.abrirnavegador()
        if site is not None and site.navegador != -1:
            # Pega o delay configurado no objeto "Site"
            delay = site.delay

            # Botão de ‘login’
            botaologin = site.verificarobjetoexiste('NAME', 'autenticar')
            # Verifica se aparece a lista de opções
            if botaologin is not None:
                cmpusuario = site.verificarobjetoexiste('NAME', 'cliente_usuario')
                # Campo de Usuário
                if cmpusuario is not None:
                    cmpusuario.send_keys(linha[Usuario])
                # Campo de Senha
                cmpsenha = site.verificarobjetoexiste('NAME', 'cliente_senha')
                if cmpsenha is not None:
                    cmpsenha.send_keys(linha[Senha])

                    # Clica no botão de ‘login’
                    site.navegador.execute_script("arguments[0].click()", botaologin)

                    site.delay = 2
                    # Verifica Mensagem de Erro
                    mensagemerro = site.verificarobjetoexiste('CLASS_NAME', 'bp-formulario-fracasso')
                    if mensagemerro is None:
                        site.delay = delay

                        site.delay = 2
                        boletos = site.verificarobjetoexiste('LINK_TEXT', 'Gerar', itemunico=False)
                        site.delay = delay

                        if boletos:
                            for i, boleto in enumerate(boletos, start=1):
                                boletotemp = boletos[i - 1]
                                caminho_arquivo = site.monitorar_downloads_sem_href(
                                    link_elemento=boletotemp, timeout=120, via_request=True)
                                if caminho_arquivo:
                                    novonomearquivo = os.path.join(
                                        objeto.pastadownload,
                                        aux.left(linha[identificador], 4) + (
                                            f"_{i - 1}.pdf" if i > 1 else ".pdf")
                                    )
                                    listaboletos.append(
                                        aux.adicionarcabecalhopdf(caminho_arquivo,
                                                                  novonomearquivo,
                                                                  aux.left(linha[
                                                                               identificador],
                                                                           4)))
                                    if os.path.isfile(novonomearquivo):
                                        numboleto += 1
                                else:
                                    print(f"Falha ao capturar o arquivo baixado para o boleto {i}.\n"
                                          f"Função: {inspect.currentframe().f_code.co_name} \n"
                                          f"Cliente: {linha[identificador]}")
                                boletos = site.verificarobjetoexiste('LINK_TEXT', 'Gerar', itemunico=False)

                        # Retorna resposta na linha
                        linha[Resposta] = respostapadrao(numboleto)

                        if site is not None:
                            site.fecharsite()
                    else:
                        # Retorna a mensagem de erro do site
                        linha[Resposta] = re.sub('\n+', '\n', mensagemerro.text)
                        linha[ProblemaLogin] = True

        if listaboletos is not None:
            return listaboletos

    except Exception as e:
        linha[Resposta] = str(e)
        linha[CheckErro] = True


    finally:
        linha[Nomefuncao] = inspect.currentframe().f_code.co_name
        if site is not None:
            site.fecharsite()


# def bcf(objeto, linha):
#     site = None
#     listaboletos = []
#     linha[Nomefuncao] = inspect.currentframe().f_code.co_name
#     try:
#         linha[ProblemaLogin] = False
#         # Variável que vai retornar a quantidade de boletos
#         numboleto = 0
#         # Prepara o objeto
#         site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
#                               os.getenv('NOMEPROFILECOND'))
#         # Abre o browser
#         site.abrirnavegador()
#         # Verifica se iniciou o site
#
#         if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
#             # Caso esteja com outra página aberta, fecha
#             if site is not None:
#                 site.fecharsite()
#             # Inicia o browser
#             site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
#                                   os.getenv('NOMEPROFILECOND'))
#             # Abre o browser
#             site.abrirnavegador()
#         if site is not None and site.navegador != -1:
#             # Pega o delay configurado no objeto "Site"
#             delay = site.delay
#             objetos = site.verificarobjetoexiste('CLASS_NAME', 'text', itemunico=False)
#             for objeto in objetos:
#                 match objeto.get_attribute('name'):
#                     # Campo de usuário
#                     case 'login':
#                         objeto.clear()
#                         objeto.send_keys(linha[Usuario])
#                     case 'senha':
#                         objeto.clear()
#                         objeto.send_keys(linha[Senha])
#
#             # Botão de Login
#             botao = site.verificarobjetoexiste('CLASS_NAME', 'submit')
#             # Verifica se achou o botão de ‘login’
#             if botao is not None:
#                 # Clica no botão de ‘login’
#                 site.navegador.execute_script("arguments[0].click()", botao)
#                 time.sleep(1)
#                 # Diminui o tempo de busca pela mensagem de erro
#                 site.delay = 2
#                 # Mensagem de erro (Teste de usuário ou senha inválido)
#                 mensagemerro = site.verificarobjetoexiste('ID', 'ucLoginSistema_lbErroEntrar')
#                 if mensagemerro is None:
#                     # Segunda tela de erro de login (o site tem duas telas de erro possíveis)
#                     mensagemerro = site.verificarobjetoexiste('ID', 'errorLogin')
#
#                 # Teste de erro depois de login validado
#                 testetelalinks = site.verificarobjetoexiste('PARTIAL_LINK_TEXT', 'via de boleto')
#                 if mensagemerro is None:
#                     mensagemerro = site.verificarobjetoexiste('CLASS_NAME', 'alert')
#
#                 # Retorna o delay ao normal
#                 site.delay = delay
#                 # Teste se não tem mensagem de erro do site
#                 if mensagemerro is None or testetelalinks is not None:
#                     numboleto = 0
#                     site.navegador.execute_script("arguments[0].click()", testetelalinks)
#                     time.sleep(1)
#                     if site.verificarobjetoexiste('CLASS_NAME', 'tableRecibos', sotestar=True):
#                         links = site.verificarobjetoexiste('CSS_SELECTOR', "a [title='Baixar PDF']")
#                         if links:
#                             for i, boleto in enumerate(links, start=1):
#                                 caminho_arquivo = site.monitorar_downloads_sem_href(link_elemento=boleto, timeout=120,
#                                                                                     clickscript=True)
#                                 if caminho_arquivo:
#                                     novonomearquivo = os.path.join(
#                                         objeto.pastadownload,
#                                         aux.left(linha[identificador], 4) + (f"_{i - 1}.pdf" if i > 1 else ".pdf")
#                                     )
#                                     listaboletos.append(
#                                         aux.adicionarcabecalhopdf(caminho_arquivo,
#                                                                   novonomearquivo,
#                                                                   aux.left(linha[identificador],4)))
#                                     if os.path.isfile(novonomearquivo):
#                                         numboleto += 1
#                                 else:
#                                     print(f"Falha ao capturar o arquivo baixado para o boleto {i}.\n"
#                                           f"Função: {inspect.currentframe().f_code.co_name} \n"
#                                           f"Cliente: {linha[identificador]}")
#
#                     # Retorna resposta na linha
#                     linha[Resposta] = respostapadrao(numboleto)
#                 else:
#                     # Retorna resposta na linha
#                     linha[Resposta] = re.sub('\n+', '\n', mensagemerro.text)
#                     linha[ProblemaLogin] = True
#
#                 # Fecha o browser
#                 site.fecharsite()
#
#         if listaboletos is not None:
#             return listaboletos
#
#     except Exception as e:
#         linha[Resposta] = str(e)
#         linha[CheckErro] = True
#
#     finally:
#         linha[Nomefuncao] = inspect.currentframe().f_code.co_name
#         if site is not None:
#             site.fecharsite()


def cipa(objeto, linha):
    # Iniciar o trabalho, só copiado de outro
    site = None
    listaboletos = []
    linha[Nomefuncao] = __name__
    try:
        # Variável que vai retornar a quantidade de boletos
        numboleto = 0
        linha[ProblemaLogin] = False
        achouapartamento = False
        mudoustatus = False

        # Prepara o objeto
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
                              os.getenv('NOMEPROFILECOND'))
        site.delay = 30
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
                                  os.getenv('NOMEPROFILECOND'))
            # Abre o browser
            site.abrirnavegador()
        if site is not None and site.navegador != -1:
            # Campo de usuário
            email = site.verificarobjetoexiste('ID', 'email')
            # Verifica se achou o campo de usuário
            if email is not None:
                # Limpa o campo
                email.clear()
                # Preenche o campo de usuário
                email.send_keys(linha[Usuario])
                delay = site.delay
                site.delay = 10
                # Campo de senha
                cmpsenha = site.verificarobjetoexiste('ID', 'password')
                site.delay = delay
                # Se o campo senha não estiver visível ele "clica" no botão entrar para exibir o campo "senha"
                if cmpsenha is None:
                    # Botão de Salvar
                    btnsalvar = site.verificarobjetoexiste('ID', 'salvar')
                    # Verifica se achou o botão Salvar
                    if btnsalvar is not None:
                        site.navegador.execute_script("arguments[0].click()", btnsalvar)

                    cmpsenha = site.verificarobjetoexiste('ID', 'senha')

                if cmpsenha is not None:
                    if cmpsenha.is_displayed():
                        cmpsenha.send_keys(linha[Senha])
                        # Botão de Salvar
                        btnsalvar = site.verificarobjetoexiste('XPATH',
                                                               '/html/body/app-root/ion-content/layout/empty-layout/div/div/auth-sign-in/div/div[2]/div/form/button')
                        # Verifica se achou o botão Salvar
                        if btnsalvar is not None:
                            site.navegador.execute_script("arguments[0].click()", btnsalvar)
                            site.delay = 5
                            # Verifica se deu erro após apertar no botão para realizar o "LOGIN"
                            msgerro = site.verificarobjetoexiste('CSS_SELECTOR', 'div.cipafacil-alert-message')
                            msgerro1 = site.verificarobjetoexiste('ID', "mat-error-2")
                            site.delay = delay
                            if msgerro is None and msgerro1 is None:
                                # Operação normal
                                site.delay = 10

                                time.sleep(3)
                                # Verifica se está na página de lista de condomínios
                                if site.verificarobjetoexiste('CSS_SELECTOR', "condominium-list", sotestar=True):
                                    icondominio = site.buscar_por_proximidade('CSS_SELECTOR',
                                                                              'div.text-md.font-semibold.leading-tight',
                                                                              linha[Condominio], 55)
                                    if icondominio is not None:
                                        # Entra no condomínio selecionado
                                        site.navegador.execute_script("arguments[0].click()", icondominio)
                                        time.sleep(4)
                                        # Busca o botão de boleto
                                        # Usar verificarobjetoexiste para encontrar e clicar no link "Boleto", ignorando espaços adicionais
                                        botaoboleto = site.verificarobjetoexiste('XPATH',
                                                                                 "//a[contains(@href, '/boletos') and contains(normalize-space(.), 'Boleto')]",
                                                                                 iraoobjeto=True)

                                        if botaoboleto is not None:
                                            # Clica no botão de boleto
                                            botaoboleto.click()
                                            time.sleep(2)
                                            # Item de Boleto
                                            botaofiltro = site.verificarobjetoexiste('XPATH',
                                                                                     "//button//span[normalize-space(text())='Filtros']")
                                            if botaofiltro is not None:
                                                site.navegador.execute_script("arguments[0].click()", botaofiltro)
                                                time.sleep(2)
                                                filtrounidade = site.verificarobjetoexiste('XPATH',
                                                                                           '/html/body/div[2]/div[2]/div/mat-dialog-container/boleto-filter-dialog/div/div[2]/mat-form-field[1]/div/div[1]/div')
                                                time.sleep(1)

                                                if filtrounidade is not None:
                                                    site.navegador.execute_script("arguments[0].click()", filtrounidade)
                                                    time.sleep(3)
                                                    listaapartamentos = site.verificarobjetoexiste('CLASS_NAME',
                                                                                                   'mat-option-text',
                                                                                                   itemunico=False)
                                                    if listaapartamentos is not None:
                                                        for apartamento in listaapartamentos:
                                                            if linha[Apartamento] in apartamento.text:
                                                                site.navegador.execute_script("arguments[0].click()",
                                                                                              apartamento)
                                                                achouapartamento = True
                                                                break
                                                    if achouapartamento:
                                                        filtrostatus = site.verificarobjetoexiste('XPATH',
                                                                                                  '/html/body/div[2]/div[2]/div/mat-dialog-container/boleto-filter-dialog/div/div[2]/mat-form-field[2]/div/div[1]/div')
                                                        time.sleep(3)
                                                        if filtrostatus is not None:
                                                            site.navegador.execute_script("arguments[0].click()",
                                                                                          filtrostatus)
                                                            time.sleep(3)
                                                            listastatus = site.verificarobjetoexiste('CLASS_NAME',
                                                                                                     'mat-option-text',
                                                                                                     itemunico=False)

                                                            for status in listastatus:
                                                                if 'Em Aberto' in status.text:
                                                                    site.navegador.execute_script(
                                                                        "arguments[0].click()",
                                                                        status)
                                                                    time.sleep(1)
                                                                    mudoustatus = True
                                                                    break

                                                if achouapartamento or mudoustatus:
                                                    botaoaplicafiltro = site.verificarobjetoexiste('XPATH',
                                                                                                   '/html/body/div[2]/div[2]/div/mat-dialog-container/boleto-filter-dialog/div/div[3]/button[2]/span[1]')
                                                    if botaoaplicafiltro is not None:
                                                        site.navegador.execute_script("arguments[0].click()",
                                                                                      botaoaplicafiltro)

                                                    if achouapartamento:
                                                        site.delay = 2
                                                        botaofecharfiltro = site.verificarobjetoexiste('XPATH',
                                                                                                       '/html/body/div[2]/div[2]/div/mat-dialog-container/boleto-filter-dialog/div/div[1]/button/span[1]/mat-icon/svg')
                                                        # site.delay = delay
                                                        if botaofecharfiltro is not None:
                                                            site.navegador.execute_script("arguments[0].click()",
                                                                                          botaofecharfiltro)
                                                            site.delay = 2
                                                        # Verifica se tem a mensagem de que não tem boleto disponível
                                                        if not site.verificarobjetoexiste('XPATH',
                                                                                          "//img[@src='assets/images/boleto/boleto-empty.svg' and @alt='Sem funcionalidade' and contains(@class, 'w-full')]",
                                                                                          sotestar=True,
                                                                                          buscar_em_iframes=True):
                                                            site.delay = delay
                                                            if site.verificarobjetoexiste('XPATH',
                                                                                          "//img[@src='assets/icons/boletos/download.svg']",
                                                                                          sotestar=True,
                                                                                          buscar_em_iframes=True):
                                                                boletos = site.verificarobjetoexiste('XPATH',
                                                                                                     "//img[@src='assets/icons/boletos/download.svg']",
                                                                                                     buscar_em_iframes=True,
                                                                                                     itemunico=False)
                                                                if boletos:
                                                                    for i, boleto in enumerate(boletos, start=1):
                                                                        caminho_arquivo = site.monitorar_downloads_sem_href(
                                                                            link_elemento=boleto, timeout=120)
                                                                        if caminho_arquivo:
                                                                            novonomearquivo = os.path.join(
                                                                                objeto.pastadownload,
                                                                                aux.left(linha[identificador], 4) + (
                                                                                    f"_{i-1}.pdf" if i > 1 else ".pdf")
                                                                            )
                                                                            listaboletos.append(
                                                                                aux.adicionarcabecalhopdf(caminho_arquivo,
                                                                                                          novonomearquivo,
                                                                                                          aux.left(linha[
                                                                                                                       identificador],
                                                                                                                   4)))
                                                                            if os.path.isfile(novonomearquivo):
                                                                                numboleto += 1
                                                                        else:
                                                                            print(f"Falha ao capturar o arquivo baixado para o boleto {i}.\n"
                                                                                  f"Função: {inspect.currentframe().f_code.co_name} \n"
                                                                                  f"Cliente: {linha[identificador]}")
                                                                else:
                                                                    # Não existem boletos em aberto para buscar na lista do apartamento dado como entrada
                                                                    linha[Resposta] = respostapadrao(-1)
                                                                    site.fecharsite()
                                                    else:
                                                        linha[Resposta] = "Apartamento não encontrado!"
                                            if site is not None:
                                                site.fecharsite()

                                else:
                                    linha[Resposta] = mensagemsemcondominio

                                # Retorna resposta na linha
                                linha[Resposta] = respostapadrao(numboleto)
                            else:
                                if msgerro:
                                    # Mensagem de erro de Login
                                    linha[Resposta] = msgerro.text
                                    linha[ProblemaLogin] = True
                                if msgerro1:
                                    if len(linha[Resposta]) == 0:
                                        # Mensagem de erro de Login
                                        linha[Resposta] = msgerro1.text
                                        linha[ProblemaLogin] = True
                                    else:
                                        # Mensagem de erro de Login
                                        linha[Resposta] = linha[Resposta] + '; ' + msgerro1.text
                                        linha[ProblemaLogin] = True
            site.fecharsite()

        if listaboletos is not None:
            return listaboletos

    except Exception as e:
        linha[Resposta] = str(e)
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = inspect.currentframe().f_code.co_name
        if site is not None:
            site.fecharsite()


def fernandoefernandes(objeto, linha):

    site = None
    listaboletos = []
    linha[Nomefuncao] = inspect.currentframe().f_code.co_name
    try:
        linha[ProblemaLogin] = False
        numboleto = 0
        # Prepara o objeto
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
                              os.getenv('NOMEPROFILECOND'))
        site.delay = 10
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
                                  os.getenv('NOMEPROFILECOND'))
            # Abre o browser
            site.abrirnavegador()
        if site is not None and site.navegador != -1:
            # Pega o delay configurado no objeto "Site"
            delay = site.delay

            # Botão de ‘login’
            botaologin = site.verificarobjetoexiste('XPATH', "//button[contains(text(), 'E N T R A R')]")
            # Verifica se aparece a lista de opções
            if botaologin is not None:
                cmpusuario = site.verificarobjetoexiste('ID', 'edit_user')
                # Campo de Usuário
                if cmpusuario is not None:
                    cmpusuario.send_keys(linha[Usuario])
                # Campo de Senha
                cmpsenha = site.verificarobjetoexiste('ID', 'edit_pw')
                if cmpsenha is not None:
                    cmpsenha.send_keys(linha[Senha])

                    # Clica no botão de ‘login’
                    site.navegador.execute_script("arguments[0].click()", botaologin)

                    site.delay = 2
                    # Verifica Mensagem de Erro
                    mensagemerro = site.verificarobjetoexiste('XPATH', '/html/body/section[2]/div/form/div[5]')
                    if mensagemerro is None:
                        site.delay = delay
                        # Opção Boletos
                        primeironivel = site.verificarobjetoexiste('LINK_TEXT', 'Boletos')
                        if primeironivel is not None:
                            # Opção segunda via
                            primeironivel.click()
                            segundonivel = site.verificarobjetoexiste('LINK_TEXT', '2ª Via de Boleto')
                            if segundonivel is not None:
                                segundonivel.click()
                                time.sleep(2)
                                # Pega a lista de boletos
                                # Extrai o elemento pai <tr> para capturar o atributo 'onclick'
                                # Localiza todos os elementos com o ícone de download
                                boletos = site.verificarobjetoexiste('CSS_SELECTOR', "[class='fa fa-download']",
                                                                     itemunico=False)
                                if boletos:
                                    for i, boleto in enumerate(boletos, start=1):

                                        # Localiza o elemento pai <tr> que contém o atributo 'onclick'
                                        elemento_tr = site.verificarobjetoexiste(
                                            'XPATH', "./ancestor::tr[@onclick]", itemunico=True, elemento_pai=boleto
                                        )

                                        # Captura o atributo 'onclick' do elemento pai
                                        onclick_script = elemento_tr.get_attribute("onclick")

                                        # Extrai a URL do atributo 'onclick' usando regex

                                        match = re.search(r"window\.open\('(?:\./)?([^']+)", onclick_script)

                                        if match:
                                            # Constrói a URL completa
                                            url_relativa = match.group(1)  # Captura a URL relativa
                                            url_base = site.navegador.current_url.rsplit('/', 1)[0]
                                            url_completa = f"{url_base}/{url_relativa}"

                                            print(f"Baixando o arquivo da URL: {url_completa}")

                                            # Chama a função monitorar_downloads_sem_href usando a URL extraída
                                            caminho_arquivo = site.monitorar_downloads_sem_href(
                                                link_elemento=None,
                                                timeout=120,
                                                via_request=True,
                                                url_override=url_completa
                                            )

                                            if caminho_arquivo:
                                                novonomearquivo = os.path.join(
                                                    objeto.pastadownload,
                                                    aux.left(linha[identificador], 4) + (
                                                        f"_{i - 1}.pdf" if i > 1 else ".pdf")
                                                )
                                                listaboletos.append(
                                                    aux.adicionarcabecalhopdf(caminho_arquivo,
                                                                              novonomearquivo,
                                                                              aux.left(linha[
                                                                                           identificador],
                                                                                       4)))
                                                if os.path.isfile(novonomearquivo):
                                                    numboleto += 1
                                        else:
                                            print("Não foi possível capturar a URL do atributo onclick.")
                        # Retorna resposta na linha
                        # linha[Resposta] = respostapadrao(numboleto)

                    else:
                        # Retorna a mensagem de erro do site
                        linha[Resposta] = re.sub('\n+', '\n', mensagemerro.text)
                        # Retorna que não conseguiu logar
                        linha[ProblemaLogin] = True

        if listaboletos is not None:
            return listaboletos

    except Exception as e:
        linha[Resposta] = str(e)
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = inspect.currentframe().f_code.co_name
        if site is not None:
            site.fecharsite()


def ICondo(objeto, linha):
    site = None
    listaboletos = []
    try:
        # Variável que vai retornar a quantidade de boletos
        numboleto = 0
        linha[Nomefuncao] = inspect.currentframe().f_code.co_name
        linha[ProblemaLogin] = False
        atualizacaocadastral = False
        # Retorna o texto do site
        textosite = "https://" + info.retornaradministradora('nomereal', linha[Administradora],
                                                             'nomereduzido') + ".icondo.com.br/"
        textosite = textosite.lower()
        # Prepara o objeto
        site = web.TratarSite(textosite, os.getenv('NOMEPROFILECOND'))
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site
        if site.url != textosite or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(textosite, os.getenv('NOMEPROFILECOND'))
            # Abre o browser
            site.abrirnavegador()
        if site is not None and site.navegador != -1:
            # Pega o delay configurado no objeto "Site"
            delay = site.delay
            # Botão de Login
            botao = site.verificarobjetoexiste('NAME', 'wp-submit')
            # Verifica se achou o botão de ‘login’
            if botao is not None:
                # Campo de Usuário
                cmpusuario = site.verificarobjetoexiste('ID', 'user_login')
                if cmpusuario is not None:
                    cmpusuario.send_keys(linha[Usuario])
                # Campo de Senha
                cmpsenha = site.verificarobjetoexiste('ID', 'user_pass')
                if cmpsenha is not None:
                    cmpsenha.send_keys(linha[Senha])

                if info.retornaradministradora('nomereal', linha[Administradora], 'nomereduzido') == 'PORTAL':
                    cmpadm = site.verificarobjetoexiste('ID', 'selectAdm')
                    if cmpadm is not None:
                        cmpadm.send_keys(aux.encontrar_administradora(linha[Administradora])['NomeSelecao'])
                        selecionaritem = site.verificarobjetoexiste('ID', 'selectBox')
                        if selecionaritem is not None:
                            selecionaritem.click()

                # Clica no botão de ‘login’
                site.navegador.execute_script("arguments[0].click()", botao)
                # time.sleep(1)
                # Diminui o tempo de busca pela mensagem de erro
                site.delay = 2
                # Mensagem de erro (Teste de usuário ou senha inválido)
                mensagemerro = site.verificarobjetoexiste('ID', 'login_error')

                # Teste se não tem mensagem de erro do site
                if mensagemerro is None:
                    # Verifica se aparece a tela pedindo atualização cadastral
                    if site.verificarobjetoexiste('ID', 'atualizacao-cadastral', sotestar=True):
                        botaoatualizacao = site.verificarobjetoexiste('CSS_SELECTOR',
                                                                      "[class = 'swal2-confirm btn green']")
                        if botaoatualizacao is not None:
                            botaoatualizacao.click()
                            # time.sleep(1)
                        else:
                            botaoatualizacao = site.verificarobjetoexiste('CSS_SELECTOR', "[class = 'btn green']")
                            if botaoatualizacao is not None:
                                botaoatualizacao.click()
                                # time.sleep(1)
                            else:
                                botaoatualizacao = site.verificarobjetoexiste('CSS_SELECTOR',
                                                                              "[class = 'swal2-confirm btn green']")
                                if botaoatualizacao is not None:
                                    botaoatualizacao.click()
                                    # time.sleep(1)

                    # Verifica se a janela foi fechada
                    if site.verificarobjetoexiste('ID', 'atualizacao-cadastral', sotestar=True):
                        linha[Resposta] = 'Exigindo atualização cadastral, entrar manualmente para realizá-la (necessita de campo que não tem na planilha)'
                        atualizacaocadastral = True

                    site.delay = delay
                    if not atualizacaocadastral:
                        # Acha o botão de segunda via
                        botao2via = site.verificarobjetoexiste('XPATH',
                                                               '//a[contains(@href,"?arquivo=segunda-via-boleto")]')
                        if botao2via is not None:
                            botao2via.click()
                            # Objeto de Mensagem sem Boleto
                            site.delay = 2
                            if site.verificarobjetoexiste('CLASS_NAME', 'alert-danger', sotestar=True):
                                linha[Resposta] = respostapadrao(numboleto)
                            else:
                                site.delay = delay
                                boletos = site.verificarobjetoexiste('PARTIAL_LINK_TEXT', 'Recibo No.', itemunico=False)
                                if boletos:
                                    for i, boleto in enumerate(boletos, start=1):
                                        boletotemp = boletos[i - 1]
                                        caminho_arquivo = site.monitorar_downloads_sem_href(
                                            link_elemento=boletotemp, timeout=120, via_request=True)
                                        if caminho_arquivo:
                                            novonomearquivo = os.path.join(
                                                objeto.pastadownload,
                                                aux.left(linha[identificador], 4) + (
                                                    f"_{i - 1}.pdf" if i > 1 else ".pdf")
                                            )
                                            listaboletos.append(
                                                aux.adicionarcabecalhopdf(caminho_arquivo,
                                                                          novonomearquivo,
                                                                          aux.left(linha[
                                                                                       identificador],
                                                                                   4)))
                                            if os.path.isfile(novonomearquivo):
                                                numboleto += 1
                                        boletos = site.verificarobjetoexiste('PARTIAL_LINK_TEXT', 'Recibo No.',
                                                                                    itemunico=False)

                    time.sleep(1)

                    # Retorna resposta na linha
                    linha[Resposta] = respostapadrao(numboleto)
                else:
                    # Retorna resposta na linha
                    linha[Resposta] = re.sub('\n+', '\n', mensagemerro.text)
                    linha[ProblemaLogin] = True

                # Fecha o browser
                site.fecharsite()

        if listaboletos is not None:
            return listaboletos

    except Exception as e:
        erro_traceback = traceback.format_exc()
        linha[Resposta] = str(e) + "\nTraceback completo:\n" + erro_traceback
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = inspect.currentframe().f_code.co_name
        if site is not None:
            site.fecharsite()


def imodata(objeto, linha):
    site = None
    listaboletos = []
    try:
        linha[ProblemaLogin] = False
        linha[Nomefuncao] = inspect.currentframe().f_code.co_name
        numboleto = 0
        # Prepara o objeto
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
                              os.getenv('NOMEPROFILECOND'))
        site.delay = 10
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
                                  os.getenv('NOMEPROFILECOND'))
            # Abre o browser
            site.abrirnavegador()
        if site is not None and site.navegador != -1:
            # Pega o delay configurado no objeto "Site"
            delay = site.delay
            site.delay = 2
            # Verifica se tem usuário logado
            sairlogin = site.verificarobjetoexiste('PARTIAL_LINK_TEXT', 'Sair com segurança')
            # Se está logado sai do usuário
            if sairlogin is not None:
                sairlogin.click()

            site.delay = delay

            # Verifica se aparece a lista de opções
            if site.verificarobjetoexiste('ID', 'login', sotestar=True):
                opcondominio = site.verificarobjetoexiste('XPATH', '/html/body/div[2]/center/form/div/ul/li[1]/label')
                if opcondominio is not None:
                    opcondominio.click()
                    # Botão de Verificar Usuário
                    botaoverificar = site.verificarobjetoexiste('XPATH', '/html/body/div[2]/center/form/input')
                    if botaoverificar is not None:
                        # Campo de Usuário
                        cmpusuario = site.verificarobjetoexiste('ID', 'login')
                        if cmpusuario is not None:
                            cmpusuario.send_keys(linha[Usuario])

                        botaoverificar = site.verificarobjetoexiste('XPATH', '/html/body/div[2]/center/form/input')

                        # Clica no botão de Verificar Usuário
                        site.navegador.execute_script("arguments[0].click()", botaoverificar)

                        # Campo de Senha
                        cmpsenha = site.verificarobjetoexiste('ID', 'senha')
                        if cmpsenha is not None:
                            cmpsenha.send_keys(linha[Senha])

                        # Botão de ‘login’
                        botaologin = site.verificarobjetoexiste('XPATH', '/html/body/center/div/form/p[2]/input')

                        if botaologin is not None:
                            # Clica no botão de ‘login’
                            site.navegador.execute_script("arguments[0].click()", botaologin)

                            site.delay = 2
                            # Verifica Mensagem de Erro
                            mensagemerro = site.verificarobjetoexiste('ID', 'ucLoginSistema_lbErroEntrar')
                            if mensagemerro is None:
                                mensagemerro = site.verificarobjetoexiste(
                                    'CSS_SELECTOR',
                                    "div[style*='color:red']",
                                    itemunico=True
                                )
                                if mensagemerro is None:
                                    # Caso tenha logado
                                    # Pedir cadastro do CPF
                                    definircpf = site.verificarobjetoexiste('ID', 'formCPF')
                                    if definircpf is not None:
                                        fechar = site.verificarobjetoexiste('XPATH',
                                                                            '/html/body/div[7]/div/div/div[1]/button')
                                        if fechar is not None:
                                            fechar.click()

                                    site.delay = delay
                                    # Botão de 2.ª via
                                    botao2via = site.verificarobjetoexiste('PARTIAL_LINK_TEXT', '2a.Via')
                                    botao2via.click()

                                    frame = site.verificarobjetoexiste('NAME', 'MyFrame')
                                    site.irparaframe(frame)
                                    time.sleep(2)
                                    boletos = site.verificarobjetoexiste('PARTIAL_LINK_TEXT',
                                                                         'Clique aqui para gerar a 2ª via',
                                                                         itemunico=False)

                                    if boletos is not None:
                                        for i, boleto in enumerate(boletos, start=1):
                                            arquivo = ''
                                            boleto = boletos[i - 1]
                                            # Clica no botão de boleto
                                            boleto.click()
                                            # Busca o ícone do boleto
                                            iconeboleto = site.verificarobjetoexiste('CLASS_NAME', 'boleto')

                                            arquivo = ''
                                            if iconeboleto:
                                                texto = iconeboleto.text
                                                if 'ITAÚ' in texto.upper():
                                                    arquivo = 'Boleto Itaú.pdf'
                                                else:
                                                    arquivo = 'Boleto Bradesco.pdf'

                                            if len(arquivo):
                                                caminho_arquivo = site.monitorar_downloads_sem_href(link_elemento=iconeboleto,
                                                                                                    timeout=120,
                                                                                                    clickscript=True,
                                                                                                    nome_arquivo=arquivo)
                                            else:
                                                caminho_arquivo = ""

                                            if caminho_arquivo:
                                                novonomearquivo = os.path.join(
                                                    objeto.pastadownload,
                                                    aux.left(linha[identificador], 4) + (
                                                        f"_{i - 1}.pdf" if i > 1 else ".pdf")
                                                )
                                                listaboletos.append(

                                                    aux.adicionarcabecalhopdf(caminho_arquivo,
                                                                              novonomearquivo,
                                                                              aux.left(linha[identificador], 4)))
                                                if os.path.isfile(novonomearquivo):
                                                    numboleto += 1
                                            else:
                                                print(f"Falha ao capturar o arquivo baixado para o boleto {i}.\n"
                                                      f"Função: {inspect.currentframe().f_code.co_name} \n"
                                                      f"Cliente: {linha[identificador]}")

                                            # Volta pra aba original
                                            site.irparaaba(1)
                                            # Recarrega a tabela de segunda via
                                            botao2via = site.verificarobjetoexiste('XPATH', '//a[contains(@href,"/2via")]')
                                            botao2via.click()
                                            # Recaptura o frame
                                            frame = site.verificarobjetoexiste('NAME', 'MyFrame')
                                            # Volta pro frame
                                            site.irparaframe(frame)
                                            # Busca de novo os boletos
                                            boletos = site.verificarobjetoexiste('PARTIAL_LINK_TEXT',
                                                                                 'Clique aqui para gerar a 2ª via',
                                                                                 itemunico=False)


                                    # Retorna resposta na linha
                                    linha[Resposta] = respostapadrao(numboleto)

                                    if site is not None:
                                        site.fecharsite()
                                else:
                                    # Retorna a mensagem de erro do site
                                    linha[Resposta] = re.sub('\n+', '\n', mensagemerro.text)
                                    # Retorna que não conseguiu logar
                                    linha[ProblemaLogin] = True

        if listaboletos is not None:
            return listaboletos

    except Exception as e:
        erro_traceback = traceback.format_exc()
        linha[Resposta] = str(e) + "\nTraceback completo:\n" + erro_traceback
        # linha[Resposta] = str(e)
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = inspect.currentframe().f_code.co_name
        if site is not None:
            site.fecharsite()


def livefacilities(objeto, linha):
    site = None
    listaboletos = []
    linha[Nomefuncao] = inspect.currentframe().f_code.co_name
    unidadeencontrada = True
    try:
        linha[ProblemaLogin] = False
        numboleto = 0
        match info.retornaradministradora('nomereal', linha[Administradora], 'nomereduzido'):
            case "HFLEX":
                sigla = "sys"
                dataid = '375'

            case "VORTEX":
                sigla = "vortex"
                dataid = '199'

            case "BCF":
                sigla = "bcf"
                dataid = '199'
                meusdadosid = '162'

            case _:
                sigla = ''
                dataid = ''

        # Abre o site dependendo da administradora
        if len(sigla) > 0:
            # Retorna o texto do site
            if linha[Loginmultiplo] == 0:
                textosite = "http://%s.livefacilities.com.br/Index.aspx" % sigla
            else:
                textosite = "https://app.bcfadm.com.br/Index.aspx?url=/Portal/Home.aspx"

            textosite = textosite.lower()
            # Prepara o objeto
            site = web.TratarSite(textosite, os.getenv('NOMEPROFILECOND'))
            # Abre o browser
            site.abrirnavegador()
            # Verifica se iniciou o site
            if site.url != textosite or site is None:
                # Caso esteja com outra página aberta, fecha
                if site is not None:
                    site.fecharsite()
                # Inicia o browser
                site = web.TratarSite(textosite, os.getenv('NOMEPROFILECOND'))
                # Abre o browser
                site.abrirnavegador()
            if site is not None and site.navegador != -1:
                # Pega o delay configurado no objeto "Site"
                delay = site.delay

                site.delay = 2
                # Botão de "Entrar"
                botao = site.verificarobjetoexiste('ID', "btLogin")
                # Verifica se achou o botão de ‘login’
                if botao is not None:
                    botao.click()

                site.delay = delay
                # Campo de Usuário
                cmpusuario = site.verificarobjetoexiste('ID', 'ucLoginSistema_tbNomeEntrar')
                if cmpusuario is not None:
                    cmpusuario.send_keys(linha[Usuario])
                # Campo de Senha
                cmpsenha = site.verificarobjetoexiste('ID', 'ucLoginSistema_tbSenhaEntrar')
                if cmpsenha is not None:
                    cmpsenha.send_keys(linha[Senha])

                botaologin = site.verificarobjetoexiste('ID', 'ucLoginSistema_btEntrar')
                if botaologin is not None:
                    # Clica no botão de ‘login’
                    botaologin.click()
                    site.delay = 2
                    # Mensagem de erro
                    erro = site.verificarobjetoexiste('ID', 'ucLoginSistema_lbErroEntrar')
                    site.delay = delay
                    if erro is None:
                        if linha[Loginmultiplo] != 0:
                            # Botão Unidade
                            unidade = site.verificarobjetoexiste('XPATH',
                                                                    '//a[contains(@data-idmenu,"%s")]' % meusdadosid)
                            time.sleep(1)
                            if unidade is not None:
                                unidade.click()
                                minhasunidades = site.verificarobjetoexiste('XPATH',
                                                                             '//a[contains(@href,"/Unidade.aspx?")]')
                                time.sleep(1)
                                if minhasunidades is not None:
                                    minhasunidades.click()
                                    site.delay = 2
                                    boletos = site.verificarobjetoexiste('XPATH',
                                                                         '//a[contains(@title,"Alterar unidade")]',itemunico=False)
                                    condominios = site.verificarobjetoexiste('CSS_SELECTOR',
                                                                     "[id^='ContentBody_rpListaEmpreendimento_lbListaEmpreendimentoNome_']", itemunico=False)
                                    blocos = site.verificarobjetoexiste('CSS_SELECTOR',
                                                                     "[id^='ContentBody_rpListaEmpreendimento_lbListaBlocoNome_']", itemunico=False)
                                    unidades = site.verificarobjetoexiste('CSS_SELECTOR',
                                                                     "[id^='ContentBody_rpListaEmpreendimento_lbListaUnidadeCodigo_']", itemunico=False)

                                    unidades_unicas = [f"{condominio.text.split("-", 1)[1].strip()} - {bloco.text.strip()} / {unidade.text.strip()}"
                                                  for condominio, bloco, unidade in zip(condominios, blocos, unidades)]

                                    item = aux.busca_fuzzy(unidades_unicas, f"{linha[Condominio].strip()} - {linha[Apartamento].strip()}", 86)
                                    unidadeencontrada = item is not None
                                    if unidadeencontrada:
                                        boletos[item].click()
                                        time.sleep(2)
                        if unidadeencontrada:
                            # Botão Financeiro
                            financeiro = site.verificarobjetoexiste('XPATH', '//a[contains(@data-idmenu,"%s")]' % dataid)
                            time.sleep(1)
                            if financeiro is not None:
                                financeiro.click()
                                # Botão de segunda via
                                botaosegundavia = site.verificarobjetoexiste('XPATH',
                                                                             '//a[contains(@href,"/boletoLista.aspx?")]')
                                time.sleep(1)
                                # Verifica se existe o botão de segunda via
                                if botaosegundavia is not None:
                                    # Clica no botão de segunda via
                                    botaosegundavia.click()
                                    site.delay = 2

                                    # Pega o objeto do frame
                                    erroboleto = site.verificarobjetoexiste('ID', 'ContentBody_ucBoletoLista_pListaErro')
                                    site.delay = delay
                                    if erroboleto is None:
                                        boletos = site.verificarobjetoexiste('XPATH','//a[contains(@title,"Abrir Boleto")]', itemunico=False)
                                        time.sleep(1)
                                        if boletos:
                                            for i, boleto in enumerate(boletos, start=1):
                                                # Clica no botão de boleto
                                                boletos[i-1].click()
                                                time.sleep(2)
                                                frame = site.verificarobjetoexiste('CLASS_NAME', 'iziModal-iframe')
                                                if frame is not None:
                                                    site.irparaframe(frame)
                                                    # Pega o objeto do frame baseado no href
                                                    # parte_href = f'https://{sigla}.livefacilities.com.br/Operacional/PopUp/pCli_BoletoNovo.aspx'
                                                    # Botão Imprimir
                                                    botaoimprimir = site.verificarobjetoexiste('XPATH', f'//a[contains(@href, "/pCli_BoletoNovo.aspx")]')
                                                    caminho_arquivo = site.monitorar_downloads_sem_href(link_elemento=botaoimprimir, timeout=120, via_request=True)
                                                    if caminho_arquivo:
                                                        novonomearquivo = os.path.join(
                                                            objeto.pastadownload,
                                                            aux.left(linha[identificador], 4) + (
                                                                f"_{i-1}.pdf" if i > 1 else ".pdf")
                                                        )
                                                        listaboletos.append(
                                                            aux.adicionarcabecalhopdf(caminho_arquivo,
                                                                                      novonomearquivo,
                                                                                      aux.left(linha[identificador], 4)))
                                                        if os.path.isfile(novonomearquivo):
                                                            numboleto += 1
                                                    else:
                                                        print(f"Falha ao capturar o arquivo baixado para o boleto {i}.\n"
                                                              f"Função: {inspect.currentframe().f_code.co_name} \n"
                                                              f"Cliente: {linha[identificador]}")

                                                # site.irparaaba(1)
                                                site.navegador.refresh()
                                                # time.sleep(2)
                                                boletos = site.verificarobjetoexiste('XPATH','//a[contains(@title,"Abrir Boleto")]', itemunico=False)

                                        # Retorna resposta na linha
                                        linha[Resposta] = respostapadrao(numboleto)
                                    else:
                                        # Retorna a tela de erro
                                        linha[Resposta] = respostapadrao(numboleto)
                        else:
                            linha[Resposta] = 'Unidade não encontrada!'
                    else:
                        # Retorna a tela de erro
                        linha[Resposta] = erro.text
                        # Retorna que não conseguiu logar
                        linha[ProblemaLogin] = True

        if listaboletos is not None:
            return listaboletos

    except Exception as e:
        # linha[Resposta] = str(e)
        erro_traceback = traceback.format_exc()
        linha[Resposta] = str(e) + "\nTraceback completo:\n" + erro_traceback
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = inspect.currentframe().f_code.co_name
        if site is not None:
            site.fecharsite()


def nacional(objeto, linha):
    site = None
    listaboletos = []
    linha[Nomefuncao] = inspect.currentframe().f_code.co_name
    try:
        linha[ProblemaLogin] = False
        numboleto = 0
        # Prepara o objeto
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
                              os.getenv('NOMEPROFILECOND'))
        site.delay = 10
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
                                  os.getenv('NOMEPROFILECOND'))
            # Abre o browser
            site.abrirnavegador()
        if site is not None and site.navegador != -1:
            # Pega o delay configurado no objeto "Site"
            delay = site.delay

            # Botão de ‘login’
            botaologin = site.verificarobjetoexiste('ID', 'submit1')
            # Verifica se aparece a lista de opções
            if botaologin is not None:
                cmpusuario = site.verificarobjetoexiste('NAME', 'login')
                # Campo de Usuário
                if cmpusuario is not None:
                    cmpusuario.send_keys(linha[Usuario])
                # Campo de Senha
                cmpsenha = site.verificarobjetoexiste('NAME', 'senha')
                if cmpsenha is not None:
                    cmpsenha.send_keys(linha[Senha])

                    # Clica no botão de ‘login’
                    site.navegador.execute_script("arguments[0].click()", botaologin)

                    site.delay = 2
                    # Verifica Mensagem de Erro
                    mensagemerro = site.verificarobjetoexiste('XPATH',
                                                              '/html/body/div[2]/div/form/table/tbody/tr[2]/td/text()[1]')
                    if mensagemerro is None:
                        site.delay = delay

                        # Pega o frame do Menu
                        framemenu = site.verificarobjetoexiste('NAME', 'menu')

                        # Volta pro frame
                        if framemenu is not None:
                            site.irparaframe(framemenu)
                            # Botão de 2.ª via
                            botao2via = site.verificarobjetoexiste('CSS_SELECTOR', '[href = "Recibo.asp"]')
                            if botao2via is not None:
                                botao2via.click()

                                boletos = site.verificarobjetoexiste(
                                    'XPATH',
                                    "//a[starts-with(@href, 'ExibeRecibo.asp?NumRec=')]",
                                    iraoobjeto=False, itemunico=False, buscar_em_iframes=True
                                )

                                if boletos:
                                    for resultado in boletos:
                                        try:
                                            # Verifica se o resultado é uma tupla (iframe, elemento)
                                            if isinstance(resultado, tuple):
                                                iframe, boleto = resultado
                                                if iframe:
                                                    # Alterna para o contexto do iframe
                                                    site.irparaframe(iframe)
                                            else:
                                                boleto = resultado

                                            # Clica no boleto
                                            boleto.click()

                                            # Retorna ao contexto principal após interagir com o iframe
                                            if isinstance(resultado, tuple) and resultado[0]:
                                                site.sairdoframe()

                                        except Exception as e:
                                            print(f"Erro ao interagir com o boleto: {e}")

                                    # Verifica novamente para 2ª via dos boletos
                                    boletos_segunda_via = site.verificarobjetoexiste(
                                        'PARTIAL_LINK_TEXT',
                                        'Clique aqui para gerar a 2ª via',
                                        itemunico=False
                                    )

                                    if boletos_segunda_via:
                                        for i, boleto in enumerate(boletos_segunda_via, start=1):
                                            try:
                                                # Monitora o download do boleto
                                                caminho_arquivo = site.monitorar_downloads_sem_href(
                                                    link_elemento=boleto, timeout=120)

                                                if caminho_arquivo:
                                                    # Renomeia e adiciona o cabeçalho ao arquivo baixado
                                                    novonomearquivo = os.path.join(
                                                        objeto.pastadownload,
                                                        aux.left(linha[identificador], 4) + (
                                                            f"_{i - 1}.pdf" if i > 1 else ".pdf")
                                                    )

                                                    listaboletos.append(
                                                        aux.adicionarcabecalhopdf(caminho_arquivo, novonomearquivo,
                                                                                  aux.left(linha[identificador], 4))
                                                    )

                                                    if os.path.isfile(novonomearquivo):
                                                        numboleto += 1
                                                else:
                                                    print(f"Falha ao capturar o arquivo baixado para o boleto {i}.\n"
                                                          f"Função: {inspect.currentframe().f_code.co_name} \n"
                                                          f"Cliente: {linha[identificador]}")

                                            except Exception as e:
                                                print(f"Erro ao processar o boleto {i}: {e}")

                                    # Retorna resposta na linha
                                    linha[Resposta] = respostapadrao(numboleto)

                                    if site is not None:
                                        site.fecharsite()
                    else:
                        # Retorna a mensagem de erro do site
                        linha[Resposta] = re.sub('\n+', '\n', mensagemerro.text)
                        # Retorna que não conseguiu logar
                        linha[ProblemaLogin] = True

        if listaboletos is not None:
            return listaboletos

    except Exception as e:
        # linha[Resposta] = str(e)
        erro_traceback = traceback.format_exc()
        linha[Resposta] = str(e) + "\nTraceback completo:\n" + erro_traceback
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = inspect.currentframe().f_code.co_name
        if site is not None:
            site.fecharsite()


def modelopagina(objeto, linha):
    site = None
    listaboletos = []
    linha[ProblemaLogin] = False
    linha[Nomefuncao] = inspect.currentframe().f_code.co_name
    try:
        numboleto = 0
        # Prepara o objeto
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
                              os.getenv('NOMEPROFILECOND'))
        site.delay = 10
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
                                  os.getenv('NOMEPROFILECOND'))
            # Abre o browser
            site.abrirnavegador()
        if site is not None and site.navegador != -1:
            # Pega o delay configurado no objeto "Site"
            delay = site.delay

            # Botão de ‘login’
            botaologin = site.verificarobjetoexiste('XPATH', '/html/body/center/div/form/p[2]/input')
            # Verifica se aparece a lista de opções
            if botaologin is not None:
                cmpusuario = site.verificarobjetoexiste('ID', 'login')
                # Campo de Usuário
                if cmpusuario is not None:
                    cmpusuario.send_keys(linha[Usuario])
                # Campo de Senha
                cmpsenha = site.verificarobjetoexiste('ID', 'senha')
                if cmpsenha is not None:
                    cmpsenha.send_keys(linha[Senha])

                    # Clica no botão de ‘login’
                    site.navegador.execute_script("arguments[0].click()", botaologin)

                    site.delay = 2
                    # Verifica Mensagem de Erro
                    mensagemerro = site.verificarobjetoexiste('ID', 'ucLoginSistema_lbErroEntrar')
                    if mensagemerro is None:
                        site.delay = delay
                        # Botão de 2ª via
                        botao2via = site.verificarobjetoexiste('PARTIAL_LINK_TEXT', '2a.Via')
                        botao2via.click()

                        boletos = site.verificarobjetoexiste('PARTIAL_LINK_TEXT', 'Clique aqui para gerar a 2ª via',
                                                             itemunico=False)

                        if boletos:
                            for i, boleto in enumerate(boletos, start=1):
                                caminho_arquivo = site.monitorar_downloads_sem_href(link_elemento=boleto, timeout=120)
                                if caminho_arquivo:
                                    novonomearquivo = os.path.join(
                                        objeto.pastadownload,
                                        aux.left(linha[identificador], 4) + (f"_{i - 1}.pdf" if i > 1 else ".pdf")
                                    )
                                    listaboletos.append(
                                        aux.adicionarcabecalhopdf(caminho_arquivo,
                                                                  novonomearquivo,
                                                                  aux.left(linha[identificador], 4)))
                                    if os.path.isfile(novonomearquivo):
                                        numboleto += 1
                                else:
                                    print(f"Falha ao capturar o arquivo baixado para o boleto {i}.\n"
                                          f"Função: {inspect.currentframe().f_code.co_name} \n"
                                          f"Cliente: {linha[identificador]}")

                        # Retorna resposta na linha
                        linha[Resposta] = respostapadrao(numboleto)

                        if site is not None:
                            site.fecharsite()
                    else:
                        # Retorna a mensagem de erro do site
                        linha[Resposta] = re.sub('\n+', '\n', mensagemerro.text)
                        # Retorna que não conseguiu logar
                        linha[ProblemaLogin] = True

        if listaboletos is not None:
            return listaboletos

    except Exception as e:
        # linha[Resposta] = str(e)
        erro_traceback = traceback.format_exc()
        linha[Resposta] = str(e) + "\nTraceback completo:\n" + erro_traceback
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = inspect.currentframe().f_code.co_name
        if site is not None:
            site.fecharsite()


def protel(objeto, linha):
    site = None
    listaboletos = []
    try:
        linha[Nomefuncao] = inspect.currentframe().f_code.co_name
        linha[ProblemaLogin] = False
        mensagemerro = None
        numboleto = 0
        # Prepara o objeto
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
                              os.getenv('NOMEPROFILECOND'))
        site.delay = 10
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
                                  os.getenv('NOMEPROFILECOND'))
            # Abre o browser
            site.abrirnavegador()
        if site is not None and site.navegador != -1:
            # Pega o delay configurado no objeto "Site"
            delay = site.delay

            # Botão de ‘login’
            botaologin = site.verificarobjetoexiste('ID', 'botao_login')
            # Verifica se aparece a lista de opções
            if botaologin is not None:
                cmpusuario = site.verificarobjetoexiste('ID', 'usuario_login')
                # Campo de Usuário
                if cmpusuario is not None:
                    cmpusuario.send_keys(linha[Usuario])
                # Campo de Senha
                cmpsenha = site.verificarobjetoexiste('ID', 'usuario_senha')
                if cmpsenha is not None:
                    cmpsenha.send_keys(linha[Senha])

                    # Clica no botão de ‘login’
                    site.navegador.execute_script("arguments[0].click()", botaologin)

                    site.delay = 2

                    # Verifica se deu erro se o botão de segunda via não aparecer
                    botao2via = site.verificarobjetoexiste('ID', 'segunda_via')
                    site.delay = delay

                    if botao2via is None:
                        # Verifica Mensagem de Erro
                        mensagemerro = site.verificarobjetoexiste('ID', 'conteudo')
                    if mensagemerro is None:
                        botao2via.click()

                        time.sleep(1)
                        # Lista de boletos
                        boletos = site.verificarobjetoexiste('XPATH',
                                                             "//a//img[contains(@src,'/images/icone_pdf.gif')]",
                                                             itemunico=False)

                        if boletos:
                            for i, boleto in enumerate(boletos, start=1):
                                caminho_arquivo = site.monitorar_downloads_sem_href(link_elemento=boleto, timeout=120)
                                if caminho_arquivo:
                                    novonomearquivo = os.path.join(
                                        objeto.pastadownload,
                                        aux.left(linha[identificador], 4) + (f"_{i - 1}.pdf" if i > 1 else ".pdf")
                                    )
                                    listaboletos.append(
                                        aux.adicionarcabecalhopdf(caminho_arquivo,
                                                                  novonomearquivo,
                                                                  aux.left(linha[identificador], 4),
                                                                  centralizado=True))
                                    if os.path.isfile(novonomearquivo):
                                        numboleto += 1
                                else:
                                    print(f"Falha ao capturar o arquivo baixado para o boleto {i}.\n"
                                          f"Função: {inspect.currentframe().f_code.co_name} \n"
                                          f"Cliente: {linha[identificador]}")

                        # Retorna resposta na linha
                        linha[Resposta] = respostapadrao(numboleto)

                        if site is not None:
                            site.fecharsite()
                    else:
                        # Retorna a mensagem de erro do site
                        linha[Resposta] = re.sub('\n+', '\n', mensagemerro.text)
                        # Retorna que não conseguiu logar
                        linha[ProblemaLogin] = True

        if listaboletos is not None:
            return listaboletos

    except Exception as e:
        # linha[Resposta] = str(e)
        erro_traceback = traceback.format_exc()
        linha[Resposta] = str(e) + "\nTraceback completo:\n" + erro_traceback
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = inspect.currentframe().f_code.co_name
        if site is not None:
            site.fecharsite()


def superlogica(objeto, linha):
    linha[ProblemaLogin] = False
    linha[Nomefuncao] = inspect.currentframe().f_code.co_name
    site = None
    listaboletos = []
    try:
        # Quebra o bloco e apartamento da informação
        if '/' not in linha[Apartamento]:
            linhabloco = linha[Apartamento] + '/'
        else:
            linhabloco = linha[Apartamento]

        apartamentobuscado, blocobuscado = linhabloco.split('/')

        apartamentobuscado = apartamentobuscado.strip()
        # Garante que o apartamento tenha 4 caracteres
        if len(apartamentobuscado) > 0 and apartamentobuscado.isdigit():
            apartamentobuscado = aux.right('0000' + apartamentobuscado, 4).strip()

        # Trata o bloco
        blocobuscado = blocobuscado.strip()

        # Variável que vai retornar a quantidade de boletos
        numboleto = 0
        boletosanalises = 0
        achouapartamento = False

        # Prepara o objeto
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
                              os.getenv('NOMEPROFILECOND'))
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'),
                                  os.getenv('NOMEPROFILECOND'))
            # Abre o browser
            site.abrirnavegador()
        if site is not None and site.navegador != -1:
            # Campo de usuário
            email = site.verificarobjetoexiste('ID', 'email')
            # Verifica se achou o campo de usuário
            if email is not None:
                # Limpa o campo
                email.clear()
                # Preenche o campo de usuário
                email.send_keys(linha[Usuario])
                delay = site.delay
                site.delay = 2
                # Campo de senha
                cmpsenha = site.verificarobjetoexiste('ID', 'senha')
                site.delay = delay
                # Se o campo senha não estiver visível ele "clica" no botão entrar para exibir o campo "senha"
                if cmpsenha is None:
                    # Botão de Salvar
                    btnsalvar = site.verificarobjetoexiste('ID', 'salvar')
                    # Verifica se achou o botão Salvar
                    if btnsalvar is not None:
                        site.navegador.execute_script("arguments[0].click()", btnsalvar)
                    cmpsenha = site.verificarobjetoexiste('ID', 'senha')

                if cmpsenha is not None:
                    if cmpsenha.is_displayed():
                        cmpsenha.send_keys(linha[Senha])
                        # Botão de Salvar
                        btnsalvar = site.verificarobjetoexiste('ID', 'salvar')
                        # Verifica se achou o botão Salvar
                        if btnsalvar is not None:
                            site.navegador.execute_script("arguments[0].click()", btnsalvar)
                            site.delay = 2
                            # Verifica se deu erro após apertar no botão para realizar o "LOGIN"
                            msgerro = site.verificarobjetoexiste('ID', 'divMsgErroArea')
                            site.delay = delay
                            if msgerro is None:
                                # Operação normal
                                testalistacondominio = site.verificarobjetoexiste('CSS_SELECTOR',
                                                                                  "[class='item-menu lista-condominio']")

                                site.delay = delay

                                if testalistacondominio is not None:
                                    listacondominios = site.verificarobjetoexiste('CSS_SELECTOR',
                                                                                  "ul.listagem#lista li a",
                                                                                  itemunico=False)
                                    achoucondominio = False
                                    for icondominio in listacondominios:
                                        if icondominio.text == linha[Condominio].replace("'", ""):
                                            achoucondominio = True
                                            break

                                    if achoucondominio:
                                        # Botão de expansão de Menu
                                        setamenu = site.verificarobjetoexiste('CLASS_NAME', "seta-menu")
                                        if setamenu is not None:
                                            # Expande o Menu
                                            site.navegador.execute_script("arguments[0].click()", setamenu)
                                            time.sleep(1)
                                            # Clica no condomínio a ser extraído
                                            site.navegador.execute_script("arguments[0].click()", icondominio)
                                            site.delay = 2
                                            maisitens = site.verificarobjetoexiste('XPATH',
                                                                                   '//*[contains(@id, "btn_mais_itens_cobrancas")]')
                                            while maisitens is not None:
                                                maisitens.click()
                                                time.sleep(1)
                                                maisitens = site.verificarobjetoexiste('XPATH',
                                                                                       '//*[contains(@id, "btn_mais_itens_cobrancas")]')
                                                time.sleep(1)

                                            site.delay = delay
                                            condominioselecionado = site.verificarobjetoexiste('XPATH',
                                                                                               '/html/body/div[3]/div[2]/div/div[1]/div[2]/div/div/div/ul/li[3]/div/b')
                                            nomecondominioselecionado = condominioselecionado.text

                                            time.sleep(1)
                                            while nomecondominioselecionado.upper() != linha[Condominio].upper():
                                                condominioselecionado = site.verificarobjetoexiste('XPATH',
                                                                                                   '/html/body/div[3]/div[2]/div/div[1]/div[2]/div/div/div/ul/li[3]/div/b')
                                                if condominioselecionado is not None:
                                                    nomecondominioselecionado = condominioselecionado.text
                                                else:
                                                    nomecondominioselecionado = ''

                                            # Verifica se tem a mensagem de que não tem boleto disponível
                                            if not site.verificarobjetoexiste('CLASS_NAME', 'conteudo-nenhuma-cobranca',
                                                                              sotestar=True):
                                                if site.verificarobjetoexiste('CLASS_NAME', 'numero', sotestar=True):
                                                    numapartamentos = site.verificarobjetoexiste('CLASS_NAME', 'numero',
                                                                                                 itemunico=False)
                                                    dadosapartamentos = site.verificarobjetoexiste('CSS_SELECTOR',
                                                                                                   "[class='infos col-md-6']",
                                                                                                   itemunico=False)
                                                    complementoapartamentos = site.verificarobjetoexiste('CLASS_NAME',
                                                                                                         "complemento",
                                                                                                         itemunico=False)
                                                    listaapartamentos = zip(numapartamentos, dadosapartamentos,
                                                                            complementoapartamentos)
                                                    for indice, linhaweb in enumerate(listaapartamentos, start=1):
                                                        # Garante que só o exibido(no caso o "a vencer") seja testado
                                                        numeroap, dadosap, complementoap = linhaweb
                                                        # Vai para a tela inicial
                                                        site.irparaaba(titulo='Areadocondomino')
                                                        if numeroap.is_displayed():
                                                            if numeroap.text.upper() == apartamentobuscado and complementoap.text.upper() == blocobuscado:
                                                                achouapartamento = True
                                                                if "Indisponível" not in dadosap.text:
                                                                    site.navegador.execute_script(
                                                                        "arguments[0].click()",
                                                                        numeroap)
                                                                    # Seleciona a opção na janela "popup" que aparece
                                                                    while not site.verificarobjetoexiste('ID', 'salvar',
                                                                                                         sotestar=True):
                                                                        time.sleep(1)
                                                                    btnsalvar = site.verificarobjetoexiste('ID',
                                                                                                           'salvar')
                                                                    site.navegador.execute_script(
                                                                        "arguments[0].click()",
                                                                        btnsalvar)
                                                                    time.sleep(1)
                                                                    # Vai para a tela inicial
                                                                    site.irparaaba(titulo='Areadocondomino')

                                                                    # Vai para segunda página que ele abre com o boleto resumido
                                                                    site.irparaaba(indice=site.num_abas())

                                                                    achoupagina = (
                                                                            'SEGUNDAVIA' in site.navegador.current_url.upper())

                                                                    time.sleep(1)

                                                                    if achoupagina:
                                                                        # Aperta no botão para gerar o boleto
                                                                        btngerarboleto = site.verificarobjetoexiste(
                                                                            'XPATH',
                                                                            "//*[contains(@class, 'botao') and contains(@class, 'pagarBoleto')]")
                                                                     #    btngerarboleto = site.verificarobjetoexiste('CSS_SELECTOR', "[class='botao pagarBoleto']",
                                                                     # itemunico=False)
                                                                        if btngerarboleto is not None:

                                                                            caminho_arquivo = site.monitorar_downloads_sem_href(
                                                                                link_elemento=btngerarboleto,
                                                                                timeout=120,
                                                                                clickscript=True)
                                                                            if caminho_arquivo:
                                                                                novonomearquivo = os.path.join(
                                                                                    objeto.pastadownload,
                                                                                    aux.left(linha[identificador],
                                                                                             4) + (
                                                                                        f"_{numboleto}.pdf" if numboleto > 0 else ".pdf")
                                                                                )
                                                                                # Renomeia o arquivo baixado para o código de cliente
                                                                                listaboletos.append(
                                                                                    aux.adicionarcabecalhopdf(
                                                                                        caminho_arquivo,
                                                                                        novonomearquivo,
                                                                                        aux.left(linha[identificador],
                                                                                                 4)))
                                                                                if os.path.isfile(novonomearquivo):
                                                                                    numboleto += 1

                                                                                time.sleep(1)

                                                                        site.irparaaba(2)
                                                                        site.fecharaba()
                                                                        site.irparaaba(1)

                                                                    else:
                                                                        boletosanalises += 1
                                                    site.navegador.close()
                                            else:
                                                # Não existem boletos em aberto para buscar na lista do apartamento dado como entrada
                                                linha[Resposta] = respostapadrao(-1)
                                    else:
                                        linha[Resposta] = mensagemsemcondominio

                                if len(linha[Resposta]) == 0:
                                    if not achouapartamento:
                                        linha[Resposta] = 'Apartamento não encontrado!'
                                    else:
                                        linha[Resposta] = respostapadrao(numboleto, boletosanalises)
                            else:
                                # Mensagem de erro de Login
                                linha[Resposta] = msgerro.text
                                linha[ProblemaLogin] = True

        if listaboletos is not None:
            return listaboletos

    except Exception as e:
        linha[Resposta] = str(e)
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = inspect.currentframe().f_code.co_name
        if site is not None:
            site.fecharsite()


def Webware(objeto, linha):
    site = None
    listaboletos = []
    linha[ProblemaLogin] = False
    loginmultiplo = False
    achou_unidade = False
    try:
        numboleto = 0
        verificaboletos = False
        linha[Nomefuncao] = __name__
        match info.retornaradministradora('nomereal', linha[Administradora], 'nomereduzido'):
            case "CONAC":
                codigo = "19834090"

            case "LOWNDES":
                codigo = "33105362"

            case "ACIR":
                codigo = "42168518"

            case _:
                codigo = ''

        # Abre o site dependendo da administradora
        if len(codigo) > 0:
            # Retorna o texto do site
            textosite = "https://www.webware.com.br/bin/administradora/default.asp?adm=" + str(codigo)
            textosite = textosite.lower()
            mensagemboletos = ''
            # Prepara o objeto
            site = web.TratarSite(textosite, os.getenv('NOMEPROFILECOND'))
            # Abre o browser
            site.abrirnavegador()
            # Verifica se iniciou o site
            if site.url != textosite or site is None:
                # Caso esteja com outra página aberta, fecha
                if site is not None:
                    site.fecharsite()
                # Inicia o browser
                site = web.TratarSite(textosite, os.getenv('NOMEPROFILECOND'))
                # Abre o browser
                site.abrirnavegador()
            if site is not None and site.navegador != -1:
                # Pega o delay configurado no objeto "Site"
                delay = site.delay
                # Botão de Login
                botao = site.verificarobjetoexiste('CSS_SELECTOR', "[type='Submit']")
                # Verifica se achou o botão de ‘login’
                if botao is not None:
                    # Campo de Usuário
                    cmpusuario = site.verificarobjetoexiste('NAME', 'mem')
                    if cmpusuario is not None:
                        cmpusuario.send_keys(linha[Usuario])
                    # Campo de Senha
                    cmpsenha = site.verificarobjetoexiste('NAME', 'pass')
                    if cmpsenha is not None:
                        cmpsenha.send_keys(linha[Senha])
                    # Clica no botão de ‘login’
                    botao.click()
                    time.sleep(2)
                    # Diminui o delay pra uma verificação rápida
                    site.delay = 2
                    # Mensagem de erro
                    erro = site.verificarobjetoexiste('CLASS_NAME', 'mensagem-erro')
                    site.delay = delay
                    if erro is None:
                        # Diminui o delay pra uma verificação rápida
                        site.delay = 2
                        erro = site.verificarobjetoexiste('ID', 'dvContainer')
                        site.delay = delay
                        if erro is None:
                            selecionarcondominio = site.verificarobjetoexiste('CLASS_NAME', "card-body", itemunico=False)
                            botoesentrada = site.verificarobjetoexiste('CLASS_NAME', "bt", itemunico=False)


                            loginmultiplo = (len(botoesentrada) > 0)
                            if loginmultiplo:
                                if len(botoesentrada) > 0:
                                    if selecionarcondominio:
                                        unidadeencontrada, candidatos_condominios, candidatos_unidade = site.filtrar_itens_por_fuzzy(
                                            selecionarcondominio,
                                            "CLASS_NAME",
                                            "txt-titulo",
                                            linha[Condominio],
                                            linha[Apartamento],
                                            corte_condominio=70,
                                            corte_apartamento=80
                                        )
                                        if len(candidatos_unidade) == 1:
                                            botoesentrada[candidatos_unidade.index()].click()
                                            achou_unidade = True
                                        else:
                                            if len(candidatos_condominios) == 1:
                                                botoesentrada[candidatos_condominios[0][0]].click()
                                                achou_unidade = True
                                else:
                                    print("Nenhum condomínio encontrado.")

                            if achou_unidade or not loginmultiplo:
                                # Vai pra área de segunda via diretamente
                                site.navegador.get("https://servc9-1.webware.com.br/bin/aplic/u_2via_net.asp")
                                time.sleep(1)

                                # Pega o objeto do frame
                                janelaboleto = site.verificarobjetoexiste('CLASS_NAME', 'prestacao-interativa')
                                if janelaboleto is not None:
                                    # Mudar para o frame dos boletos
                                    site.irparaframe(janelaboleto)
                                    time.sleep(2)
                                    site.delay = 2
                                    # Verifica se tem e está visível o ícone de carregando no frame
                                    telacarregando = site.verificarobjetoexiste('CLASS_NAME', 'carregando')
                                    site.delay = delay
                                    if telacarregando is not None:
                                        while telacarregando.is_displayed():
                                            time.sleep(1)
                                    botaoalerta = site.verificarobjetoexiste('CLASS_NAME', 'popup-alerta-botoes',
                                                                             iraoobjeto=True)

                                    time.sleep(2)
                                    # Verifica se tem o botão de aceitar o alerta
                                    if botaoalerta is not None:
                                        botaoalerta.click()

                                    opcaoencontrada = True
                                    site.delay = 2
                                    combobox = site.verificarobjetoexiste('CLASS_NAME', 'v-input__control')
                                    site.delay = delay
                                    if combobox:
                                        # Clica na combobox para expandir as opções
                                        combobox.click()
                                        time.sleep(1)
                                        # Obtém todas as opções do dropdown
                                        opcoes = site.verificarobjetoexiste('CLASS_NAME', 'v-list-item__content', itemunico=False)

                                        # Variável com o apartamento que você está procurando

                                        if len((linha[Apartamento] or "").strip()) == 0:
                                            bloco = linha[Usuario]
                                            bloco = bloco.replace(linha[Senha],'')
                                            bloco = bloco.replace('BL', 'PB')
                                            item = bloco + '/Apart/' + str(linha[Senha].replace('APTO','')).zfill(4)
                                            if item:
                                                for opcao in opcoes:
                                                    texto_opcao = opcao.text.strip()
                                                    if texto_opcao == item:
                                                        opcaoencontrada = True
                                                        # Seleciona o apartamento
                                                        opcao.click()
                                                        time.sleep(1)
                                                        btnconsulta = site.verificarobjetoexiste('ID', 'container-btn-consulta')
                                                        btnconsulta.click()
                                                        break
                                                else:
                                                    opcaoencontrada = False
                                                    print(f"Apartamento {item} não encontrado na lista.")
                                        else:
                                            item = linha[Apartamento].strip()
                                            itemselecionado = site.buscar_por_proximidade('CLASS_NAME','v-list-item__content', item, 40, elemento_do_item=True)
                                            if type(itemselecionado) is not str:
                                                itemselecionado.click()
                                            else:
                                                mensagemboletos = itemselecionado
                                                verificaboletos = False

                                        if len(mensagemboletos) == 0:
                                            mensagemboletos = site.verificarobjetoexiste('ID', 'mensagem-indisponivel')

                                        if not mensagemboletos:
                                            verificaboletos = True

                                    else:
                                        verificaboletos = True

                                    if verificaboletos:
                                        # Pega todos os botões de gerar boleto do frame
                                        boletos = site.verificarobjetoexiste('XPATH',
                                                                             '//a[contains(@href,"/relatorioboleto/BuscarBoletoPorRecibo?")]',
                                                                             itemunico=False,
                                                                             esperar_clicavel=True)
                                        if boletos:

                                            for i, boleto in enumerate(boletos, start=1):
                                                caminho_arquivo = site.monitorar_downloads_sem_href(link_elemento=boleto,
                                                                                                    timeout=120, via_request=True)
                                                if caminho_arquivo:
                                                    novonomearquivo = os.path.join(
                                                        objeto.pastadownload,
                                                        aux.left(linha[identificador], 4) + (f"_{i - 1}.pdf" if i > 1 else ".pdf")
                                                    )
                                                    listaboletos.append(
                                                        aux.adicionarcabecalhopdf(caminho_arquivo,
                                                                                  novonomearquivo,
                                                                                  aux.left(linha[identificador], 4)))
                                                    if os.path.isfile(novonomearquivo):
                                                        numboleto += 1
                                                else:
                                                    print(f"Falha ao capturar o arquivo baixado para o boleto {i}.\n"
                                                          f"Função: {inspect.currentframe().f_code.co_name} \n"
                                                          f"Cliente: {linha[identificador]}")

                                    # Volta pra aba original
                                    site.irparaaba(1)
                                    # Pega o objeto do frame
                                    janelaboleto = site.verificarobjetoexiste('CLASS_NAME', 'prestacao-interativa')
                                    if janelaboleto:
                                        # Volta pro frame
                                        site.irparaframe(janelaboleto)

                                    # Retorna resposta na linha
                                    if verificaboletos:
                                        if isinstance(mensagemboletos, str):
                                            if not mensagemboletos:  # Verifica se a string está vazia
                                                linha[Resposta] = respostapadrao(numboleto)
                                        elif hasattr(mensagemboletos, 'text'):
                                            if not mensagemboletos.text:  # Verifica se o atributo 'text' está vazio
                                                linha[Resposta] = respostapadrao(numboleto)
                                            else:
                                                linha[Resposta] = mensagemboletos.text
                                        else:
                                            linha[Resposta] = respostapadrao(numboleto)
                                    else:
                                        if isinstance(mensagemboletos, str):
                                            if not mensagemboletos:  # Verifica se a string está vazia
                                                linha[Resposta] = respostapadrao(numboleto)
                                        elif hasattr(mensagemboletos, 'text'):
                                            if not mensagemboletos.text:  # Verifica se o atributo 'text' está vazio
                                                linha[Resposta] = respostapadrao(numboleto)
                                            else:
                                                linha[Resposta] = mensagemboletos.text
                                        else:
                                            linha[Resposta] = respostapadrao(numboleto)

                                else:
                                    linha[Resposta] = 'Unidade não encontrada!'
                        else:
                            # Retorna a tela de erro
                            linha[Resposta] = erro.text
                            # Retorna que não conseguiu logar
                            linha[ProblemaLogin] = True

                    else:
                        # Retorna a tela de erro
                        linha[Resposta] = erro.text
                        # Retorna que não conseguiu logar
                        linha[ProblemaLogin] = True

        if site is not None:
            site.fecharsite()

        if listaboletos is not None:
            return listaboletos

    except Exception as e:
        # linha[Resposta] = str(e)
        erro_traceback = traceback.format_exc()
        linha[Resposta] = str(e) + "\nTraceback completo:\n" + erro_traceback
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = inspect.currentframe().f_code.co_name
        if site is not None:
            site.fecharsite()


def respostapadrao(quant, quantanalise=0):
    if quantanalise == 0 or quant == -1:
        match quant:
            case 0:
                resposta = 'Sem Boleto'
            case 1:
                resposta = 'Salvo ' + str(quant) + ' boleto'
            case -1:
                resposta = 'Sem boletos em aberto para analisar'
            case _:
                resposta = 'Salvo ' + str(quant) + ' boletos'
    else:
        match quant:
            case 0:
                match quantanalise:
                    case 1:
                        resposta = 'Tem ' + str(quantanalise) + 'boleto para analisar (provável acordo)'
                    case _:
                        resposta = 'Tem ' + str(quantanalise) + 'boletos para analisar (provável acordo)'

            case 1:
                resposta = 'Salvo ' + quant + ' boleto'
                match quantanalise:
                    case 1:
                        resposta = resposta + " e tem " + str(
                            quantanalise) + " boleto para ser analisado (provável em acordo)"
                    case _:
                        resposta = resposta + " e tem " + str(
                            quantanalise) + " boletos para ser analisado (provável em acordo)"

            case _:
                resposta = 'Salvo ' + quant + ' boletos'
                match quantanalise:
                    case 1:
                        resposta = resposta + " e tem " + str(
                            quantanalise) + " boleto para ser analisado (provável em acordo)"
                    case _:
                        resposta = resposta + " e tem " + str(
                            quantanalise) + " boletos para ser analisado (provável em acordo)"

    return resposta
