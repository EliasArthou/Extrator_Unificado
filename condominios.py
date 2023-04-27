import time
import auxiliares as aux
import sensiveis as info
import web
import re
import os

identificador = 0
Usuario = 1
Senha = 2
Administradora = 3
Condominio = 4
Apartamento = 5
Resposta = 6
CheckArquivo = 7
CheckErro = 8
Nomefuncao = 9
ProblemaLogin = 10

timeout = 10
pastadownload = aux.caminhoprojeto('Downloads')
tempoesperadownload = 180
mensagemerropadrao = 'Deu erro! Tentar novamente!'
mensagemsemcondominio = "Condomínio não encontrado!"


def abrj(objeto, linha):
    site = None
    try:
        # Quebra o bloco e apartamento da informação
        if '/' not in linha[Apartamento]:
            linhabloco = linha[Apartamento] + '/'
        else:
            linhabloco = linha[Apartamento]

        apartamentobuscado, blocobuscado = linhabloco.split('/')

        # Garante que o apartamento tenha 4 caracteres
        if len(apartamentobuscado.strip()) > 0:
            apartamentobuscado = aux.right('0000' + apartamentobuscado.strip(), 4).strip()

        # Trata o bloco
        blocobuscado = blocobuscado.strip()

        # Variável que vai retornar a quantidade de boletos
        numboleto = 0
        boletosanalises = 0
        achouapartamento = False

        # Prepara o objeto
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
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
                                testalistacondominio = site.verificarobjetoexiste('CSS_SELECTOR', "[class='item-menu lista-condominio']")
                                if testalistacondominio is not None:
                                    listacondominios = site.verificarobjetoexiste('CSS_SELECTOR', "[href = '#']", itemunico=False)
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
                                            condominioselecionado = site.verificarobjetoexiste('XPATH', '/html/body/div[3]/div[2]/div/div[1]/div[2]/div/div/div/ul/li[3]/div/b')
                                            condominioselecionado = condominioselecionado.text

                                            time.sleep(1)
                                            while condominioselecionado != linha[Condominio].upper():
                                                condominioselecionado = site.verificarobjetoexiste('XPATH', '/html/body/div[3]/div[2]/div/div[1]/div[2]/div/div/div/ul/li[3]/div/b')
                                                if condominioselecionado is not None:
                                                    condominioselecionado = condominioselecionado.text
                                                else:
                                                    condominioselecionado = ''

                                            # Verifica se tem a mensagem de que não tem boleto disponível
                                            if not site.verificarobjetoexiste('CLASS_NAME', 'conteudo-nenhuma-cobranca', sotestar=True):
                                                if site.verificarobjetoexiste('CLASS_NAME', 'numero', sotestar=True):
                                                    numapartamentos = site.verificarobjetoexiste('CLASS_NAME', 'numero', itemunico=False)
                                                    dadosapartamentos = site.verificarobjetoexiste('CSS_SELECTOR', "[class='infos col-md-6']", itemunico=False)
                                                    complementoapartamentos = site.verificarobjetoexiste('CLASS_NAME', "complemento", itemunico=False)
                                                    listaapartamentos = zip(numapartamentos, dadosapartamentos, complementoapartamentos)
                                                    for indice, linhaweb in enumerate(listaapartamentos):
                                                        # Garante que só o exibido(no caso o "a vencer") seja testado
                                                        numeroap, dadosap, complementoap = linhaweb
                                                        # Vai para a tela inicial
                                                        site.irparaaba(titulo='Areadocondomino')
                                                        if numeroap.is_displayed():
                                                            if numeroap.text == apartamentobuscado and complementoap.text == blocobuscado:
                                                                achouapartamento = True
                                                                if "Indisponível" not in dadosap.text:
                                                                    site.navegador.execute_script("arguments[0].click()", numeroap)
                                                                    # Seleciona a opção na janela "popup" que aparece
                                                                    while not site.verificarobjetoexiste('ID', 'salvar', sotestar=True):
                                                                        time.sleep(1)
                                                                    btnsalvar = site.verificarobjetoexiste('ID', 'salvar')
                                                                    site.navegador.execute_script("arguments[0].click()", btnsalvar)
                                                                    time.sleep(1)
                                                                    # Vai para a tela inicial
                                                                    site.irparaaba(titulo='Areadocondomino')

                                                                    # Vai para segunda página que ele abre com o boleto resumido
                                                                    site.irparaaba(indice=site.num_abas())

                                                                    achoupagina = ('SEGUNDAVIA' in site.navegador.current_url.upper())

                                                                    if achoupagina:
                                                                        # Aperta no botão para gerar o boleto
                                                                        btngerarboleto = site.verificarobjetoexiste('CSS_SELECTOR', "[class='botao pagarBoleto pagarBoletoParcelado']")
                                                                        if btngerarboleto is not None:
                                                                            site.navegador.execute_script("arguments[0].click()", btngerarboleto)
                                                                            btngerarboleto = site.verificarobjetoexiste('ID', 'btnSubmitParcelamentoCartao')
                                                                            while btngerarboleto is None:
                                                                                btngerarboleto = site.verificarobjetoexiste('ID', 'btnSubmitParcelamentoCartao')
                                                                                time.sleep(1)

                                                                            while not btngerarboleto.is_displayed():
                                                                                time.sleep(1)

                                                                            site.navegador.execute_script("arguments[0].click()", btngerarboleto)

                                                                            # Espera o download finalizar
                                                                            site.esperadownloads(objeto.pastadownload, timeout)

                                                                            # Define o nome (baseado no nr do boleto)
                                                                            if numboleto == 0:
                                                                                novonomearquivo = objeto.pastadownload + '\\' + linha[identificador] + '_' + \
                                                                                                  info.retornaradministradora('nomereal', linha[Administradora], 'nomereduzido') + '.pdf'
                                                                            else:
                                                                                novonomearquivo = objeto.pastadownload + '\\' + linha[identificador] + '_' + \
                                                                                                  info.retornaradministradora('nomereal', linha[Administradora], 'nomereduzido') + '_' \
                                                                                                  + str(numboleto) + '.pdf'

                                                                            # Espera o download finalizar e "pega" o arquivo baixado
                                                                            arquivobaixado = site.pegaarquivobaixado(tempoesperadownload)
                                                                            time.sleep(1)

                                                                            # Verifica se o arquivo baixado de fato existe
                                                                            if os.path.isfile(objeto.pastadownload + '\\' + arquivobaixado):
                                                                                # Renomeia o arquivo baixado para o código de cliente
                                                                                aux.adicionarcabecalhopdf(objeto.pastadownload + '\\' + arquivobaixado, novonomearquivo, linha[identificador], True)
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
                                        linha[Resposta] = "Condomínio não encontrado!"

                                if len(linha[Resposta]) == 0:
                                    if not achouapartamento:
                                        linha[Resposta] = 'Apartamento não encontrado!'
                                    else:
                                        linha[Resposta] = respostapadrao(numboleto, boletosanalises)
                            else:
                                # Mensagem de erro de Login
                                linha[Resposta] = msgerro.text
                                linha[ProblemaLogin] = True

    except Exception as e:
        linha[Resposta] = str(e)
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = __name__
        if site is not None:
            site.fecharsite()


def apsa(objeto, linha):
    # XPATH ícone loading : '/html/body/app-root/app-loading-screen/div/div/div[2]'
    # Tela de Loading que fica "invisível" (XPATH): '/html/body/app-root/app-loading-screen'
    site = None
    telaloading = '/html/body/app-root/app-loading-screen'

    try:
        # Variável que vai retornar a quantidade de boletos
        numboleto = 0
        # Prepara o objeto
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
        # Abre o browser
        site.abrirnavegador(False)

        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
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
                        mensagemerro = site.verificarobjetoexiste('XPATH', '/html/body/app-ligthboxes-default[1]/div/div/div[2]/div')
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
                                            iconesboletos = site.verificarobjetoexiste('CLASS_NAME', 'Actions_Option_Label', itemunico=False)

                                            # "Varre" os itens clicáveis da lista de boletos
                                            for boleto in iconesboletos:
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
                                                            lembrardepois = site.verificarobjetoexiste('ID', 'ctl00_ContentPlaceHolder1_btnLembrarDepois')
                                                            # Verifica se achou o botão
                                                            if lembrardepois is not None:
                                                                # Clica no botão de lembrar depois
                                                                site.navegador.execute_script("arguments[0].click()", lembrardepois)
                                                                time.sleep(1)
                                                                # Define o novo nome e caminho do arquivo baixado (baseado no nr do boleto)
                                                                if numboleto == 0:
                                                                    novonomearquivo = objeto.pastadownload + '\\' + linha[identificador] + '_' + linha[Administradora] + '.pdf'
                                                                else:
                                                                    novonomearquivo = objeto.pastadownload + '\\' + linha[identificador] + '_' + linha[Administradora] + '_' + str(numboleto + 1) + '.pdf'

                                                                # Espera o download finalizar e "pega" o arquivo baixado
                                                                arquivobaixado = site.pegaarquivobaixado(tempoesperadownload)
                                                                time.sleep(1)

                                                                # Verifica se o arquivo baixado de fato existe
                                                                if os.path.isfile(objeto.pastadownload + '\\' + arquivobaixado):
                                                                    # Renomeia o arquivo baixado para o código de cliente
                                                                    aux.adicionarcabecalhopdf(objeto.pastadownload + '\\' + arquivobaixado, novonomearquivo, linha[identificador], True)
                                                                    numboleto += 1
                                                                    time.sleep(1)

                                                                time.sleep(1)
                                                                site.irparaaba(2)
                                                                time.sleep(1)
                                                                site.navegador.refresh()
                                                                time.sleep(1)

                                                            time.sleep(1)
                                                            site.fecharaba()
                                                            time.sleep(1)
                                                            site.irparaaba(1)
                                                            time.sleep(1)
                                                            pesquisasatisfacao = site.verificarobjetoexiste('XPATH', '/html/body/yes-or-no-with-link[1]/div/div/div[4]/button')
                                                            if pesquisasatisfacao is not None:
                                                                site.navegador.execute_script("arguments[0].click()", pesquisasatisfacao)
                                                                # Verifica site carregando
                                                                # ====================================================================
                                                                objtelaloading = site.verificarobjetoexiste('XPATH', telaloading)
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

    except Exception as e:
        linha[Resposta] = str(e)
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = __name__
        if site is not None:
            site.fecharsite()


def bap(objeto, linha):
    site = None
    try:
        numboleto = 0
        # Prepara o objeto
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
        site.delay = 10
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
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
                        # site.delay = delay

                        # site.delay = 2
                        boletos = site.verificarobjetoexiste('LINK_TEXT', 'Gerar', itemunico=False)
                        site.delay = delay

                        if len(boletos) > 0:
                            for i, boleto in enumerate(boletos, start=1):
                                # Clica no botão de boleto
                                boleto.click()
                                # Define que o nome do arquivo ficará
                                if i == 1:
                                    novonomearquivo = os.path.join(aux.caminho, aux.left(linha[identificador], 4) + '_' + info.retornaradministradora('nomereal', linha[Administradora],
                                                                                                                                                      'nomereduzido') + '.pdf')
                                else:
                                    novonomearquivo = os.path.join(aux.caminho, aux.left(linha[identificador], 4) + '_' + info.retornaradministradora('nomereal', linha[Administradora],
                                                                                                                                                      'nomereduzido') + '_' + str(i) + '.pdf')
                                time.sleep(1)

                                # Espera o download finalizar e "pega" o arquivo baixado (Espera o download pela página download do chrome)
                                arquivobaixado = os.path.join(objeto.pastadownload, site.pegaarquivobaixado(tempoesperadownload, 1))

                                # Verifica se o arquivo baixado de fato existe
                                if os.path.isfile(arquivobaixado):
                                    # Move o arquivo para o caminho escolhido adicionando o cabeçalho
                                    aux.adicionarcabecalhopdf(objeto.arquivobaixado, novonomearquivo, aux.left(linha[identificador], 4))
                                    # Verifica se o arquivo foi gerado
                                    if os.path.isfile(novonomearquivo):
                                        numboleto += 1
                                        time.sleep(1)

                        # Retorna resposta na linha
                        linha[Resposta] = respostapadrao(numboleto)

                        if site is not None:
                            site.fecharsite()
                    else:
                        # Retorna a mensagem de erro do site
                        linha[Resposta] = re.sub('\n+', '\n', mensagemerro.text)
                        linha[ProblemaLogin] = True

    except Exception as e:
        linha[Resposta] = str(e)
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = __name__
        if site is not None:
            site.fecharsite()


def bcf(linha):
    site = None
    try:
        # Variável que vai retornar a quantidade de boletos
        numboleto = 0
        # Prepara o objeto
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site

        if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
            # Abre o browser
            site.abrirnavegador()
        if site is not None and site.navegador != -1:
            # Pega o delay configurado no objeto "Site"
            delay = site.delay
            objetos = site.verificarobjetoexiste('CLASS_NAME', 'text', itemunico=False)
            for objeto in objetos:
                match objeto.get_attribute('name'):
                    # Campo de usuário
                    case 'login':
                        objeto.clear()
                        objeto.send_keys(linha[Usuario])
                    case 'senha':
                        objeto.clear()
                        objeto.send_keys(linha[Senha])

            # Botão de Login
            botao = site.verificarobjetoexiste('CLASS_NAME', 'submit')
            # Verifica se achou o botão de ‘login’
            if botao is not None:
                # Clica no botão de ‘login’
                site.navegador.execute_script("arguments[0].click()", botao)
                time.sleep(1)
                # Diminui o tempo de busca pela mensagem de erro
                site.delay = 2
                # Mensagem de erro (Teste de usuário ou senha inválido)
                mensagemerro = site.verificarobjetoexiste('ID', 'ucLoginSistema_lbErroEntrar')
                if mensagemerro is None:
                    # Segunda tela de erro de login (o site tem duas telas de erro possíveis)
                    mensagemerro = site.verificarobjetoexiste('ID', 'errorLogin')

                # Teste de erro depois de login validado
                testetelalinks = site.verificarobjetoexiste('PARTIAL_LINK_TEXT', 'via de boleto')
                if mensagemerro is None:
                    mensagemerro = site.verificarobjetoexiste('CLASS_NAME', 'alert')

                # Retorna o delay ao normal
                site.delay = delay
                # Teste se não tem mensagem de erro do site
                if mensagemerro is None or testetelalinks is not None:
                    numboleto = 0
                    site.navegador.execute_script("arguments[0].click()", testetelalinks)
                    time.sleep(1)
                    if site.verificarobjetoexiste('CLASS_NAME', 'tableRecibos', sotestar=True):
                        links = site.verificarobjetoexiste('CSS_SELECTOR', "a [title='Baixar PDF']")
                        if links is not None:
                            for link in links:
                                # Clica no link
                                site.navegador.execute_script("arguments[0].click()", link)
                                # Espera o download finalizar
                                site.esperadownloads(objeto.pastadownload, timeout)
                                # Define o nome (baseado no nr do boleto)
                                if numboleto == 0:
                                    novonomearquivo = objeto.pastadownload + '\\' + linha[identificador] + '_' + linha[Administradora] + '.pdf'
                                else:
                                    novonomearquivo = objeto.pastadownload + '\\' + linha[identificador] + '_' + linha[Administradora] + '_' \
                                                      + str(numboleto) + '.pdf'

                                # Espera o download finalizar e "pega" o arquivo baixado
                                arquivobaixado = site.pegaarquivobaixado(tempoesperadownload)

                                # Verifica se o arquivo baixado de fato existe
                                if os.path.isfile(objeto.pastadownload + '\\' + arquivobaixado):
                                    aux.renomeararquivo(objeto.pastadownload + '\\' + arquivobaixado, novonomearquivo)
                                    if os.path.isfile(novonomearquivo):
                                        numboleto += 1
                                    time.sleep(1)

                    # Retorna resposta na linha
                    linha[Resposta] = respostapadrao(numboleto)
                else:
                    # Retorna resposta na linha
                    linha[Resposta] = re.sub('\n+', '\n', mensagemerro.text)
                    linha[ProblemaLogin] = True

                # Fecha o browser
                site.fecharsite()

    except Exception as e:
        linha[Resposta] = str(e)
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = __name__
        if site is not None:
            site.fecharsite()


def cipa(objeto, linha):
    # Iniciar o trabalho, só copiado de outro
    site = None
    # try:
    # Variável que vai retornar a quantidade de boletos

    numboleto = 0
    boletosanalises = 0
    achouapartamento = False
    mudoustatus = False
    textoobjetologado = '/html/body/app-root/ion-content/layout/app-layout/div/div[1]/div/div/div/div[2]/user-menu/div/div/div/div'
    # apartamentobuscado, blocobuscado = linha[Apartamento].split('/')
    # apartamentobuscado = apartamentobuscado.strip()
    # blocobuscado = blocobuscado.strip()

    # Prepara o objeto
    site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
    site.delay = 30
    # Abre o browser
    site.abrirnavegador()
    # Verifica se iniciou o site
    if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
        # Caso esteja com outra página aberta, fecha
        if site is not None:
            site.fecharsite()
        # Inicia o browser
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
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
                    btnsalvar = site.verificarobjetoexiste('XPATH', '/html/body/app-root/ion-content/layout/empty-layout/div/div/auth-sign-in/div/div[2]/div/form/button')
                    # Verifica se achou o botão Salvar
                    if btnsalvar is not None:
                        site.navegador.execute_script("arguments[0].click()", btnsalvar)
                        site.delay = 10
                        # Verifica se deu erro após apertar no botão para realizar o "LOGIN"
                        msgerro = site.verificarobjetoexiste('ID', 'divMsgErroArea')
                        site.delay = delay
                        if msgerro is None:
                            # Operação normal
                            site.delay = 10
                            usuariologado = site.verificarobjetoexiste('XPATH', textoobjetologado)
                            textousuariologado = usuariologado.text
                            textousuariologado = textousuariologado.upper()
                            site.delay = delay
                            if usuariologado is not None:
                                if 'WMARTINS' in textousuariologado or 'ELENICE' in textousuariologado:
                                    site.delay = 10
                                    # Botão para se o usuário estiver em um condomínio específico volta para lista de condomínios
                                    botaolistacondominios = site.verificarobjetoexiste('XPATH', '/html/body/app-root/ion-content/layout/app-layout/div/div[1]/div/div[2]/div/div/div[2]/mat-icon')
                                    site.delay = delay
                                    if botaolistacondominios is not None:
                                        site.navegador.execute_script("arguments[0].click()", botaolistacondominios)

                                    time.sleep(1)
                                    # Verifica se está na página de lista de condomínios
                                    if site.verificarobjetoexiste('CSS_SELECTOR', "condominium-list", sotestar=True):
                                        # Aciona a barra de pesquisa para verificar se o condomínio dado como entrada está na lista de condomínios do usuário
                                        barrapesquisa = site.verificarobjetoexiste('ID', "mat-input-2")
                                        time.sleep(2)
                                        if barrapesquisa is not None:
                                            barrapesquisa.send_keys(linha[Condominio])
                                            botaopesquisa = site.verificarobjetoexiste('XPATH',
                                                                                       "/html/body/app-root/ion-content/layout/app-layout/div/div[2]/div/condominium-list/div/div/div[2]/div/button/span[1]")
                                            if botaopesquisa is not None:
                                                site.navegador.execute_script("arguments[0].click()", botaopesquisa)
                                                time.sleep(3)
                                                icondominio = site.verificarobjetoexiste('XPATH',
                                                                                         '/html/body/app-root/ion-content/layout/app-layout/div/div[2]/div/condominium-list/div/div/div[3]/div/div/div')
                                                if icondominio is not None:
                                                    # Entra no condomínio selecionado
                                                    site.navegador.execute_script("arguments[0].click()", icondominio)
                                                    time.sleep(2)
                                                    # Botão de expansão de Menu de Financeiro
                                                    financeiro = site.verificarobjetoexiste('XPATH',
                                                                                            '/html/body/app-root/ion-content/layout/app-layout/div/div[1]/div/div[2]/div/cipafacil-horizontal-navigation/div/cipafacil-horizontal-navigation-branch-item[1]/div/div/div/div/div/span')
                                                    if financeiro is not None:
                                                        # Expande o Menu
                                                        site.navegador.execute_script("arguments[0].click()", financeiro)
                                                        time.sleep(1)
                                                        # Item de Boleto
                                                        itensfinanceiro = site.verificarobjetoexiste('CLASS_NAME', 'cipafacil-horizontal-navigation-item-title', itemunico=False)
                                                        for item in itensfinanceiro:
                                                            if item.text == 'Boleto':
                                                                site.navegador.execute_script("arguments[0].click()", item)
                                                                time.sleep(1)
                                                                botaofiltro = site.verificarobjetoexiste('XPATH',
                                                                                                         '/html/body/app-root/ion-content/layout/app-layout/div/div[2]/div/boleto-list/div/div/div[1]/div[2]/button/span[1]')
                                                                if botaofiltro is not None:
                                                                    site.navegador.execute_script("arguments[0].click()", botaofiltro)
                                                                    time.sleep(1)
                                                                    filtrounidade = site.verificarobjetoexiste('XPATH',
                                                                                                               '/html/body/div[2]/div[2]/div/mat-dialog-container/boleto-filter-dialog/div/div[2]/mat-form-field[1]/div/div[1]/div')

                                                                    if filtrounidade is not None:
                                                                        site.navegador.execute_script("arguments[0].click()", filtrounidade)
                                                                        # listaunidade = site.verificarobjetoexiste('ID', 'mat-autocomplete-0')
                                                                        listaapartamentos = site.verificarobjetoexiste('CLASS_NAME', 'mat-option-text', itemunico=False)
                                                                        if listaapartamentos is not None:
                                                                            for apartamento in listaapartamentos:
                                                                                if linha[Apartamento] in apartamento.text:
                                                                                    # site.navegador.execute_script("arguments[0].click()", filtrounidade)
                                                                                    site.navegador.execute_script("arguments[0].click()", apartamento)
                                                                                    achouapartamento = True
                                                                                    break
                                                                        if achouapartamento:
                                                                            filtrostatus = site.verificarobjetoexiste('XPATH',
                                                                                                                      '/html/body/div[2]/div[2]/div/mat-dialog-container/boleto-filter-dialog/div/div[2]/mat-form-field[2]/div/div[1]/div')
                                                                            if filtrostatus is not None:
                                                                                site.navegador.execute_script("arguments[0].click()", filtrostatus)
                                                                                listastatus = site.verificarobjetoexiste('CLASS_NAME', 'mat-option-text', itemunico=False)

                                                                                for status in listastatus:
                                                                                    if 'Em Aberto' in status.text:
                                                                                        site.navegador.execute_script("arguments[0].click()", status)
                                                                                        mudoustatus = True

                                                                    if achouapartamento or mudoustatus:

                                                                        botaoaplicafiltro = site.verificarobjetoexiste('XPATH',
                                                                                                                       '/html/body/div[2]/div[2]/div/mat-dialog-container/boleto-filter-dialog/div/div[3]/button[2]/span[1]')
                                                                        if botaoaplicafiltro is not None:
                                                                            site.navegador.execute_script("arguments[0].click()", botaoaplicafiltro)
                                                                            break
                                                                    else:
                                                                        botaofecharfiltro = site.verificarobjetoexiste('XPATH',
                                                                                                                       '/html/body/div[2]/div[2]/div/mat-dialog-container/boleto-filter-dialog/div/div[1]/button/span[1]/mat-icon/svg')
                                                                        if botaofecharfiltro is not None:
                                                                            site.navegador.execute_script("arguments[0].click()", botaofecharfiltro)
                                                                            break

                                        # CONTINUAR DAQUI, VERIFICAR SE TEM BOLETO E SE TIVER FAZER O DOWNLOAD

                                        # Verifica se tem a mensagem de que não tem boleto disponível
                                        if not site.verificarobjetoexiste('CLASS_NAME', 'conteudo-nenhuma-cobranca', sotestar=True):
                                            if site.verificarobjetoexiste('CLASS_NAME', 'numero', sotestar=True):
                                                numapartamentos = site.verificarobjetoexiste('CLASS_NAME', 'numero', itemunico=False)
                                                dadosapartamentos = site.verificarobjetoexiste('CSS_SELECTOR', "[class='infos col-md-6']", itemunico=False)
                                                complementoapartamentos = site.verificarobjetoexiste('CLASS_NAME', "complemento", itemunico=False)
                                                listaapartamentos = zip(numapartamentos, dadosapartamentos, complementoapartamentos)
                                                for indice, linhaweb in enumerate(listaapartamentos):
                                                    # Garante que só o exibido(no caso o "a vencer") seja testado
                                                    numeroap, dadosap, complementoap = linhaweb
                                                    # Vai para a tela inicial
                                                    site.irparaaba(titulo='Areadocondomino')
                                                    if numeroap.is_displayed():
                                                        if numeroap.text == apartamentobuscado and complementoap.text == blocobuscado:
                                                            achouapartamento = True
                                                            if "Indisponível" not in dadosap.text:
                                                                site.navegador.execute_script("arguments[0].click()", numeroap)
                                                                # Seleciona a opção na janela "popup" que aparece
                                                                while not site.verificarobjetoexiste('ID', 'salvar', sotestar=True):
                                                                    time.sleep(1)
                                                                btnsalvar = site.verificarobjetoexiste('ID', 'salvar')
                                                                site.navegador.execute_script("arguments[0].click()", btnsalvar)
                                                                time.sleep(1)
                                                                # Vai para a tela inicial
                                                                site.irparaaba(titulo='Areadocondomino')

                                                                # Vai para segunda página que ele abre com o boleto resumido
                                                                site.irparaaba(indice=site.num_abas())

                                                                achoupagina = ('SEGUNDAVIA' in site.navegador.current_url.upper())

                                                                if achoupagina:
                                                                    # Aperta no botão para gerar o boleto
                                                                    btngerarboleto = site.verificarobjetoexiste('CSS_SELECTOR', "[class='botao pagarBoleto pagarBoletoParcelado']")
                                                                    if btngerarboleto is not None:
                                                                        site.navegador.execute_script("arguments[0].click()", btngerarboleto)
                                                                        btngerarboleto = site.verificarobjetoexiste('ID', 'btnSubmitParcelamentoCartao')
                                                                        while btngerarboleto is None:
                                                                            btngerarboleto = site.verificarobjetoexiste('ID', 'btnSubmitParcelamentoCartao')
                                                                            time.sleep(1)

                                                                        while not btngerarboleto.is_displayed():
                                                                            time.sleep(1)

                                                                        site.navegador.execute_script("arguments[0].click()", btngerarboleto)

                                                                        # Espera o download finalizar
                                                                        site.esperadownloads(objeto.pastadownload, timeout)

                                                                        # Define o nome (baseado no nr do boleto)
                                                                        if numboleto == 0:
                                                                            novonomearquivo = objeto.pastadownload + '\\' + linha[identificador] + '_' + \
                                                                                              info.retornaradministradora('nomereal', linha[Administradora], 'nomereduzido') + '.pdf'
                                                                        else:
                                                                            novonomearquivo = objeto.pastadownload + '\\' + linha[identificador] + '_' + \
                                                                                              info.retornaradministradora('nomereal', linha[Administradora], 'nomereduzido') + '_' \
                                                                                              + str(numboleto) + '.pdf'

                                                                        # Espera o download finalizar e "pega" o arquivo baixado
                                                                        arquivobaixado = site.pegaarquivobaixado(tempoesperadownload)

                                                                        # Verifica se o arquivo baixado de fato existe
                                                                        if os.path.isfile(objeto.pastadownload + '\\' + arquivobaixado):
                                                                            aux.renomeararquivo(objeto.pastadownload + '\\' + arquivobaixado, novonomearquivo)
                                                                            if os.path.isfile(novonomearquivo):
                                                                                numboleto += 1
                                                                            time.sleep(1)

                                                                else:
                                                                    boletosanalises += 1
                                                site.navegador.close()
                                        else:
                                            # Não existem boletos em aberto para buscar na lista do apartamento dado como entrada
                                            linha[Resposta] = respostapadrao(-1)
                                else:
                                    linha[Resposta] = "Condomínio não encontrado!"

                            if len(linha[Resposta]) == 0:
                                if not achouapartamento:
                                    linha[Resposta] = 'Apartamento não encontrado!'
                                else:
                                    linha[Resposta] = respostapadrao(numboleto, boletosanalises)
                        else:
                            # Mensagem de erro de Login
                            linha[Resposta] = msgerro.text
                            linha[ProblemaLogin] = True

    # except Exception as e:
    #     linha[Resposta] = str(e)
    #     linha[CheckErro] = True
    #     linha[Nomefuncao] = __name__
    #
    # finally:
    #     if site is not None:
    #         site.fecharsite()


def ICondo(objeto, linha):
    site = None
    # try:
    # Variável que vai retornar a quantidade de boletos
    numboleto = 0
    atualizacaocadastral = False
    # Retorna o texto do site
    textosite = "https://" + info.retornaradministradora('nomereal', linha[Administradora], 'nomereduzido') + ".icondo.com.br/"
    textosite = textosite.lower()
    # Prepara o objeto
    site = web.TratarSite(textosite, info.nomeprofilecond)
    # Abre o browser
    site.abrirnavegador()
    # Verifica se iniciou o site
    if site.url != textosite or site is None:
        # Caso esteja com outra página aberta, fecha
        if site is not None:
            site.fecharsite()
        # Inicia o browser
        site = web.TratarSite(textosite, info.nomeprofilecond)
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
                    botaoatualizacao = site.verificarobjetoexiste('CSS_SELECTOR', "[class = 'swal2-confirm btn green']")
                    if botaoatualizacao is not None:
                        botaoatualizacao.click()
                        # time.sleep(1)
                    else:
                        botaoatualizacao = site.verificarobjetoexiste('CSS_SELECTOR', "[class = 'btn green']")
                        if botaoatualizacao is not None:
                            botaoatualizacao.click()
                            # time.sleep(1)
                        else:
                            botaoatualizacao = site.verificarobjetoexiste('CSS_SELECTOR', "[class = 'swal2-confirm btn green']")
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
                    botao2via = site.verificarobjetoexiste('XPATH', '//a[contains(@href,"?arquivo=segunda-via-boleto")]')
                    if botao2via is not None:
                        botao2via.click()
                        time.sleep(3)
                        # Objeto de Mensagem sem Boleto
                        site.delay = 2
                        if site.verificarobjetoexiste('CLASS_NAME', 'alert-danger', sotestar=True):
                            linha[Resposta] = respostapadrao(numboleto)
                        else:
                            site.delay = delay
                            objetosboletos = site.verificarobjetoexiste('PARTIAL_LINK_TEXT', 'Recibo No.', itemunico=False)  # Debug
                            for i, boleto in enumerate(objetosboletos, start=1):
                                boleto.click()
                                time.sleep(1)

                                if i == 1:
                                    novonomearquivo = os.path.join(aux.caminho, aux.left(linha[identificador], 4) + '_' + info.retornaradministradora('nomereal', linha[Administradora],
                                                                                                                                                      'nomereduzido') + '.pdf')
                                else:
                                    novonomearquivo = os.path.join(aux.caminho, aux.left(linha[identificador], 4) + '_' + info.retornaradministradora('nomereal', linha[Administradora],
                                                                                                                                                      'nomereduzido') + '_' + str(i) + '.pdf')

                                linkimpressao = site.verificarobjetoexiste('LINK_TEXT', 'Imprimir/salvar pdf')
                                if linkimpressao is not None:
                                    linkimpressao.click()

                                    # Espera o download finalizar e "pega" o arquivo baixado
                                    arquivobaixado = site.pegaarquivobaixado(tempoesperadownload, caminhobaixado=objeto.pastadownload)

                                    # Verifica se o arquivo baixado de fato existe
                                    if os.path.isfile(objeto.pastadownload + '\\' + arquivobaixado):
                                        # aux.renomeararquivo(pastadownload + '\\' + arquivobaixado, novonomearquivo)
                                        aux.adicionarcabecalhopdf(objeto.pastadownload + '\\' + arquivobaixado, novonomearquivo, aux.left(linha[identificador], 4))
                                        if os.path.isfile(novonomearquivo):
                                            numboleto += 1
                                        time.sleep(1)

                time.sleep(1)

                # Retorna resposta na linha
                linha[Resposta] = respostapadrao(numboleto)
            else:
                # Retorna resposta na linha
                linha[Resposta] = re.sub('\n+', '\n', mensagemerro.text)
                linha[ProblemaLogin] = True

            # Fecha o browser
            site.fecharsite()

    # except Exception as e:
    #     linha[Resposta] = str(e)
    #     linha[CheckErro] = True
    #
    # finally:
    #     linha[Nomefuncao] = __name__
    #     if site is not None:
    #         site.fecharsite()


def Webware(objeto, linha):
    site = None
    try:
        numboleto = 0
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
            # Prepara o objeto
            site = web.TratarSite(textosite, info.nomeprofilecond)
            # Abre o browser
            site.abrirnavegador()
            # Verifica se iniciou o site
            if site.url != textosite or site is None:
                # Caso esteja com outra página aberta, fecha
                if site is not None:
                    site.fecharsite()
                # Inicia o browser
                site = web.TratarSite(textosite, info.nomeprofilecond)
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

                    site.delay = 2
                    # Mensagem de erro
                    erro = site.verificarobjetoexiste('CLASS_NAME', 'mensagem-erro')
                    site.delay = delay
                    if erro is None:
                        # Botão de segunda via
                        botaosegundavia = site.verificarobjetoexiste('CLASS_NAME', 'statistic__item ')

                        # Verifica se existe o botão de segunda via
                        if botaosegundavia is not None:
                            botaosegundavia.click()
                            # Pega o objeto do frame
                            janelaboleto = site.verificarobjetoexiste('CLASS_NAME', 'prestacao-interativa')
                            if janelaboleto is not None:
                                # Mudar para o frame dos boletos
                                site.irparaframe(janelaboleto)

                                # Verifica se tem e está visível o ícone de carregando no frame
                                telacarregando = site.verificarobjetoexiste('CLASS_NAME', 'carregando')
                                if telacarregando is not None:
                                    while telacarregando.is_displayed():
                                        time.sleep(1)

                                # Verifica se acha botao de alerta
                                botaoalerta = site.verificarobjetoexiste('CLASS_NAME', 'popup-alerta-botoes')
                                time.sleep(2)

                                # Verifica se tem o botão de aceitar o alerta
                                if botaoalerta is not None:
                                    # Clica no botão de aceitar o alerta
                                    botaoalerta.click()

                                # Pega todos os botões de gerar boleto do frame
                                objetosboletos = site.verificarobjetoexiste('XPATH', '//a[contains(@href,"/relatorioboleto/BuscarBoletoPorRecibo?")]', itemunico=False)  # Debug
                                for i, boleto in enumerate(objetosboletos, start=1):
                                    # Clica no botão de boleto
                                    boleto.click()
                                    # Define que o nome do arquivo ficará
                                    if i == 1:
                                        novonomearquivo = os.path.join(aux.caminho,
                                                                       aux.left(linha[identificador], 4) + '_' + info.retornaradministradora('nomereal', linha[Administradora],
                                                                                                                                             'nomereduzido') + '.pdf')
                                    else:
                                        novonomearquivo = os.path.join(aux.caminho,
                                                                       aux.left(linha[identificador], 4) + '_' + info.retornaradministradora('nomereal', linha[Administradora],
                                                                                                                                             'nomereduzido') + '_' + str(
                                                                           i) + '.pdf')

                                    # Espera o download finalizar e "pega" o arquivo baixado
                                    arquivobaixado = site.pegaarquivobaixado(tempoesperadownload, 1)

                                    # Verifica se o arquivo baixado de fato existe
                                    if os.path.isfile(objeto.pastadownload + '\\' + arquivobaixado):
                                        # Move o arquivo para o caminho escolhido
                                        if info.retornaradministradora('nomereal', linha[Administradora], 'nomereduzido') == "LOWNDES":
                                            aux.adicionarcabecalhopdf(objeto.pastadownload + '\\' + arquivobaixado, novonomearquivo, aux.left(linha[identificador], 4), protegido=True)
                                        else:
                                            aux.adicionarcabecalhopdf(objeto.pastadownload + '\\' + arquivobaixado, novonomearquivo, aux.left(linha[identificador], 4))
                                        # Verifica se o arquivo foi gerado
                                        if os.path.isfile(novonomearquivo):
                                            numboleto += 1
                                        time.sleep(1)

                                    # Volta pra aba original
                                    site.irparaaba(1)
                                    # Volta pro frame
                                    site.irparaframe(janelaboleto)

                                # Retorna resposta na linha
                                linha[Resposta] = respostapadrao(numboleto)
                    else:
                        # Retorna a tela de erro
                        linha[Resposta] = erro.text
                        # Retorna que não conseguiu logar
                        linha[ProblemaLogin] = True

    except Exception as e:
        linha[Resposta] = str(e)
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = __name__
        if site is not None:
            site.fecharsite()


def livefacilities(objeto, linha):
    site = None
    try:
        numboleto = 0
        match info.retornaradministradora('nomereal', linha[Administradora], 'nomereduzido'):
            case "HFLEX":
                sigla = "sys"
                dataid = '375'

            case "VORTEX":
                sigla = "vortex"
                dataid = '199'

            case _:
                sigla = ''

        # Abre o site dependendo da administradora
        if len(sigla) > 0:
            # Retorna o texto do site
            textosite = "http://%s.livefacilities.com.br/Index.aspx" % sigla
            textosite = textosite.lower()
            # Prepara o objeto
            site = web.TratarSite(textosite, info.nomeprofilecond)
            # Abre o browser
            site.abrirnavegador()
            # Verifica se iniciou o site
            if site.url != textosite or site is None:
                # Caso esteja com outra página aberta, fecha
                if site is not None:
                    site.fecharsite()
                # Inicia o browser
                site = web.TratarSite(textosite, info.nomeprofilecond)
                # Abre o browser
                site.abrirnavegador()
            if site is not None and site.navegador != -1:
                # Pega o delay configurado no objeto "Site"
                delay = site.delay
                # Botão de "Entrar"
                botao = site.verificarobjetoexiste('ID', "btLogin")
                # Verifica se achou o botão de ‘login’
                if botao is not None:
                    botao.click()
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
                            # Botão Financeiro
                            financeiro = site.verificarobjetoexiste('XPATH', '//a[contains(@data-idmenu,"%s")]' % dataid)
                            if financeiro is not None:
                                financeiro.click()
                                # Botão de segunda via
                                botaosegundavia = site.verificarobjetoexiste('XPATH', '//a[contains(@href,"/boletoLista.aspx?")]')

                                # Verifica se existe o botão de segunda via
                                if botaosegundavia is not None:
                                    # Clica no botão de segunda via
                                    botaosegundavia.click()
                                    site.delay = 2

                                    # Pega o objeto do frame
                                    erroboleto = site.verificarobjetoexiste('ID', 'ContentBody_ucBoletoLista_pListaErro')
                                    site.delay = delay
                                    if erroboleto is None:
                                        # Pega todos os botões de gerar boleto do frame
                                        objetosboletos = site.verificarobjetoexiste('XPATH', '//a[contains(@title,"Abrir Boleto")]', itemunico=False)
                                        totalboletos = len(objetosboletos)
                                        for i in range(totalboletos):
                                            # Clica no botão de boleto
                                            objetosboletos[i].click()
                                            time.sleep(2)
                                            frame = site.verificarobjetoexiste('CLASS_NAME', 'iziModal-iframe')
                                            if frame is not None:
                                                site.irparaframe(frame)
                                                # Botão Imprimir
                                                botaoimprimir = site.verificarobjetoexiste('ID', 'print')
                                                if botaoimprimir is not None:
                                                    time.sleep(2)
                                                    # Aperta o botão de impressão
                                                    botaoimprimir.click()
                                                    # Define que o nome do arquivo ficará
                                                    if i == 0:
                                                        novonomearquivo = os.path.join(aux.caminho,
                                                                                       aux.left(linha[identificador], 4) + '_' + info.retornaradministradora('nomereal', linha[Administradora],
                                                                                                                                                             'nomereduzido') + '.pdf')
                                                    else:
                                                        novonomearquivo = os.path.join(aux.caminho,
                                                                                       aux.left(linha[identificador], 4) + '_' + info.retornaradministradora('nomereal', linha[Administradora],
                                                                                                                                                             'nomereduzido') + '_' + str(
                                                                                           i + 1) + '.pdf')
                                                    time.sleep(2)
                                                    # Espera o download finalizar
                                                    site.esperadownloads(objeto.pastadownload, timeout)
                                                    # Espera o download finalizar e "pega" o arquivo baixado
                                                    arquivobaixado = aux.ultimoarquivo(objeto.pastadownload, 'pdf')

                                                    # Verifica se o arquivo baixado de fato existe
                                                    if os.path.isfile(arquivobaixado):
                                                        # Move o arquivo para o caminho escolhido
                                                        aux.adicionarcabecalhopdf(arquivobaixado, novonomearquivo, aux.left(linha[identificador], 4))
                                                        # Verifica se o arquivo foi gerado
                                                        if os.path.isfile(novonomearquivo):
                                                            numboleto += 1

                                            site.navegador.refresh()
                                            time.sleep(2)
                                            objetosboletos = site.verificarobjetoexiste('XPATH', '//a[contains(@title,"Abrir Boleto")]', itemunico=False)

                                        # Retorna resposta na linha
                                        linha[Resposta] = respostapadrao(numboleto)
                                    else:
                                        # Retorna a tela de erro
                                        linha[Resposta] = respostapadrao(numboleto)
                        else:
                            # Retorna a tela de erro
                            linha[Resposta] = erro.text
                            # Retorna que não conseguiu logar
                            linha[ProblemaLogin] = True

    except Exception as e:
        linha[Resposta] = str(e)
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = __name__
        if site is not None:
            site.fecharsite()


def imodata(objeto, linha):
    site = None
    try:
        numboleto = 0
        # Prepara o objeto
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
        site.delay = 10
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
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
                                # Caso tenha logado
                                # Pedir cadastro do CPF
                                definircpf = site.verificarobjetoexiste('ID', 'formCPF')
                                if definircpf is not None:
                                    fechar = site.verificarobjetoexiste('XPATH', '/html/body/div[7]/div/div/div[1]/button')
                                    if fechar is not None:
                                        fechar.click()

                                site.delay = delay
                                # Botão de 2.ª via
                                botao2via = site.verificarobjetoexiste('PARTIAL_LINK_TEXT', '2a.Via')
                                botao2via.click()

                                frame = site.verificarobjetoexiste('NAME', 'MyFrame')
                                site.irparaframe(frame)
                                boletos = site.verificarobjetoexiste('PARTIAL_LINK_TEXT', 'Clique aqui para gerar a 2ª via', itemunico=False)
                                # boletos = site.verificarobjetoexiste('XPATH', '//a[contains(@href,"/condo/2via/econdo_bolmenu_v4.asp?")]', itemunico=False)

                                if boletos is not None:
                                    for i, boleto in enumerate(boletos, start=1):
                                        boleto = boletos[i - 1]
                                        # Clica no botão de boleto
                                        boleto.click()
                                        # Busca o ícone do boleto
                                        iconeboleto = site.verificarobjetoexiste('CLASS_NAME', 'boleto')
                                        if iconeboleto is not None:
                                            # Clica no ícone do boleto
                                            site.navegador.execute_script("arguments[0].click()", iconeboleto)

                                        # Define que o nome do arquivo ficará
                                        if i == 1:
                                            novonomearquivo = os.path.join(aux.caminho, aux.left(linha[identificador], 4) + '_' + info.retornaradministradora('nomereal', linha[Administradora],
                                                                                                                                                              'nomereduzido') + '.pdf')
                                        else:
                                            novonomearquivo = os.path.join(aux.caminho, aux.left(linha[identificador], 4) + '_' + info.retornaradministradora('nomereal', linha[Administradora],
                                                                                                                                                              'nomereduzido') + '_' + str(i) + '.pdf')

                                        time.sleep(2)
                                        # Verifica se o download acabou
                                        site.esperadownloads(objeto.pastadownload, timeout)
                                        # Espera o download finalizar e "pega" o arquivo baixado
                                        arquivobaixado = aux.ultimoarquivo(objeto.pastadownload, 'pdf')

                                        # Vai na nova janela e fecha
                                        achoujanela = site.irparaaba(titulo='Boleto Bradesco')
                                        time.sleep(1)
                                        if not achoujanela:
                                            achoujanela = site.irparaaba(titulo='Boleto Itaú')
                                        if achoujanela:
                                            site.navegador.close()
                                        # Verifica se o arquivo baixado de fato existe
                                        if os.path.isfile(arquivobaixado):
                                            # Move o arquivo para o caminho escolhido
                                            aux.adicionarcabecalhopdf(arquivobaixado, novonomearquivo, aux.left(linha[identificador], 4))
                                            # Verifica se o arquivo foi gerado
                                            if os.path.isfile(novonomearquivo):
                                                numboleto += 1
                                                time.sleep(1)

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
                                        boletos = site.verificarobjetoexiste('PARTIAL_LINK_TEXT', 'Clique aqui para gerar a 2ª via', itemunico=False)

                                # Retorna resposta na linha
                                linha[Resposta] = respostapadrao(numboleto)

                                if site is not None:
                                    site.fecharsite()
                            else:
                                # Retorna a mensagem de erro do site
                                linha[Resposta] = re.sub('\n+', '\n', mensagemerro.text)
                                # Retorna que não conseguiu logar
                                linha[ProblemaLogin] = True

    except Exception as e:
        linha[Resposta] = str(e)
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = __name__
        if site is not None:
            site.fecharsite()


def fernandoefernandes(objeto, linha):
    site = None
    try:
        numboleto = 0
        # Prepara o objeto
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
        site.delay = 10
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
            # Abre o browser
            site.abrirnavegador()
        if site is not None and site.navegador != -1:
            # Pega o delay configurado no objeto "Site"
            delay = site.delay

            # Botão de ‘login’
            botaologin = site.verificarobjetoexiste('XPATH', '/html/body/section[2]/div/form/div[4]/button')
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
                                time.sleep(1)
                                # Pega a lista de boletos
                                boletos = site.verificarobjetoexiste('CSS_SELECTOR', "[class='fa fa-download']", itemunico=False)
                                if boletos is not None:
                                    for i, boleto in enumerate(boletos, start=1):
                                        # Clica no botão de boleto
                                        boleto.click()
                                        # Define que o nome do arquivo ficará
                                        if i == 1:
                                            novonomearquivo = os.path.join(aux.caminho, aux.left(linha[identificador], 4) + '_' + info.retornaradministradora('nomereal', linha[Administradora],
                                                                                                                                                              'nomereduzido') + '.pdf')
                                        else:
                                            novonomearquivo = os.path.join(aux.caminho, aux.left(linha[identificador], 4) + '_' + info.retornaradministradora('nomereal', linha[Administradora],
                                                                                                                                                              'nomereduzido') + '_' + str(i) + '.pdf')
                                        time.sleep(1)

                                        # Espera o download finalizar e "pega" o arquivo baixado (Espera o download pela página download do chrome)
                                        arquivobaixado = os.path.join(objeto.pastadownload, site.pegaarquivobaixado(tempoesperadownload, 1))

                                        # Pega o último arquivo baixado da pasta
                                        # arquivobaixado = aux.ultimoarquivo(pastadownload, 'pdf')

                                        # Verifica se o arquivo baixado de fato existe
                                        if os.path.isfile(arquivobaixado):
                                            # Move o arquivo para o caminho escolhido adicionando o cabeçalho
                                            aux.adicionarcabecalhopdf(arquivobaixado, novonomearquivo, aux.left(linha[identificador], 4))
                                            # Verifica se o arquivo foi gerado
                                            if os.path.isfile(novonomearquivo):
                                                numboleto += 1
                                                time.sleep(1)
                                        site.irparaaba(1)

                        # Retorna resposta na linha
                        linha[Resposta] = respostapadrao(numboleto)

                    else:
                        # Retorna a mensagem de erro do site
                        linha[Resposta] = re.sub('\n+', '\n', mensagemerro.text)
                        # Retorna que não conseguiu logar
                        linha[ProblemaLogin] = True

    except Exception as e:
        linha[Resposta] = str(e)
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = __name__
        if site is not None:
            site.fecharsite()


def protel(objeto, linha):
    site = None
    try:
        mensagemerro = None
        numboleto = 0
        # Prepara o objeto
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
        site.delay = 10
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
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
                        boletos = site.verificarobjetoexiste('XPATH', "//a//img[contains(@src,'/images/icone_pdf.gif')]", itemunico=False)

                        if len(boletos) > 0:
                            for i, boleto in enumerate(boletos, start=1):
                                # Clica no botão de boleto
                                boleto.click()
                                # Define que o nome do arquivo ficará
                                if i == 1:
                                    novonomearquivo = os.path.join(aux.caminho, aux.left(linha[identificador], 4) + '_' + info.retornaradministradora('nomereal', linha[Administradora],
                                                                                                                                                      'nomereduzido') + '.pdf')
                                else:
                                    novonomearquivo = os.path.join(aux.caminho, aux.left(linha[identificador], 4) + '_' + info.retornaradministradora('nomereal', linha[Administradora],
                                                                                                                                                      'nomereduzido') + '_' + str(i) + '.pdf')
                                time.sleep(2)

                                # Espera o download finalizar e "pega" o arquivo baixado (Espera o download pela página download do chrome)
                                arquivobaixado = os.path.join(objeto.pastadownload, site.pegaarquivobaixado(tempoesperadownload, 1))

                                time.sleep(1)
                                # Pega o último arquivo baixado da pasta
                                # arquivobaixado = aux.ultimoarquivo(pastadownload, 'pdf')

                                # Verifica se o arquivo baixado de fato existe
                                if os.path.isfile(arquivobaixado):
                                    # Move o arquivo para o caminho escolhido adicionando o cabeçalho
                                    aux.adicionarcabecalhopdf(arquivobaixado, novonomearquivo, aux.left(linha[identificador], 4))
                                    # Verifica se o arquivo foi gerado
                                    if os.path.isfile(novonomearquivo):
                                        numboleto += 1
                                        time.sleep(1)

                        # Retorna resposta na linha
                        linha[Resposta] = respostapadrao(numboleto)

                        if site is not None:
                            site.fecharsite()
                    else:
                        # Retorna a mensagem de erro do site
                        linha[Resposta] = re.sub('\n+', '\n', mensagemerro.text)
                        # Retorna que não conseguiu logar
                        linha[ProblemaLogin] = True

    except Exception as e:
        linha[Resposta] = str(e)
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = __name__
        if site is not None:
            site.fecharsite()


def nacional(objeto, linha):
    site = None
    try:
        numboleto = 0
        # Prepara o objeto
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
        site.delay = 10
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
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
                    mensagemerro = site.verificarobjetoexiste('XPATH', '/html/body/div[2]/div/form/table/tbody/tr[2]/td/text()[1]')
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

                                # Vai para o frame de boletos
                                frameboletos = site.verificarobjetoexiste('NAME', 'principal')
                                if frameboletos is not None:
                                    # Vai para os frames do boleto
                                    site.irparaframe(frameboletos)

# SEM BOLETOS, NÃO TEM COMO TESTAR DAQUI PRA BAIXO
# ==============================================================================================================================================================
                                    boletos = site.verificarobjetoexiste('PARTIAL_LINK_TEXT', 'Clique aqui para gerar a 2ª via', itemunico=False)

                                    if boletos is not None:
                                        for i, boleto in enumerate(boletos, start=1):
                                            # boleto = boletos[i - 1]
                                            # Clica no botão de boleto
                                            boleto.click()
                                            # Define que o nome do arquivo ficará
                                            if i == 1:
                                                novonomearquivo = os.path.join(aux.caminho, aux.left(linha[identificador], 4) + '_' + info.retornaradministradora('nomereal', linha[Administradora],
                                                                                                                                                                  'nomereduzido') + '.pdf')
                                            else:
                                                novonomearquivo = os.path.join(aux.caminho, aux.left(linha[identificador], 4) + '_' + info.retornaradministradora('nomereal', linha[Administradora],
                                                                                                                                                                  'nomereduzido') + '_' + str(i) + '.pdf')
                                            time.sleep(2)

                                            # Espera o download finalizar e "pega" o arquivo baixado (Espera o download pela página download do chrome)
                                            arquivobaixado = os.path.join(objeto, objeto.pastadownload, site.pegaarquivobaixado(tempoesperadownload))

                                            # Espera o download finalizar
                                            # site.esperadownloads(pastadownload, timeout)
                                            # Pega o último arquivo baixado da pasta
                                            # arquivobaixado = aux.ultimoarquivo(pastadownload, 'pdf')

                                            # Verifica se o arquivo baixado de fato existe
                                            if os.path.isfile(arquivobaixado):
                                                # Move o arquivo para o caminho escolhido adicionando o cabeçalho
                                                aux.adicionarcabecalhopdf(arquivobaixado, novonomearquivo, aux.left(linha[identificador], 4))
                                                # Verifica se o arquivo foi gerado
                                                if os.path.isfile(novonomearquivo):
                                                    numboleto += 1
                                                    time.sleep(1)

                                    # Retorna resposta na linha
                                    linha[Resposta] = respostapadrao(numboleto)

                                    if site is not None:
                                        site.fecharsite()
                    else:
                        # Retorna a mensagem de erro do site
                        linha[Resposta] = re.sub('\n+', '\n', mensagemerro.text)
                        # Retorna que não conseguiu logar
                        linha[ProblemaLogin] = True

    except Exception as e:
        linha[Resposta] = str(e)
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = __name__
        if site is not None:
            site.fecharsite()


def modelopagina(objeto, linha):
    site = None
    try:
        numboleto = 0
        # Prepara o objeto
        site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
        site.delay = 10
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[Administradora], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.retornaradministradora('nomereal', linha[Administradora], 'site'), info.nomeprofilecond)
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
                        # Botão de 2.ª via
                        botao2via = site.verificarobjetoexiste('PARTIAL_LINK_TEXT', '2a.Via')
                        botao2via.click()

                        boletos = site.verificarobjetoexiste('PARTIAL_LINK_TEXT', 'Clique aqui para gerar a 2ª via', itemunico=False)

                        if boletos is not None:
                            for i, boleto in enumerate(boletos, start=1):
                                # boleto = boletos[i - 1]
                                # Clica no botão de boleto
                                boleto.click()
                                # Define que o nome do arquivo ficará
                                if i == 1:
                                    novonomearquivo = os.path.join(aux.caminho, aux.left(linha[identificador], 4) + '_' + info.retornaradministradora('nomereal', linha[Administradora],
                                                                                                                                                      'nomereduzido') + '.pdf')
                                else:
                                    novonomearquivo = os.path.join(aux.caminho, aux.left(linha[identificador], 4) + '_' + info.retornaradministradora('nomereal', linha[Administradora],
                                                                                                                                                      'nomereduzido') + '_' + str(i) + '.pdf')
                                time.sleep(2)

                                # Espera o download finalizar e "pega" o arquivo baixado (Espera o download pela página download do chrome)
                                arquivobaixado = os.path.join(objeto.pastadownload, site.pegaarquivobaixado(tempoesperadownload))

                                # Espera o download finalizar
                                # site.esperadownloads(objeto.pastadownload, timeout)
                                # Pega o último arquivo baixado da pasta
                                # arquivobaixado = aux.ultimoarquivo(objeto.pastadownload, 'pdf')

                                # Verifica se o arquivo baixado de fato existe
                                if os.path.isfile(arquivobaixado):
                                    # Move o arquivo para o caminho escolhido adicionando o cabeçalho
                                    aux.adicionarcabecalhopdf(arquivobaixado, novonomearquivo, aux.left(linha[identificador], 4))
                                    # Verifica se o arquivo foi gerado
                                    if os.path.isfile(novonomearquivo):
                                        numboleto += 1
                                        time.sleep(1)

                        # Retorna resposta na linha
                        linha[Resposta] = respostapadrao(numboleto)

                        if site is not None:
                            site.fecharsite()
                    else:
                        # Retorna a mensagem de erro do site
                        linha[Resposta] = re.sub('\n+', '\n', mensagemerro.text)
                        # Retorna que não conseguiu logar
                        linha[ProblemaLogin] = True

    except Exception as e:
        linha[Resposta] = str(e)
        linha[CheckErro] = True

    finally:
        linha[Nomefuncao] = __name__
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
                        resposta = resposta + " e tem " + str(quantanalise) + " boleto para ser analisado (provável em acordo)"
                    case _:
                        resposta = resposta + " e tem " + str(quantanalise) + " boletos para ser analisado (provável em acordo)"

            case _:
                resposta = 'Salvo ' + quant + ' boletos'
                match quantanalise:
                    case 1:
                        resposta = resposta + " e tem " + str(quantanalise) + " boleto para ser analisado (provável em acordo)"
                    case _:
                        resposta = resposta + " e tem " + str(quantanalise) + " boletos para ser analisado (provável em acordo)"

    return resposta
