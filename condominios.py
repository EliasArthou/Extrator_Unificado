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

timeout = 10
pastadownload = aux.caminhoprojeto() + '\\' + 'Downloads'
tempoesperadownload = 180
mensagemerropadrao = 'Deu erro! Tentar novamente!'
mensagemsemcondominio = "Condomínio não encontrado!"


def abrj(linha):
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
                                                                            site.esperadownloads(pastadownload, timeout)

                                                                            # Define o nome (baseado no nr do boleto)
                                                                            if numboleto == 0:
                                                                                novonomearquivo = pastadownload + '\\' + linha[identificador] + '_' + \
                                                                                                  info.retornaradministradora('nomereal', linha[Administradora], 'nomereduzido') + '.pdf'
                                                                            else:
                                                                                novonomearquivo = pastadownload + '\\' + linha[identificador] + '_' + \
                                                                                                  info.retornaradministradora('nomereal', linha[Administradora], 'nomereduzido') + '_' \
                                                                                                  + str(numboleto) + '.pdf'

                                                                            # Retorna o último arquivo na pasta de download
                                                                            lastfile = aux.ultimoarquivo(pastadownload, '.pdf')
                                                                            # Verifica se achou um arquivo na pasta de download
                                                                            if len(lastfile) > 0:
                                                                                # Nâo chamo adicionando o cabeçalho porque o arquivo vem protegido
                                                                                aux.renomeararquivo(lastfile, novonomearquivo)
                                                                                numboleto += 1

                                                                    else:
                                                                        boletosanalises += 1
                                                    site.navegador.close()
                                            else:
                                                # Não existem boletos em aberto para buscar na lista do apartamento dado como entrada
                                                linha[Resposta] = respostaducessopadrao(-1)
                                    else:
                                        linha[Resposta] = "Condomínio não encontrado!"

                                if len(linha[Resposta]) == 0:
                                    if not achouapartamento:
                                        linha[Resposta] = 'Apartamento não encontrado!'
                                    else:
                                        linha[Resposta] = respostaducessopadrao(numboleto, boletosanalises)
                            else:
                                # Mensagem de erro de Login
                                linha[Resposta] = msgerro.text

    except Exception as e:
        if site is not None:
            site.fecharsite()
        linha[Resposta] = str(e)


def apsa(linha):
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
                                                                                  '/html/body/app-root/ion-app/ion-router-outlet/app-layout/ion-app/div[2]/section/ion-router-outlet/app-index/div/home-desktop/div/div[2]/div[1]/div/app-cotas-com-vencimento-no-mes/a/ds-button/button/div/div')
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
                                                                    novonomearquivo = pastadownload + '\\' + linha[identificador] + '_' + linha[Administradora] + '.pdf'
                                                                else:
                                                                    novonomearquivo = pastadownload + '\\' + linha[identificador] + '_' + linha[Administradora] + '_' + str(numboleto + 1) + '.pdf'

                                                                # Espera o download finalizar e "pega" o arquivo baixado
                                                                arquivobaixado = site.pegaarquivobaixado(tempoesperadownload)
                                                                time.sleep(1)

                                                                # Verifica se o arquivo baixado de fato existe
                                                                if os.path.isfile(pastadownload + '\\' + arquivobaixado):
                                                                    # Renomeia o arquivo baixado para o código de cliente
                                                                    aux.adicionarcabecalhopdf(pastadownload + '\\' + arquivobaixado, novonomearquivo, linha[identificador])
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
                                        linha[Resposta] = respostaducessopadrao(numboleto)
                                        break

                        else:
                            if hasattr(mensagemerro, 'text'):
                                linha[Resposta] = mensagemerro.text
                            else:
                                linha[Resposta] = mensagemerropadrao

                        # Fecha o browser
                        site.fecharsite()

    except Exception as e:
        linha[Resposta] = str(e)

    finally:
        if site is not None:
            site.fecharsite()


def bap(linha):
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
                # Campo de usuário
                cmpusuario = site.verificarobjetoexiste('NAME', 'cliente_usuario')
                # Verifica se achou o campo de usuário
                if cmpusuario is not None:
                    # Limpa o campo de usuário
                    cmpusuario.clear()
                    # Carrega o campo de usuário
                    cmpusuario.send_keys(linha[Usuario])
                    # Campo de senha
                    cmpsenha = site.verificarobjetoexiste('NAME', 'cliente_senha')
                    # Verifica se achou o campo de senha
                    if cmpsenha is not None:
                        # Limpa o campo de senha
                        cmpsenha.clear()
                        # Carrega o campo de senha
                        cmpsenha.send_keys(linha[Senha])
                        # Botão de Login
                        botao = site.verificarobjetoexiste('NAME', 'autenticar')
                        # Verifica se achou o botão de ‘login’
                        if botao is not None:
                            # Clica no botão de ‘login’
                            site.navegador.execute_script("arguments[0].click()", botao)
                            time.sleep(1)
                            # Diminui o tempo de busca pela mensagem de erro
                            site.delay = 2
                            # Mensagem de erro
                            mensagemerro = site.verificarobjetoexiste('CLASS_NAME', 'bp-formulario-fracasso')
                            # Retorna o delay ao normal
                            site.delay = delay
                            # Teste se não tem mensagem de erro do site
                            if mensagemerro is None:
                                # Acho os links de gerar boleto
                                boletos = site.verificarobjetoexiste('LINK_TEXT', 'Gerar', itemunico=True)
                                # Verifica se achou os links de boleto
                                if boletos is not None:
                                    # Inicia o contador de boletos
                                    numboleto = 0
                                    # "Varre" os links de boleto
                                    for boleto in boletos:
                                        # Verifica se é um link válido
                                        if boleto.get_attribute('href') != '':
                                            # Clica no link
                                            site.navegador.execute_script("arguments[0].click()", boleto)
                                            # Define o novo nome e caminho do arquivo baixado (baseado no nr do boleto)
                                            if numboleto == 0:
                                                novonomearquivo = pastadownload + '\\' + linha[identificador] + '_' + linha[Administradora] + '.pdf'
                                            else:
                                                novonomearquivo = pastadownload + '\\' + linha[identificador] + '_' + linha[Administradora] + '_' + str(numboleto + 1) + '.pdf'

                                            # Espera o download finalizar e "pega" o arquivo baixado
                                            arquivobaixado = site.pegaarquivobaixado(tempoesperadownload)

                                            # Verifica se o arquivo baixado de fato existe
                                            if os.path.isfile(pastadownload + '\\' + arquivobaixado):
                                                aux.renomeararquivo(pastadownload + '\\' + arquivobaixado, novonomearquivo)
                                                numboleto += 1
                                                time.sleep(1)

                                # Retorna resposta na linha
                                linha[Resposta] = respostaducessopadrao(numboleto)
                            else:
                                # Retorna resposta na linha
                                if hasattr(mensagemerro, 'text'):
                                    linha[Resposta] = mensagemerro.text
                                else:
                                    linha[Resposta] = mensagemerropadrao

                            # Fecha o browser
                            site.fecharsite()
    except Exception as e:
        linha[Resposta] = str(e)

    finally:
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
        print(info.retornaradministradora('nomereal', linha[Administradora], 'site'), site.url)
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
                                site.esperadownloads(pastadownload, timeout)
                                # Define o nome (baseado no nr do boleto)
                                if numboleto == 0:
                                    novonomearquivo = pastadownload + '\\' + linha[identificador] + '_' + linha[Administradora] + '.pdf'
                                else:
                                    novonomearquivo = pastadownload + '\\' + linha[identificador] + '_' + linha[Administradora] + '_' \
                                                      + str(numboleto) + '.pdf'

                                # Retorna o último arquivo na pasta de download
                                lastfile = aux.ultimoarquivo(pastadownload, '.pdf')

                                # Retorna o último arquivo na pasta de download
                                lastfile = aux.ultimoarquivo(pastadownload, '.pdf')
                                # Verifica se achou um arquivo na pasta de download
                                if len(lastfile) > 0:
                                    aux.renomeararquivo(lastfile, novonomearquivo, linha[identificador])
                                    numboleto += 1

                    # Retorna resposta na linha
                    linha[Resposta] = respostaducessopadrao(numboleto)
                else:
                    # Retorna resposta na linha
                    linha[Resposta] = re.sub('\n+', '\n', mensagemerro.text)

                # Fecha o browser
                site.fecharsite()

    except Exception as e:
        if site is not None:
            site.fecharsite()
        linha[Resposta] = str(e)


def cipa(linha):
    # Iniciar o trabalho, só copiado de outro
    site = None
    try:
        # Variável que vai retornar a quantidade de boletos
        numboleto = 0
        boletosanalises = 0
        achouapartamento = False
        textoobjetologado = '/html/body/app-root/div[1]/div[1]/app-topbar/div[1]/div/div[2]/app-menu-rapido/div/div[1]/span'

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
                                                                            site.esperadownloads(pastadownload, timeout)

                                                                            # Define o nome (baseado no nr do boleto)
                                                                            if numboleto == 0:
                                                                                novonomearquivo = pastadownload + '\\' + linha[identificador] + '_' + \
                                                                                                  info.retornaradministradora('nomereal', linha[Administradora], 'nomereduzido') + '.pdf'
                                                                            else:
                                                                                novonomearquivo = pastadownload + '\\' + linha[identificador] + '_' + \
                                                                                                  info.retornaradministradora('nomereal', linha[Administradora], 'nomereduzido') + '_' \
                                                                                                  + str(numboleto) + '.pdf'

                                                                            # Retorna o último arquivo na pasta de download
                                                                            lastfile = aux.ultimoarquivo(pastadownload, '.pdf')
                                                                            # Verifica se achou um arquivo na pasta de download
                                                                            if len(lastfile) > 0:
                                                                                # Nâo chamo adicionando o cabeçalho porque o arquivo vem protegido
                                                                                aux.renomeararquivo(lastfile, novonomearquivo)
                                                                                numboleto += 1

                                                                    else:
                                                                        boletosanalises += 1
                                                    site.navegador.close()
                                            else:
                                                # Não existem boletos em aberto para buscar na lista do apartamento dado como entrada
                                                linha[Resposta] = respostaducessopadrao(-1)
                                    else:
                                        linha[Resposta] = "Condomínio não encontrado!"

                                if len(linha[Resposta]) == 0:
                                    if not achouapartamento:
                                        linha[Resposta] = 'Apartamento não encontrado!'
                                    else:
                                        linha[Resposta] = respostaducessopadrao(numboleto, boletosanalises)
                            else:
                                # Mensagem de erro de Login
                                linha[Resposta] = msgerro.text

    except Exception as e:
        if site is not None:
            site.fecharsite()
        linha[Resposta] = str(e)


def respostaducessopadrao(quant, quantanalise=0):
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
