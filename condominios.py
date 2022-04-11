
import web
import sensiveis as info
import time
import auxiliares as aux
import os


identificador = 0
Usuario = 1
Senha = 2
Administradora = 3
Condominio = 4
Apartamento = 5
Resposta = 6
CheckArquivo = 7

timeout = 3
pastadownload = aux.caminhoprojeto() + '\\' + 'Downloads'


def abrj(linha):
    site = None
    try:
        # Quebra o bloco e apartamento da informação
        linhabloco = linha[Apartamento]+'/'
        apartamentobuscado, blocobuscado = linhabloco.split('/')

        # Garante que o apartamento tenha 4 caracteres
        if len(apartamentobuscado.strip()) > 0:
            apartamentobuscado = aux.right('0000' + apartamentobuscado.strip(), 4).strip()

        # Trata o bloco
        blocobuscado = blocobuscado.strip()

        # Variável que vai retornar a quantidade de boletos
        numboleto = 0
        boletosanalises = 0
        achooapartamento = False

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
                # Campo de senha
                cmpsenha = site.verificarobjetoexiste('ID', 'senha')
                # Se o campo senha não estiver visível ele "clica" no botão entrar para exibir o campo "senha"
                if cmpsenha is None:
                    # Botão de Salvar
                    btnsalvar = site.verificarobjetoexiste('ID', 'salvar')
                    # Verifica se achou o botão Salvar
                    if btnsalvar is not None:
                        site.navegador.execute_script("arguments[0].click()", btnsalvar)

                if cmpsenha is not None:
                    if cmpsenha.is_displayed():
                        cmpsenha.send_keys(linha[Senha])
                        # Botão de Salvar
                        btnsalvar = site.verificarobjetoexiste('ID', 'salvar')
                        # Verifica se achou o botão Salvar
                        if btnsalvar is not None:
                            site.navegador.execute_script("arguments[0].click()", btnsalvar)
                            # Verifica se deu erro após apertar no botão para realizar o "LOGIN"
                            msgerro = site.verificarobjetoexiste('ID', 'divMsgErroArea')
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
                                            # Verifica se tem a mensagem de que não tem boleto disponível
                                            if not site.verificarobjetoexiste('CLASS_NAME', 'conteudo-nenhuma-cobranca', sotestar=True):
                                                if site.verificarobjetoexiste('CLASS_NAME', 'numero', sotestar=True):
                                                    numapartamentos = site.verificarobjetoexiste('CLASS_NAME', 'numero', itemunico=False)
                                                    dadosapartamentos = site.verificarobjetoexiste('CSS_SELECTOR', "[class='infos col-md-6']", itemunico=False)
                                                    complementoapartamentos = site.verificarobjetoexiste('CLASS_NAME', "complemento", itemunico=False)
                                                    for numeroap, dadosap, complementoap in zip(numapartamentos, dadosapartamentos, complementoapartamentos):
                                                        # Garante que só o exibido(no caso o "a vencer") seja testado
                                                        if numeroap.is_displayed():
                                                            if numeroap.text == apartamentobuscado and complementoap == blocobuscado:
                                                                achouapartamento = True
                                                                if "Indisponível" not in dadosap.text:
                                                                    site.navegador.execute_script("arguments[0].click()", numeroap)
                                                                    # Seleciona a opção na janela "popup" que aparece
                                                                    while not site.verificarobjetoexiste('ID', 'salvar', sotestar=True):
                                                                        time.sleep(1)
                                                                    btnsalvar = site.verificarobjetoexiste('ID', 'salvar')
                                                                    site.navegador.execute_script("arguments[0].click()", btnsalvar)
                                                                    # Vai para segunda página que ele abre com o boleto resumido
                                                                    achoupagina = site.navegador.irparaaba(titulo='segundavia')
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
                                                                                novonomearquivo = pastadownload + '\\' + linha[identificador] + '_' + linha[Administradora] + '.pdf'
                                                                            else:
                                                                                novonomearquivo = pastadownload + '\\' + linha[identificador] + '_' + linha[
                                                                                    Administradora] + '_' + str(numboleto) + '.pdf'

                                                                            # Retorna o último arquivo na pasta de download
                                                                            lastfile = aux.ultimoarquivo(pastadownload, '.pdf')
                                                                            # Verifica se achou um arquivo na pasta de download
                                                                            if len(lastfile) > 0:
                                                                                aux.renomeararquivo(lastfile, novonomearquivo)
                                                                                numboleto += 1

                                                                            site.navegador.close()

                                                                        achoupagina = site.navegador.irparaaba(titulo='Areadocondominio')
                                                                    else:
                                                                        boletosanalises += 1
                                            else:
                                                # Não existem boletos em aberto para buscar na lista o apartamento dado como entrada
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
    print(info.retornaradministradora('nomereal', linha[Administradora], 'site'))


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
                            # Mensagem de erro
                            mensagemerro = site.verificarobjetoexiste('CLASS_NAME', 'bp-formulario-fracasso')
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
                                            # Espera o download finalizar
                                            site.esperadownloads(pastadownload, timeout)

                                            # Define o nome (baseado no nr do boleto)
                                            if numboleto == 0:
                                                novonomearquivo = pastadownload + '\\' + linha[identificador] + '_' + linha[Administradora] + '.pdf'
                                            else:
                                                novonomearquivo = pastadownload + '\\' + linha[identificador] + '_' + linha[Administradora] + '_' + str(numboleto) + '.pdf'

                                            # Retorna o último arquivo na pasta de download
                                            lastfile = aux.ultimoarquivo(pastadownload, '.pdf')
                                            # Verifica se achou um arquivo na pasta de download
                                            if len(lastfile) > 0:
                                                aux.renomeararquivo(lastfile, novonomearquivo)
                                                numboleto += 1

                                # Retorna resposta na linha
                                linha[Resposta] = respostaducessopadrao(numboleto)
                            else:
                                # Retorna resposta na linha
                                linha[Resposta] = mensagemerro.text

                            # Fecha o browser
                            site.fecharsite()
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
                resposta = 'Salvo ' + str(quant) + 'boleto'
            case -1:
                resposta = 'Sem boletos em aberto para analisar'
            case _:
                resposta = 'Salvo ' + str(quant) + 'boletos'
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
