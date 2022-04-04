
import web
import sensiveis as info
import time
import auxiliares as aux
import os

timeout = 3
pastadownload = aux.caminhoprojeto() + '\\' + 'Downloads'


def abrj (linha, objetobrowser):
    site = None
    try:
        # Quebra o bloco e apartamento da informação
        apartamentobuscado, blocobuscado = linha[5].split('/')

        # Garante que o apartamento tenha 4 caracteres
        if len(apartamentobuscado.strip())>0:
            apartamentobuscado = aux.right('0000' + apartamentobuscado.strip(), 4).strip()

        # Trata o bloco
        blocobuscado = blocobuscado.strip()


        # Variável que vai retornar a quantidade de boletos
        numboleto = 0
        # Prepara o objeto
        site = objetobrowser.TratarSite(info.retornaradministradora('nomereal', linha[3], 'site'), info.nomeprofilecond)
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[3], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.siteiptu, info.nomeprofileIPTU)
            # Abre o browser
            site.abrirnavegador()
            if site is not None and site.navegador != -1:
                # Campo de usuário
                usuario = site.verificarobjetoexiste('NAME', 'cliente_usuario')
                # Verifica se achou o campo de usuário
                if usuario is not None:

    except Exception as e:
        if site is not None:
            site.fecharsite()
        linha[6] = str(e)

def apsa(linha):
    print(info.retornaradministradora('nomereal', linha[3], 'site'))


def bap(linha, objetobrowser):
    site = None
    try:
        # Variável que vai retornar a quantidade de boletos
        numboleto = 0
        # Prepara o objeto
        site = objetobrowser.TratarSite(info.retornaradministradora('nomereal', linha[3], 'site'), info.nomeprofilecond)
        # Abre o browser
        site.abrirnavegador()
        # Verifica se iniciou o site
        if site.url != info.retornaradministradora('nomereal', linha[3], 'site') or site is None:
            # Caso esteja com outra página aberta, fecha
            if site is not None:
                site.fecharsite()
            # Inicia o browser
            site = web.TratarSite(info.siteiptu, info.nomeprofileIPTU)
            # Abre o browser
            site.abrirnavegador()
            if site is not None and site.navegador != -1:
                # Campo de usuário
                usuario = site.verificarobjetoexiste('NAME', 'cliente_usuario')
                # Verifica se achou o campo de usuário
                if usuario is not None:
                    # Limpa o campo de usuário
                    usuario.clear()
                    # Carrega o campo de usuário
                    usuario.send_keys(linha[1])
                    # Campo de senha
                    senha = site.verificarobjetoexiste('NAME', 'cliente_senha')
                    # Verifica se achou o campo de senha
                    if senha is not None:
                        # Limpa o campo de senha
                        senha.clear()
                        # Carrega o campo de senha
                        senha.send_keys(linha[2])
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
                                                novonomearquivo = pastadownload + '\\' + linha[0] + '_' + linha[3] + '.pdf'
                                            else:
                                                novonomearquivo = pastadownload + '\\' + linha[0] + '_' + linha[3] + '_' + str(numboleto) + '.pdf'

                                            # Retorna o último arquivo na pasta de download
                                            lastfile = aux.ultimoarquivo(pastadownload, '.pdf')
                                            # Verifica se achou um arquivo na pasta de download
                                            if len(lastfile) > 0:
                                                aux.renomeararquivo(lastfile, novonomearquivo)
                                                numboleto += 1

                                # Retorna resposta na linha
                                linha[6] = respostaducessopadrao(numboleto)
                            else:
                                # Retorna resposta na linha
                                linha[6] = mensagemerro.text

                            # Fecha o browser
                            site.fecharsite()
    except Exception as e:
        if site is not None:
            site.fecharsite()
        linha[6] = str(e)


def respostaducessopadrao(quant, quantanalise):
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
                        resposta = resposta + " e tem " + quantanalise + " boleto para ser analisado (provável em acordo)"
                    case _:
                        resposta = resposta + " e tem " + quantanalise + " boletos para ser analisado (provável em acordo)"

            case _:
                resposta = 'Salvo ' + quant + ' boletos'
                match quantanalise:
                    case 1:
                        resposta = resposta + " e tem " + quantanalise + " boleto para ser analisado (provável em acordo)"
                    case _:
                        resposta = resposta + " e tem " + quantanalise + " boletos para ser analisado (provável em acordo)"

    return resposta
