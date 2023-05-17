"""
Extração de IPTUs
"""

import os
import web
import auxiliares as aux
import sensiveis as senha
import time
import sys

Codigo = 0
NrIPTU = 1
IndiceIPTU = 2

tempoesperadownload = 180


def extrairboletos(objeto, linha, nrguia=0, site=None):
    """
    : param linha: a linha de dados a ser analisada.
    : param objeto: janela a ser manipulada.
    """
    dadosiptu = []
    # dadosintermediario = []
    df = None
    caminhodestino = ''
    guias = None
    mensagemerro = None
    mensagemguia = None

    try:
        errocarregamentosite = 'Exception: Message TelaSelecao was received; the expected message was SegundaTela.<br />' \
                               'WebMessage: Message TelaSelecao was received; the expected message was SegundaTela.'

        # Define se precisa gerar o boleto
        gerarboleto = not bool(objeto.visual.somentevalores.get())
        # Define se precisa salvar o PDF
        salvardadospdf = objeto.visual.codigosdebarra.get()
        # Define com qual data de vencimento tem que gerar
        resposta = str(objeto.visual.tipopagamento.get())

        if site is None:
            site = web.TratarSite(senha.siteiptu, senha.nomeprofileIPTU)
        else:
            site.delay = 2
            guias = site.verificarobjetoexiste('CSS_SELECTOR', '[title="Clique aqui para visualizar / imprimir esta guia."]', itemunico=False)
            site.delay = 10
            if guias is not None:
                mensagemerro = None
                mensagemguia = None

        codigocliente = linha[Codigo]
        # Informa que a parte da extração está sendo feita
        objeto.visual.mudartexto('labelstatus', 'Extraindo boleto...')
        if gerarboleto:
            caminhodestino = os.path.join(objeto.pastadownload, codigocliente + '_' + linha[NrIPTU] + '.pdf')

        # Verifica se o arquivo já existe e se não está pedindo pra pegar as informações do site
        # Se o arquivo existir ele pega as informações do arquivo
        if not os.path.isfile(caminhodestino) or not gerarboleto:
            if guias is None:
                if site is not None:
                    site.fecharsite()
                site = web.TratarSite(senha.siteiptu, senha.nomeprofileIPTU)
                site.abrirnavegador()
                if site.url != senha.siteiptu or site is None:
                    if site is not None:
                        site.fecharsite()
                    site = web.TratarSite(senha.siteiptu, senha.nomeprofileIPTU)
                    site.abrirnavegador()

            if site is not None and site.navegador != -1:
                # Campo de Inscrição da tela Inicial
                inscricao = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_inscricao_input')
                if inscricao is not None:
                    inscricao.send_keys(linha[NrIPTU])
                    if site.url == senha.siteiptu:
                        # Botão pra entrar na área de boletos
                        botaogerar = site.verificarobjetoexiste('NAME', 'ctl00$ePortalContent$DefiniGuia')
                        if botaogerar is not None:
                            if getattr(sys, 'frozen', False):
                                botaogerar.click()
                            else:
                                site.navegador.execute_script("arguments[0].click()", botaogerar)

                            site.delay = 2
                            # Verifica se teve mensagem de erro
                            mensagemerro = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_MENSAGEM')
                            site.delay = 10
                            if mensagemerro is not None:
                                if hasattr(mensagemerro, 'text'):
                                    while mensagemerro.text == errocarregamentosite:
                                        if site is not None:
                                            site.fecharsite()
                                        site = web.TratarSite(senha.siteiptu, senha.nomeprofileIPTU)
                                        site.abrirnavegador()

                                        if site is not None and site.navegador != -1:
                                            # Campo de Inscrição da tela Inicial
                                            inscricao = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_inscricao_input')
                                            if inscricao is not None:
                                                inscricao.clear()
                                                inscricao.send_keys(linha[NrIPTU])
                                                if site.url == senha.siteiptu:
                                                    botaogerar = site.verificarobjetoexiste('NAME', 'ctl00$ePortalContent$DefiniGuia')
                                                    if botaogerar is not None:
                                                        if getattr(sys, 'frozen', False):
                                                            botaogerar.click()
                                                        else:
                                                            site.navegador.execute_script("arguments[0].click()", botaogerar)

                                                        mensagemerro = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_MENSAGEM')
                                                        if mensagemerro is None:
                                                            break
                                    if mensagemerro is not None:
                                        if hasattr(mensagemerro, 'text'):
                                            if mensagemerro.text != '':
                                                dadosiptu = [codigocliente, linha[NrIPTU], '', '', '', '', '', mensagemerro.text]

                if mensagemerro is None:
                    if guias is None:
                        site.delay = 2
                        mensagemguia = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_TELA_M1')
                        site.delay = 10

                    if mensagemguia is None:
                        if guias is None:
                            guias = site.verificarobjetoexiste('CSS_SELECTOR', '[title="Clique aqui para visualizar / imprimir esta guia."]', itemunico=False)
                        if guias is not None:
                            guia = guias[nrguia]
                            if guia is not None:
                                if getattr(sys, 'frozen', False):
                                    guia.click()
                                else:
                                    site.navegador.execute_script("arguments[0].click()", guia)

                                guiaexercicio = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_GuiaExercicio')
                                if guiaexercicio is None:
                                    guiaexercicio = ''
                                else:
                                    guiaexercicio = guiaexercicio.text

                                contribuinte = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_TELA_CONTRIBUINTE')
                                if contribuinte is None:
                                    contribuinte = ''
                                else:
                                    contribuinte = contribuinte.text

                                endereco = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_TELA_ENDERECO')
                                if endereco is None:
                                    endereco = ''
                                else:
                                    endereco = endereco.text

                                botaogerarid = 'ctl00_ePortalContent_btnDarmIndiv'

                                match resposta:
                                    case '1':
                                        idselecionado = 'ctl00$ePortalContent$cbCotaUnica'
                                        namevalor = 'ctl00_ePortalContent_TELA_VALOR_COTA_UNICA'
                                        valores = site.verificarobjetoexiste('ID', namevalor)
                                        dadosiptu = [codigocliente, linha[NrIPTU], guiaexercicio, 1, valores.text, contribuinte, endereco]

                                    case '2' | '3' | '4':
                                        idselecionado = 'ctl00_ePortalContent_Chk_00' + str(int(resposta) - 1)
                                        if resposta == '2':
                                            namevalor = 'Valor_Prim'
                                        elif resposta == '3':
                                            namevalor = 'Valor_Segu'
                                        else:
                                            namevalor = 'Valor_Terc'

                                        valores = site.verificarobjetoexiste('NAME', namevalor, itemunico=False)

                                        for index, valor in enumerate(valores):
                                            dadosiptuintermediario = [codigocliente, linha[NrIPTU], guiaexercicio, str(index + 1), valor.text,
                                                                      contribuinte, endereco, 'Ok']
                                            if dadosiptuintermediario:
                                                dadosiptu.append(dadosiptuintermediario)
                                    case _:
                                        idselecionado = ''

                                if gerarboleto:
                                    cota = site.verificarobjetoexiste('ID', idselecionado, iraoobjeto=True)
                                    if cota is not None:
                                        site.descerrolagem()
                                        botaogerar = site.verificarobjetoexiste('ID', botaogerarid, iraoobjeto=True)
                                        if botaogerar is not None:
                                            confirmar = site.verificarobjetoexiste('ID', 'popup_ok')
                                            if confirmar is not None:
                                                if getattr(sys, 'frozen', False):
                                                    confirmar.click()
                                                else:
                                                    site.navegador.execute_script("arguments[0].click()", confirmar)

                                                if site.navegador.current_url == senha.telaboletoIPTU:
                                                    linkdownload = site.verificarobjetoexiste('LINK_TEXT', 'aqui')
                                                    if linkdownload is not None:
                                                        if botaogerar is not None:
                                                            if getattr(sys, 'frozen', False):
                                                                linkdownload.click()
                                                            else:
                                                                site.navegador.execute_script("arguments[0].click()", linkdownload)

                                                    baixado = site.pegaarquivobaixado(tempoesperadownload, 1)
                                                    if baixado:
                                                        baixado = os.path.join(site.caminhodownload, baixado)
                                                    else:
                                                        baixado = ''

                                                    if len(baixado) > 0:
                                                        objeto.visual.mudartexto('labelstatus', 'Salvando Boleto...')
                                                        caminhodestino = aux.to_raw(caminhodestino)
                                                        aux.adicionarcabecalhopdf(baixado, caminhodestino, codigocliente)
                                if nrguia == 0 and len(guias):
                                    for indice, guia in enumerate(guias):
                                        dadosiptutemp, dftemp = extrairboletos(objeto, linha, indice, site)
                                        if dadosiptutemp is not None:
                                            if dadosiptu is None:
                                                dadosiptu = dadosiptutemp
                                            else:
                                                dadosiptu.append(dadosiptutemp)

                                        if dftemp is not None:
                                            if df is None:
                                                df = dftemp
                                            else:
                                                df.append(dftemp)

                    else:
                        guiaexercicio = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_TELA_Guia1')
                        if guiaexercicio is None:
                            guiaexercicio = ''
                        else:
                            guiaexercicio = guiaexercicio.text

                        contribuinte = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_TELA_CONTRIBUINTE')
                        if contribuinte is None:
                            contribuinte = ''
                        else:
                            contribuinte = contribuinte.text

                        endereco = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_TELA_ENDERECO')
                        if endereco is None:
                            endereco = ''
                        else:
                            endereco = endereco.text

                        valoresimpostotela = site.verificarobjetoexiste('CSS_SELECTOR', '[class ="ValorDevido col bordas"]', itemunico=False)
                        # valorimpostotela = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_TELA_Valor1')
                        if valoresimpostotela is not None:
                            for valortela in valoresimpostotela:

                                if valortela is not None:
                                    if len(valortela.text) > 0:
                                        valor = valortela.text
                                        valor = valor.replace('.', '')
                                        valor = valor.replace(',', '.')
                                    else:
                                        valor = 0
                                else:
                                    valor = 0

                                if float(valor) == 0:
                                    dadosiptu.append([codigocliente, linha[NrIPTU], guiaexercicio, '0', '0,00', contribuinte, endereco, 'Sem Guia (Provável Isento)'])
                                else:
                                    dadosiptu.append([codigocliente, linha[NrIPTU], guiaexercicio, '0', valor, contribuinte, endereco, 'Verificar (Extrair Manualmente)'])

        if df is None or len(dadosiptu) == 0:
            if os.path.isfile(caminhodestino) and salvardadospdf:
                listacodigo = []
                listatipopag = []
                listaarquivo = []
                df = aux.extrairtextopdf(caminhodestino)
                total_rows = df[df.columns[0]].count()
                for linhatotais in range(total_rows):
                    listacodigo.append("'" + codigocliente + "'")
                    listatipopag.append("'PARCELADO'")
                    listaarquivo.append("'" + os.path.basename(caminhodestino) + "'")

                df.insert(loc=0, column='Codigo', value=listacodigo)
                df.insert(loc=4, column='TpoPagto', value=listatipopag)
                df.insert(loc=5, column='Arquivo', value=listaarquivo)

            # objeto.bd.adicionardf('Codigos IPTUs', df, 7)
        if site is not None:
            site.fecharsite()

        if df is None:
            return dadosiptu, None
        else:
            return dadosiptu, df

    finally:
        if site is not None:
            site.fecharsite()



def extrairnadaconsta(objeto, linha):
    """

    : param linha: a linha de dados a ser analisada.
    : param objeto: janela a ser manipulada.
    """

    site = None
    dadosiptu = []
    df = None

    # try:

    # Define se precisa gerar o boleto
    gerarboleto = not objeto.var1.get()
    # Define se precisa salvar o PDF
    salvardadospdf = objeto.var2.get()

    site = web.TratarSite(senha.siteiptu, senha.nomeprofileIPTU)

    codigocliente = linha[Codigo]
    caminhodestino = objeto.pastadownload + '/' + codigocliente + '_' + linha[NrIPTU] + '.pdf'

    # Verifica se o arquivo já existe e se não está pedindo pra pegar as informações do site
    # Se o arquivo existir ele pega as informações do arquivo
    if not os.path.isfile(caminhodestino) or not gerarboleto:
        if site.url != senha.sitenadaconsta or site is None:
            if site is not None:
                site.fecharsite()
            site = web.TratarSite(senha.sitenadaconsta, senha.nomeprofileIPTU)
            site.abrirnavegador()

        if site is not None and site.navegador != -1:
            # Campo de Inscrição da tela Inicial
            inscricao = site.verificarobjetoexiste('ID', 'Inscricao')
            if inscricao is not None:
                inscricao.clear()
                inscricao.send_keys(linha['iptu'])
                if site.url == senha.sitenadaconsta:
                    botaogerar = site.verificarobjetoexiste('ID', 'Avancar')
                    if botaogerar is not None:
                        if getattr(sys, 'frozen', False):
                            botaogerar.click()
                        else:
                            site.navegador.execute_script("arguments[0].click()", botaogerar)

                        site.delay = 2
                        mensagemerro = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_MSG')
                        site.delay = 10
                        if mensagemerro is None:
                            linkdownload = site.verificarobjetoexiste('LINK_TEXT', 'link')
                            if linkdownload is not None:
                                if botaogerar is not None:
                                    if getattr(sys, 'frozen', False):
                                        linkdownload.click()
                                    else:
                                        site.navegador.execute_script("arguments[0].click()", linkdownload)

                                dadosiptu = [codigocliente, linha[NrIPTU], '2022']

                                baixado = os.path.join(objeto.pastadownload, site.pegaarquivobaixado(tempoesperadownload, 1))

                                if len(baixado) > 0:
                                    caminhodestino = aux.to_raw(caminhodestino)
                                    aux.adicionarcabecalhopdf(baixado, caminhodestino, codigocliente)
                                    # if salvardadospdf:
                                    #     listacodigo = []
                                    #     listatipopag = []
                                    #     listaarquivo = []
                                    #     df = aux.extrairtextopdf(caminhodestino)
                                    #     total_rows = df[df.columns[0]].count()
                                    #     for linhatotais in range(total_rows):
                                    #         listacodigo.append("'" + codigocliente + "'")
                                    #         listatipopag.append("'PARCELADO'")
                                    #         listaarquivo.append("'" + caminhodestino + "'")
                                    #
                                    #     df.insert(loc=0, column='Codigo', value=listacodigo)
                                    #     df.insert(loc=4, column='TpoPagto', value=listatipopag)
                                    #     df.insert(loc=5, column='Arquivo', value=listaarquivo)

                            else:
                                dadosiptu = [codigocliente, linha[NrIPTU], '2022', 'Verificar (Extrair Manualmente)']

                        else:
                            dadosiptu = [codigocliente, linha['iptu'], '2022', mensagemerro.text]

    if os.path.isfile(caminhodestino) and salvardadospdf:
        listacodigo = []
        listatipopag = []
        listaarquivo = []
        df = aux.extrairtextopdf(caminhodestino)
        total_rows = df[df.columns[0]].count()
        for linhatotais in range(total_rows):
            listacodigo.append("'" + codigocliente + "'")
            listatipopag.append("'PARCELADO'")
            listaarquivo.append("'" + caminhodestino + "'")

        df.insert(loc=0, column='Codigo', value=listacodigo)
        df.insert(loc=4, column='TpoPagto', value=listatipopag)
        df.insert(loc=5, column='Arquivo', value=listaarquivo)

    if site:
        site.fecharsite()

    #     if df is None:
    #         return dadosiptu
    #     else:
    #         return dadosiptu, df
    #
    # finally:
    #     if site is not None:
    #         site.fecharsite()


def extraircertidaonegativa(objeto, linha):
    """
    : param linha: a linha de dados a ser analisada.
    : param objeto: janela a ser manipulada.
    """
    site = None
    dadosiptu = []
    df = None

    contartentativas = 0
    limitetentativas = 30
    problemacarregamento = False
    resolveu = False

    try:
        gerarboleto = not objeto.var1.get()
        salvardadospdf = objeto.var2.get()

        codigocliente = linha[Codigo]
        caminhodestino = objeto.pastadownload + '/certidao' + codigocliente + '_' + linha[NrIPTU] + '.pdf'

        # listachaves = ['Código Cliente', 'Inscrição', 'Guia do Exercício', 'Status']
        # listaexcel = []
        site = web.TratarSite(senha.siteiptu, senha.nomeprofileIPTU)

        objeto.mudartexto('labelstatus', 'Extraindo boleto...')

        if not os.path.isfile(caminhodestino) or not gerarboleto:
            while not problemacarregamento:
                if site.url != senha.sitecertidaonegativa or site is None:
                    if site is not None:
                        site.fecharsite()
                    site = web.TratarSite(senha.sitecertidaonegativa, senha.nomeprofileIPTU)
                    site.abrirnavegador()
                    time.sleep(2)

                if site is not None and site.navegador != -1:
                    if site.url == senha.sitecertidaonegativa and site.navegador.title != 'Request Rejected':
                        # Campo de Inscrição da tela Inicial
                        inscricao = site.verificarobjetoexiste('NAME', 'inscricao')
                        if inscricao is not None:
                            inscricao.clear()
                            inscricao.send_keys(linha['iptu'])
                            # Campo de Ano de extração
                            exercicio = site.verificarobjetoexiste('NAME', 'exercicio', '2022')
                            if exercicio is not None:
                                salvou, baixou = site.baixarimagem('ID', 'img', objeto.pastadownload + '/captcha.png')
                                if salvou and baixou:
                                    while not resolveu:
                                        resolveu = site.resolvercaptcha('ID', 'texto_imagem', 'NAME', 'btConsultar')
                                        if site.navegador.title != 'Request Rejected':
                                            contartentativas += 1
                                            if site.verificarobjetoexiste('ID', 'texto_imagem') or site.verificarobjetoexiste('NAME', 'btConsultar') or contartentativas == limitetentativas:
                                                resolveu = False
                                                break
                                        else:
                                            problemacarregamento = True
                                            break
                                    if not problemacarregamento:
                                        if resolveu:
                                            site.delay = 2
                                            mensagemerro = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_MSG')
                                            site.delay = 10
                                            if mensagemerro is None:
                                                linkdownload = site.verificarobjetoexiste('LINK_TEXT', 'link')
                                                if linkdownload is not None:
                                                    dadosiptu = [codigocliente, linha['iptu'], '2022']
                                                    # site.esperadownloads(objeto.pastadownload, 10)
                                                    # baixado = aux.ultimoarquivo(objeto.pastadownload, '.pdf')
                                                    baixado = os.path.join(objeto.pastadownload, site.pegaarquivobaixado(tempoesperadownload, 1))

                                                    if len(baixado) > 0:
                                                        caminhodestino = aux.to_raw(caminhodestino)
                                                        aux.adicionarcabecalhopdf(baixado, caminhodestino, codigocliente)

                                                else:
                                                    dadosiptu = [codigocliente, linha['iptu'], '2022', 'Verificar (Extrair Manualmente)']

                                            else:
                                                dadosiptu = [codigocliente, linha['iptu'], '2022', mensagemerro.text]

                                        else:
                                            dadosiptu = [codigocliente, linha['iptu'], '2022', 'Problema Captcha']

        if os.path.isfile(caminhodestino) and salvardadospdf:
            listacodigo = []
            listatipopag = []
            listaarquivo = []
            df = aux.extrairtextopdf(caminhodestino)
            total_rows = df[df.columns[0]].count()
            for linhatotais in range(total_rows):
                listacodigo.append("'" + codigocliente + "'")
                listatipopag.append("'PARCELADO'")
                listaarquivo.append("'" + caminhodestino + "'")

            df.insert(loc=0, column='Codigo', value=listacodigo)
            df.insert(loc=4, column='TpoPagto', value=listatipopag)
            df.insert(loc=5, column='Arquivo', value=listaarquivo)

        if df is None:
            return dadosiptu
        else:
            return dadosiptu, df

    finally:
        if site is not None:
            site.fecharsite()

