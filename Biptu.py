"""
Extração de IPTUs
"""

import os
import web
import auxiliares as aux
import sensiveis as senha
import sys
import pandas as pd
from selenium.webdriver.support.ui import Select
from datetime import date
from datetime import datetime
import datetime as dt
import messagebox as msg

Codigo = 0
NrIPTU = 1
NrCBM = 1
IndiceIPTU = 2

tempoesperadownload = 180


def extrairboletos(objeto, linha, nrguia=0, site=None):
    """
    : param linha: a linha de dados a ser analisada.
    : param objeto: janela a ser manipulada.
    """
    dadosiptu = []
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
            if nrguia == 0:
                caminhodestino = os.path.join(objeto.pastadownload, codigocliente + '_' + linha[NrIPTU] + '.pdf')
            else:
                caminhodestino = os.path.join(objeto.pastadownload, codigocliente + '_' + linha[NrIPTU] + '_' + str(nrguia + 1) + '.pdf')
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
                            mensagemerro = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_MSG')
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
                                                        aux.adicionarcabecalhopdf(baixado, caminhodestino, codigocliente, codigobarras=False)
                                if nrguia == 0 and len(guias):
                                    for indice, guia in enumerate(guias):
                                        if site is not None:
                                            site.fecharsite()
                                            site = None
                                        dadosiptutemp, dftemp = extrairboletos(objeto, linha, indice, site)
                                        if dadosiptutemp is not None:
                                            if len(dadosiptutemp) > 0:
                                                if dadosiptu is None:
                                                    # self.listadados.extend(sublist for sublist in dadosiptu)
                                                    dadosiptu = []
                                                    dadosiptu.extend(sublist for sublist in dadosiptutemp)
                                                    # dadosiptu = dadosiptutemp
                                                else:
                                                    dadosiptu.extend(sublist for sublist in dadosiptutemp)
                                                    # dadosiptu.append(dadosiptutemp)

                                        if dftemp is not None:
                                            if len(dftemp) > 0:
                                                if df is None:
                                                    df = dftemp
                                                else:
                                                    df = pd.concat([df, dftemp], ignore_index=True)
                                                    # df = df.append(dftemp, ignore_index=True)

                    else:
                        guiaexercicio = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_TELA_Guia1')
                        if guiaexercicio is None:
                            guiaexercicio = ''
                        else:
                            guiaexercicio = guiaexercicio.accessible_name
                            guiaexercicio = guiaexercicio.split("/")[-1]

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

                        if valoresimpostotela is not None:
                            for valortela in valoresimpostotela:
                                if valortela is not None:
                                    if len(valortela.text) > 0:
                                        valor = valortela.text
                                        valor = valor.replace('.', '')
                                        valor = valor.replace(',', '.')
                                        if float(valor) == 0:
                                            dadosiptu.append([codigocliente, linha[NrIPTU], guiaexercicio, '0', '0,00', contribuinte, endereco, 'Sem Guia (Provável Isento)'])
                                        else:
                                            dadosiptu.append([codigocliente, linha[NrIPTU], guiaexercicio, '0', valor, contribuinte, endereco, 'Verificar (Extrair Manualmente)'])

        if df is None or len(dadosiptu) == 0:
            if os.path.isfile(caminhodestino) and salvardadospdf:
                listacodigo = []
                listatipopag = []
                listaarquivo = []
                df = aux.extrairtextopdf(caminhodestino, 'Boletos')
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


def extrairnadaconsta(objeto, linha, dataatual=''):
    """
    : param linha: a linha de dados a ser analisada.
    : param objeto: janela a ser manipulada.
    """

    site = None
    dadosiptu = []
    df = None

    try:

        # Define se precisa gerar o boleto
        gerarboleto = not bool(objeto.visual.somentevalores.get())
        if dataatual == '':
            dataatual = aux.hora('America/Sao_Paulo', 'DATA')

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
                    inscricao.send_keys(linha[NrIPTU])
                    anosite = site.verificarobjetoexiste('ID', 'Exercicio')
                    if anosite is not None:
                        combobox = Select(anosite)
                        anoextracao = combobox.first_selected_option.text
                        # selected_value = combobox.first_selected_option.get_attribute("value")
                    else:
                        anoextracao = str(dataatual.year)
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
                                linkdownload = site.verificarobjetoexiste('LINK_TEXT', 'aqui')
                                if linkdownload is not None:
                                    if botaogerar is not None:
                                        if getattr(sys, 'frozen', False):
                                            linkdownload.click()
                                        else:
                                            site.navegador.execute_script("arguments[0].click()", linkdownload)

                                    dadosiptu = [codigocliente, linha[NrIPTU], anoextracao]

                                    baixado = site.pegaarquivobaixado(tempoesperadownload, 1)
                                    if baixado:
                                        baixado = os.path.join(site.caminhodownload, baixado)
                                    else:
                                        baixado = ''

                                    if len(baixado) > 0:
                                        caminhodestino = aux.to_raw(caminhodestino)
                                        aux.adicionarcabecalhopdf(baixado, caminhodestino, codigocliente, codigobarras=False)

                                else:
                                    dadosiptu = [codigocliente, linha[NrIPTU], anoextracao, 'Verificar (Extrair Manualmente)']

                            else:
                                dadosiptu = [codigocliente, linha[NrIPTU], anoextracao, mensagemerro.text]

        if os.path.isfile(caminhodestino):
            df = aux.extrairtextopdf(caminhodestino, 'NADA CONSTA')
            if df is not None:
                df.insert(loc=0, column='Codigo', value="'" + codigocliente + "'")
                df.insert(loc=5, column='Arquivo', value="'" + caminhodestino + "'")
                if len(dadosiptu) == len(linha) + 1:
                    data_atual = date(dataatual.year, dataatual.month, 1)
                    filtros = df[df['Vencimentos'].apply(lambda x: datetime.strptime(str(x).replace("'", ""), "%d/%m/%Y").date() < data_atual)]
                    if filtros is not None:
                        if len(filtros) > 0:
                            dadosiptu.append('Imóvel com dívida!')
                        else:
                            dadosiptu.append('Imóvel com parcelas em dia!')
                    else:
                        dadosiptu.append('')
            else:
                dadosiptu = [codigocliente, linha[NrIPTU], anoextracao, 'Sem valores a pagar!']

        if site:
            site.fecharsite()

        if df is not None:
            if len(linha) == len(dadosiptu):
                dadosiptu.append('')

        return dadosiptu, df

    finally:
        if site is not None:
            site.fecharsite()


def extraircertidaonegativa(objeto, linha, dataatual=''):
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

    # try:
    # gerarboleto = not objeto.var1.get()
    # salvardadospdf = objeto.var2.get()
    if dataatual == '':
        dataatual = aux.hora('America/Sao_Paulo', 'DATA')

    if dataatual:
        anoextracao = str(dataatual.year)

    codigocliente = linha[Codigo]
    caminhodestino = objeto.pastadownload + '/certidao_' + codigocliente + '_' + linha[NrIPTU] + '.pdf'

    site = web.TratarSite(senha.siteiptu, senha.nomeprofileIPTU)

    objeto.visual.mudartexto('labelstatus', 'Extraindo boleto...')

    # if not os.path.isfile(caminhodestino) or not gerarboleto:
    if not os.path.isfile(caminhodestino):
        while not problemacarregamento:
            if site.url != senha.sitecertidaoefipeutica or site is None:
                if site is not None:
                    site.fecharsite()
                site = web.TratarSite(senha.sitecertidaoefipeutica, senha.nomeprofileIPTU)
                site.abrirnavegador()
                # time.sleep(2)

            if site is not None and site.navegador != -1:
                if site.url == senha.sitecertidaoefipeutica and site.navegador.title != 'Request Rejected':
                    # Campo de Inscrição da tela Inicial
                    inscricao = site.verificarobjetoexiste('NAME', 'inscricao')
                    if inscricao is not None:
                        inscricao.clear()
                        inscricao.send_keys(linha[NrIPTU])
                        # Campo de Ano de extração
                        exercicio = site.verificarobjetoexiste('NAME', 'exercicio', anoextracao)
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
                                                dadosiptu = [codigocliente, linha[NrIPTU], anoextracao]
                                                baixado = os.path.join(objeto.pastadownload, site.pegaarquivobaixado(tempoesperadownload, 1))

                                                if len(baixado) > 0:
                                                    caminhodestino = aux.to_raw(caminhodestino)
                                                    aux.adicionarcabecalhopdf(baixado, caminhodestino, codigocliente, codigobarras=False)

                                            else:
                                                dadosiptu = [codigocliente, linha[NrIPTU], anoextracao, 'Verificar (Extrair Manualmente)']

                                        else:
                                            dadosiptu = [codigocliente, linha[NrIPTU], anoextracao, mensagemerro.text]

                                    else:
                                        dadosiptu = [codigocliente, linha[NrIPTU], anoextracao, 'Problema Captcha']

    # if os.path.isfile(caminhodestino) and salvardadospdf:
    if os.path.isfile(caminhodestino):
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

    # finally:
    #     if site is not None:
    #         site.fecharsite()


def extrairbombeiros(objeto, linha, dataatual=''):
    """
    : param linha: a linha de dados a ser analisada.
    : param objeto: janela a ser manipulada.
    """
    site = None
    dadosbombeiros = []
    df = None

    contartentativas = 0
    limitetentativas = 30
    problemacarregamento = False
    resolveucaptcha = False

    # try:
    gerarboleto = not objeto.somentevalores.get()

    if dataatual == '':
        dataatual = aux.hora('America/Sao_Paulo', 'DATA')

    if dataatual:
        anoextracao = str(dataatual.year)

    codigocliente = linha[Codigo]
    caminhodestino = objeto.pastadownload + '/' + codigocliente + '_' + linha[NrCBM] + '.pdf'

    site = web.TratarSite(senha.siteiptu, senha.nomeprofileIPTU)

    objeto.visual.mudartexto('labelstatus', 'Extraindo boleto...')

    if not os.path.isfile(caminhodestino):
        while not resolveucaptcha and aux.hora('America/Sao_Paulo', 'HORA') < dt.time(22, 00, 00):
            # Variável que vai receber o texto de erro do site (caso exista)
            texto = ''
            # Verifica a hora para entrar no site, caso esteja fora do horário válido, nem inicia
            if aux.hora('America/Sao_Paulo', 'HORA') < dt.time(22, 00, 00):
                # Verifica se o chrome está aberta
                if site is not None:
                    # Fecha o site
                    site.fecharsite()
                # Carrega o site na memória
                site = web.TratarSite(senha.siteCBM, senha.nomeprofileCBM)
                # Inicia o browser carregado na memória
                site.abrirnavegador()
                # Verifica se não carregou ou abriu o site errado
                if site.url != senha.siteCBM or site is None:
                    # Verifica se o chrome está aberta
                    if site is not None:
                        # Fecha o site
                        site.fecharsite()
                    # Carrega o site na memória
                    site = web.TratarSite(senha.siteCBM, senha.nomeprofileCBM)
                    # Inicia o browser carregado na memória
                    site.abrirnavegador()

                if site is not None and site.navegador != -1:
                    # Campo de Inscrição da tela Inicial
                    inscricao = site.verificarobjetoexiste('NAME', 'num_cbmerj')
                    # Campo de dígito verificador
                    dv = site.verificarobjetoexiste('NAME', 'dv_cbmerj')
                    # Testa se tem os dois campos supracitados
                    if inscricao is not None and dv is not None:
                        # "Limpa" o campo de inscrição
                        inscricao.clear()
                        # Preenche o campo de inscrição com os dados do banco de dados (sem o dígito verificador)
                        inscricao.send_keys(aux.left(linha[NrCBM], 7))
                        # "Limpa" o campo do dígito verificador
                        dv.clear()
                        # Preenche o campo do dígito verificador com os dados do banco de dados
                        dv.send_keys(aux.right(linha[NrCBM], 1))
                        # Baixa a imagem do captcha para ser solucionado
                        achouimagem, baixouimagem = site.baixarimagem('ID', 'cod_seguranca', aux.caminhoprojeto() + '\\' + 'Downloads\\captcha.png')
                        # Verifica se achou e salvou a imagem do captcha
                        if achouimagem and baixouimagem:
                            # Função que pede como entrada a caixa de texto do código de segurança e o botão de confirmar, como retorno
                            # dá um booleano que diz se conseguiu resolver o captcha e a mensagem de erro (qualquer uma), caso exista.
                            resolveucaptcha, mensagemerro = site.resolvercaptcha('NAME', 'txt_Seguranca', 'XPATH',
                                                                                 '/html/body/table[2]/tbody/tr/td[4]/table[2]/tbody/tr[2]/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td/form/input[5]')
                            # Testa se teve erro de imóvel não encontrado e começa a busca pelo IPTU
                            if 'Imóvel não encontrado. Confira os dados que você informou.' in mensagemerro or \
                                    'PREZADO(A) CONTRIBUINTE\nPara verificação dos débitos da Taxa de Incêndios dos imóveis dos municípios de Macaé, São Gonçalo e Campos dos Goytacazes, realize a consulta através do número da inscrição municipal vigente.' in mensagemerro:
                                # Limpa a mensagem de erro
                                mensagemerro = ''
                                # Se a página ficou aberta ele fecha
                                if site is not None:
                                    # Fecha o site
                                    site.fecharsite()
                                # Reabre o navegador na página de busca por código IPTU
                                site = web.TratarSite(senha.siteporCBM, senha.nomeprofileCBM)
                                # Inicia o navegador
                                site.abrirnavegador()
                                if site is not None and site.navegador != -1:
                                    # Campo de Inscrição da tela Inicial
                                    campoiptu = site.verificarobjetoexiste('NAME', 'inscricao_e')
                                    # Combobox com a lista de municípios
                                    campomunicipio = site.verificarobjetoexiste('NAME', 'cod_mun_e', 'Rio de Janeiro')
                                    # Botão para abrir a tabela de consulta
                                    btnconsultar = site.verificarobjetoexiste('XPATH',
                                                                              '/html/body/table[2]/tbody/tr/td[4]/table[2]/tbody/tr/td/table/tbody/tr/td/form/input[3]')
                                    # Verifica se tem todos os elementos na tela para chamar a resolução do captcha
                                    if campoiptu is not None and campomunicipio is not None and btnconsultar is not None:
                                        campoiptu.send_keys(linha[IndiceIPTU])
                                        if btnconsultar is not None:
                                            # Clique no botão
                                            if getattr(sys, 'frozen', False):
                                                btnconsultar.click()
                                            else:
                                                site.navegador.execute_script("arguments[0].click()", btnconsultar)

                                            achouimagem, baixouimagem = site.baixarimagem('ID', 'cod_seguranca', aux.caminhoprojeto() + '\\' + 'Downloads\\captcha.png')
                                            if achouimagem and baixouimagem:
                                                resolveucaptcha, mensagemerro = site.resolvercaptcha('NAME',
                                                                                                     'txt_Seguranca', 'XPATH',
                                                                                                     '/html/body/table[2]/tbody/tr/td[4]/table[2]/tbody/tr/td/table/tbody/tr/td/form/input[5]')

                            if resolveucaptcha and len(mensagemerro) == 0:
                                # Verifica se está na página de geração de boletos
                                if site.url == senha.siteCBM:
                                    # Área de retorno dos campos de informação do imposto
                                    # ==================================================================================================================================
                                    # Tabela de dados exclusivas do imóvel
                                    tabelaimovel = site.retornartabela(1)
                                    # Tabela de dados das cobranças (inclusive se tiver mais de um ano em aberto)
                                    tabelacobranca = site.retornartabela(2)
                                    # ==================================================================================================================================
                                    # Lista dos anos em aberto (ou parcelas do ano no caso da tx de incêndio do ano corrente)
                                    listaanos = []
                                    # Laço para percorrer as tabelas para "juntar" os dados do imóvel com os dados da cobrança
                                    for linhacobranca in tabelacobranca:
                                        # Dicionário que vai virar a linha que será adicionada no Excel
                                        linhanova = {'Cod Cliente': codigocliente}
                                        # Pega a primeira lista da íntegra
                                        linhanova.update(tabelaimovel[0])
                                        # Pega a linha de cobrança conforme o item do laço
                                        linhanova.update(linhacobranca)
                                        # Adiciona o Status do boleto conforme a data atual
                                        if int(linhacobranca.get('taxa[Exercicio]')) < anoextracao - 1:
                                            linhanova.update({'Status': 'Vencido'})
                                        else:
                                            linhanova.update({'Status': 'Imposto Ano Corrente'})

                                        # Insere a linha formada da junção das listas no Excel
                                        dadosbombeiros.append(linhanova)
                                        # Popula a lista de anos que fazem parte da extração
                                        listaanos.append(linhanova['taxa[Exercicio]'])

                                    # Começa a gerar os boletos (se marcado essa opção)
                                    if gerarboleto:
                                        # Cria uma lista com todos os objetos da classe com o nome "botao"
                                        botoes = site.verificarobjetoexiste('CLASS_NAME', 'botao', itemunico=False)
                                        # Verifica se tem o botão de gerar boleto (caso tenha só um é o botão de dívida ativa)
                                        if len(botoes) > 1:
                                            indicebotao = 0
                                            indiceano = 0
                                            anoanterior = ''
                                            # Laço para percorrer todos os botões (inclusive os que não são de boleto)
                                            for botao in botoes:
                                                # Verifica se é um botão de gerar boleto para poder interagir com ele
                                                if botao.get_attribute("value") == "Gerar Boleto":
                                                    # Ação de clicar no botão
                                                    if getattr(sys, 'frozen', False):
                                                        botao.click()
                                                    else:
                                                        site.navegador.execute_script("arguments[0].click()", botao)

                                                    # Anda com o índice cada vez que o botão é um botão que gera boleto
                                                    indicebotao += 1

                                                    # Verifica se tem que criar um índice nos arquivos por ano
                                                    # (quanto tem mais de um boleto por ano)
                                                    if listaanos[indicebotao - 1] != anoanterior:
                                                        indiceano = 0
                                                        anoanterior = listaanos[indicebotao - 1]
                                                    else:
                                                        indiceano += 1

                                                    # Trata o alerta (caso apareça), fechando o mesmo
                                                    site.trataralerta()

                                                    # Verifica a quantidade de abas
                                                    if site.num_abas() > 1:
                                                        # Vai para a última aba
                                                        site.irparaaba(site.num_abas())
                                                        # Verifica se está na tela do boleto
                                                        if site.navegador.current_url == senha.telaboletoCBM:
                                                            # Botão de impressão
                                                            imprimir = site.verificarobjetoexiste('CLASS_NAME', 'no_print')
                                                            # Verifica se o botão de impressão existe
                                                            if imprimir is not None:
                                                                # Muda o status na tela
                                                                objeto.visual.mudartexto('labelstatus', 'Salvando Boleto...')
                                                                # Ativa o botão de imprimir do visualizar impressão
                                                                site.navegador.execute_script('window.print();')
                                                                baixado = os.path.join(objeto.pastadownload, site.pegaarquivobaixado(tempoesperadownload, 1))

                                                                if len(baixado) > 0:
                                                                    if indiceano == 0:
                                                                        caminhodestino = objeto.pastadownload + '/' + listaanos[indicebotao - 1] + '_' + codigocliente + '_' + \
                                                                                         linha[NrCBM] + '.pdf'
                                                                    else:
                                                                        caminhodestino = objeto.pastadownload + '/' + listaanos[indicebotao - 1] + '_' + codigocliente + '_' + \
                                                                                         linha[NrCBM] + '_' + str(indiceano) + '.pdf'
                                                                    caminhodestino = aux.to_raw(caminhodestino)
                                                                    aux.adicionarcabecalhopdf(baixado, caminhodestino, codigocliente, codigobarras=False)


                                                        # Verifica se tem abas abertas de boletos
                                                        while site.num_abas() > 1:
                                                            # Vai para última aba para não fechar a aba que tem os botões de gerar boleto
                                                            site.irparaaba(site.num_abas())
                                                            # Fecha as abas de boletos
                                                            site.fecharaba()
                                                            # Vai para última aba para não fechar a aba que tem os botões de gerar boleto
                                                            site.irparaaba(site.num_abas())
                                                    # Se a opção de boleto único for marcada ele não continua extraindo os boletos do parcelamento
                                                    if objeto.visual.resposta == 1 and indiceano == 1 and listaanos[len(listaanos) - 1] == listaanos[indicebotao - 1]:
                                                        break

                                        # Verifica se o site está na memória
                                        if site is not None:
                                            # Fecha o site
                                            site.fecharsite()
                            # Condição se o captcha não for resolvido ou tiver uma mensagem de erro.
                            else:
                                # Verifica se o problema não foi o captcha
                                if resolveucaptcha:
                                    # Linha de erro para carregar no Excel
                                    dadosbombeiros = [codigocliente, aux.left(str(linha[NrCBM]), 7) + '-' + aux.right(str(linha[NrCBM]), 1), '', '', '', '', '', '', '', '', '', '', '', '', mensagemerro]
                                # Fecha o site
                                site.fecharsite()
                    else:
                        texto = site.retornartabela(3)
                        if texto == 'Este serviço encontra-se temporariamente indisponível.':
                            break
    else:
        # Mensagem de horário inválido para gerar boleto
        objeto.visual.acertaconfjanela(False)
        # Texto quando o serviço está indisponível
        if texto != 'Este serviço encontra-se temporariamente indisponível.':
            # Caso o erro não seja de serviço indisponível o horário é inválido
            msg.msgbox('Impossível gerar boletos depois das 22:00!', msg.MB_OK, 'Horário Inválido')
        else:
            # Mensagem de serviço indisponível
            msg.msgbox('Serviço fora do ar!', msg.MB_OK, 'Serviço com problemas')

        # Sai do Looping0
        # break
                # if len(baixado) > 0:
                #     caminhodestino = aux.to_raw(caminhodestino)
                #     aux.adicionarcabecalhopdf(baixado, caminhodestino, codigocliente)
                #
                # else:
                #     dadoscbm = [codigocliente, linha[NrIPTU], anoextracao, 'Verificar (Extrair Manualmente)']
            #
            # else:
            #     dadoscbm = [codigocliente, linha[NrIPTU], anoextracao, mensagemerro.text]

        # else:
        #     dadoscbm = [codigocliente, linha[NrIPTU], anoextracao, 'Problema Captcha']

    # if os.path.isfile(caminhodestino) and salvardadospdf:
    if os.path.isfile(caminhodestino):
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

    # finally:
    #     if site is not None:
    #         site.fecharsite()
