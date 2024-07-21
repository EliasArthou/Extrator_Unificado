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
from bs4 import BeautifulSoup
from urllib.parse import urlparse, parse_qs


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
            anoextracao = str(dataatual.year)

        site = web.TratarSite(senha.siteiptu, senha.nomeprofileIPTU)

        codigocliente = linha[Codigo]
        caminhodestino = objeto.pastadownload + '/' + str(codigocliente) + '_' + str(linha[NrIPTU]) + '.pdf'

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
            anoextracao = str(dataatual.year)
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
    textoerro = ''

    # try:
    if dataatual == '':
        dataatual = aux.hora('America/Sao_Paulo', 'DATA')

    if dataatual:
        anoextracao = str(dataatual.year)

    codigocliente = linha[Codigo]
    caminhodestino = objeto.pastadownload + '/certidao_' + str(codigocliente) + '_' + str(linha[NrIPTU]) + '.pdf'

    objeto.visual.mudartexto('labelstatus', 'Extraindo boleto...')

    if not os.path.isfile(caminhodestino):
        while not problemacarregamento and not resolveu:
            textoerro = ''
            if site is not None:
                if site.navegador.current_url != senha.sitecertidaoefipeutica:
                    site.fecharsite()
                    site = None
            if site is None:
                site = web.TratarSite(senha.sitecertidaoefipeutica, senha.nomeprofileIPTU)
                site.abrirnavegador()
                if site.navegador is not None:
                    # Use o método find() para encontrar o primeiro objeto com a formatação desejada
                    pagina = BeautifulSoup(site.navegador.page_source, 'html.parser')
                    objeto_encontrado = pagina.find(lambda tag: tag.name == 'font' and tag.get('size') == '2' and tag.get('color') == 'red')
                    if objeto_encontrado:
                        # Verifica se o objeto contém texto
                        if objeto_encontrado.text.strip():
                            textoerro = objeto_encontrado.text.strip()

            if site is not None and site.navegador != -1:
                if site.url == senha.sitecertidaoefipeutica and site.navegador.title != 'Request Rejected':
                    # Campo de Inscrição da tela Inicial
                    inscricao = site.verificarobjetoexiste('NAME', 'inscricao')
                    if str(linha[NrIPTU]).strip() != '':
                        if inscricao is not None:
                            inscricao.clear()
                            inscricao.send_keys(linha[NrIPTU])
                            salvou, baixou = site.baixarimagem('ID', 'img', objeto.pastadownload + '/captcha.png')
                            if salvou and baixou:
                                while not resolveu:
                                    resolveu, textoerro = site.resolvercaptcha('ID', 'texto_imagem', 'NAME', 'btConsultar')
                                    if len(textoerro) > 0:
                                        print(textoerro.upper())
                                    if 'INSCRIÇÃO IMOBILIÁRIA INVÁLIDA' not in textoerro.upper():
                                        if site.navegador.title != 'Request Rejected':
                                            contartentativas += 1
                                            delay = site.delay
                                            site.delay = 2
                                            if site.verificarobjetoexiste('ID', 'texto_imagem') or site.verificarobjetoexiste('NAME', 'btConsultar') or contartentativas == limitetentativas:
                                                resolveu = False
                                                break
                                            site.delay = delay
                                        else:
                                            problemacarregamento = True
                                            break

                                        if resolveu:
                                            site.delay = 2
                                            mensagemerro = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_MSG')
                                            site.delay = 10
                                            if mensagemerro is None:
                                                botaoimpressao = site.verificarobjetoexiste('ID', 'btimprimir')
                                                if botaoimpressao is not None:
                                                    arquivos_originais = site.obter_arquivos_atuais(site.caminhodownload)
                                                    botaoimpressao.click()
                                                    arquivobaixado = site.esperar_novo_download(site.caminhodownload, arquivos_originais)

                                                    # Verifica se o arquivo baixado de fato existe
                                                    if os.path.isfile(site.caminhodownload + '\\' + arquivobaixado):
                                                        aux.adicionarcabecalhopdf(site.caminhodownload + '\\' + arquivobaixado, caminhodestino, aux.left(codigocliente, 4), posicao_y = 10)

                                                else:
                                                    dadosiptu = [codigocliente, linha[NrIPTU], anoextracao, 'Verificar (Extrair Manualmente)']

                                            else:
                                                dadosiptu = [codigocliente, linha[NrIPTU], anoextracao, mensagemerro.text]

                                        else:
                                            dadosiptu = [codigocliente, linha[NrIPTU], anoextracao, 'Problema Captcha']
                                    else:
                                        dadosiptu = [codigocliente, linha[NrIPTU], anoextracao, 'Inscrição inválida!']
                                        resolveu = True
                    else:
                        dadosiptu = [codigocliente, linha[NrIPTU], anoextracao, 'Inscrição inválida!']
                        resolveu = True

    # if os.path.isfile(caminhodestino) and salvardadospdf:
    if os.path.isfile(caminhodestino):
        listacodigo = []
        # listatipopag = []
        listaarquivo = []

        listacodigo.append("'" + codigocliente + "'")
        listaarquivo.append("'" + caminhodestino + "'")

        df = {'Codigo': listacodigo, 'Arquivo': listaarquivo}
        df = pd.DataFrame(df)

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
    gerarboleto = not bool(objeto.visual.somentevalores.get())

    if dataatual == '':
        dataatual = aux.hora('America/Sao_Paulo', 'DATA')

    if dataatual:
        anoextracao = str(dataatual.year)

    codigocliente = linha[Codigo]
    caminhodestino = objeto.pastadownload + '/' + codigocliente + '_' + linha[NrCBM] + '.pdf'

    site = web.TratarSite(senha.siteiptu, senha.nomeprofileIPTU)

    objeto.visual.mudartexto('labelstatus', 'Extraindo boleto...')

    if not os.path.isfile(caminhodestino):
        while not resolveucaptcha: #and aux.hora('America/Sao_Paulo', 'HORA') < dt.time(23, 59, 00):
            # Variável que vai receber o texto de erro do site (caso exista)
            texto = ''
            # Verifica a hora para entrar no site, caso esteja fora do horário válido, nem inicia
            # if aux.hora('America/Sao_Paulo', 'HORA') < dt.time(22, 00, 00):
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
                # Carrega a página num dataframe
                # site.retornarpaginaemdf()
                # Campo de Inscrição da tela Inicial
                inscricao = site.verificarobjetoexiste('ID', 'cbmerj', buscar_em_iframes=True)
                # inscricao = site.buscarobjetoemdf({'id': 'cbmerj'})

                # Campo de dígito verificador
                dv = site.verificarobjetoexiste('ID', 'cbmerj_dv', buscar_em_iframes=True)
                # dv = site.buscarobjetoemdf({'id': 'cbmerj_dv'})

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
                    # Pega na URL o chave do CAPTCHA para a resolução
                    elementocaptcha = site.buscarobjetoemdf({'id': 'recaptcha-token', 'nodeName': 'INPUT', 'value': '^NaN'}, 'baseURI')

                    # Parsear a URL para extrair a parte da query
                    parsed_url = urlparse(elementocaptcha)
                    query_params = parse_qs(parsed_url.query)
                    # Extrair o valor da chave 'k'
                    k_value = query_params['k'][0]
                    # Chama a função da solução de CAPTCHA
                    resposta = site.resolvecaptchatipo2(k_value)
                    # Se o CAPTCHA retornar valor ele manda a resposta para o text oculto
                    if resposta is not None:
                        # Como o text está oculto envio através da execução de script em Javascript
                        site.navegador.execute_script(f"document.getElementById('g-recaptcha-response').innerHTML = '{resposta}'")
                        # "Pega" o botão de enviar
                        botaoenvio = site.verificarobjetoexiste('ID', 'btnEnviar', buscar_em_iframes=True)
                        if botaoenvio is not None:
                            # Clica no botão de enviar
                            botaoenvio.click()


    else:
        # Mensagem de horário inválido para gerar boleto
        objeto.visual.acertaconfjanela(False)
        # Texto quando o serviço está indisponível
        # if texto != 'Este serviço encontra-se temporariamente indisponível.':
        #     # Caso o erro não seja de serviço indisponível o horário é inválido
        #     msg.msgbox('Impossível gerar boletos depois das 22:00!', msg.MB_OK, 'Horário Inválido')
        # else:
        #     # Mensagem de serviço indisponível
        #     msg.msgbox('Serviço fora do ar!', msg.MB_OK, 'Serviço com problemas')

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


    # finally:
    #     if site is not None:
    #         site.fecharsite()
