"""
Extração de IPTUs
"""

import os
from core import web
from auxiliares import utils as aux
import sys
import pandas as pd
from selenium.webdriver.support.ui import Select
from datetime import date
from datetime import datetime
from bs4 import BeautifulSoup
from urllib.parse import urlparse, parse_qs
from dotenv import load_dotenv
import inspect

# Carrega as variáveis do arquivo .env
load_dotenv()


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
            site = web.TratarSite(os.getenv('SITEIPTU'), os.getenv('NOMEPROFILEIPTU'))
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
                site = web.TratarSite(os.getenv('SITEIPTU'), os.getenv('NOMEPROFILEIPTU'))
                site.abrirnavegador()
                if site.url != os.getenv('SITEIPTU') or site is None:
                    if site is not None:
                        site.fecharsite()
                    site = web.TratarSite(os.getenv('SITEIPTU'), os.getenv('NOMEPROFILEIPTU'))
                    site.abrirnavegador()

            if site is not None and site.navegador != -1:
                # Campo de Inscrição da tela Inicial
                inscricao = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_inscricao_input')
                if inscricao is not None:
                    inscricao.send_keys(linha[NrIPTU])
                    if site.url == os.getenv('SITEIPTU'):
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
                                        site = web.TratarSite(os.getenv('SITEIPTU'), os.getenv('NOMEPROFILEIPTU'))
                                        site.abrirnavegador()

                                        if site is not None and site.navegador != -1:
                                            # Campo de Inscrição da tela Inicial
                                            inscricao = site.verificarobjetoexiste('ID', 'ctl00_ePortalContent_inscricao_input')
                                            if inscricao is not None:
                                                inscricao.clear()
                                                inscricao.send_keys(linha[NrIPTU])
                                                if site.url == os.getenv('SITEIPTU'):
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

                                                if site.navegador.current_url == os.getenv('TELABOLETOIPTU'):
                                                    linkdownload = site.verificarobjetoexiste('LINK_TEXT', 'aqui', esperar_clicavel=True)
                                                    if linkdownload:
                                                        if botaogerar is not None:
                                                            caminho_arquivo = site.monitorar_downloads_sem_href(
                                                                    link_elemento=linkdownload,
                                                                    timeout=120,
                                                                    nome_arquivo='iptu.pdf'
                                                                    # clickscript=not(getattr(sys, 'frozen', False))
                                                                )
                                                            if caminho_arquivo:
                                                                # novonomearquivo = os.path.join(
                                                                #     objeto.pastadownload,
                                                                #     codigocliente) + ".pdf"


                                                                aux.adicionarcabecalhopdf(caminho_arquivo,
                                                                                        caminhodestino,
                                                                                        codigocliente,
                                                                                        codigobarras=False)
                                                        else:
                                                            print(f"Falha ao capturar o arquivo baixado para o boleto.\n"
                                                        f"Função: {inspect.currentframe().f_code.co_name} \n"
                                                        f"Cliente: {linha[identificador]}")

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

        site = web.TratarSite(os.getenv('SITEIPTU'), os.getenv('NOMEPROFILEIPTU'))

        codigocliente = linha[Codigo]
        caminhodestino = objeto.pastadownload + '/' + str(codigocliente) + '_' + str(linha[NrIPTU]) + '.pdf'

        # Verifica se o arquivo já existe e se não está pedindo pra pegar as informações do site
        # Se o arquivo existir ele pega as informações do arquivo
        if not os.path.isfile(caminhodestino) or not gerarboleto:
            if site.url != os.getenv('SITENADACONSTA') or site is None:
                if site is not None:
                    site.fecharsite()
                site = web.TratarSite(os.getenv('SITENADACONSTA'), os.getenv('NOMEPROFILEIPTU'))
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
                    if site.url == os.getenv('SITENADACONSTA'):
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
                                    for i, boleto in enumerate(linkdownload, start=1):
                                        if botaogerar is not None:
                                            if getattr(sys, 'frozen', False):
                                                caminho_arquivo = site.monitorar_downloads_sem_href(
                                                    link_elemento=boleto,
                                                    timeout=120)

                                                if caminho_arquivo:
                                                    novonomearquivo = os.path.join(
                                                        objeto.pastadownload,
                                                        codigocliente) + (
                                                                          f"_{i - 1}.pdf" if i > 1 else ".pdf")

                                                    aux.adicionarcabecalhopdf(caminho_arquivo,
                                                                              novonomearquivo,
                                                                              codigocliente,
                                                                              codigobarras=False)
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
    dfdivida = None

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
    caminhodestino = 'certidao_' + str(codigocliente) + '_' + str(linha[NrIPTU]) + '.pdf'

    objeto.visual.mudartexto('labelstatus', 'Extraindo boleto...')

    pasta_com_divida = os.path.join(objeto.pastadownload, "Com_Divida")
    pasta_sem_divida = os.path.join(objeto.pastadownload, "Sem_Divida")

    # Criar pastas de destino
    os.makedirs(pasta_com_divida, exist_ok=True)
    os.makedirs(pasta_sem_divida, exist_ok=True)
    caminho_arquivo = aux.is_valid_file_in_path(caminhodestino, objeto.pastadownload)
    if not caminho_arquivo:
        while not problemacarregamento and not resolveu:
            textoerro = ''
            if site is not None:
                if site.navegador.current_url != os.getenv('SITECERTIDAOENFITEUTICA'):
                    site.fecharsite()
                    site = None
            if site is None:
                site = web.TratarSite(os.getenv('SITECERTIDAOENFITEUTICA'), os.getenv('NOMEPROFILEIPTU'))
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
                if site.url == os.getenv('SITECERTIDAOENFITEUTICA') and site.navegador.title != 'Request Rejected':
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
                                                if site.navegador.current_url == 'https://www2.rio.rj.gov.br/smf/iptucertfiscal/default.asp':
                                                    caminho_arquivo = os.path.join(objeto.pastadownload, caminhodestino)
                                                    site.salvar_pagina_como_pdf(caminho_arquivo)

                                                    if caminho_arquivo:
                                                        dfdivida = aux.extrairtextopdf(caminho_arquivo,'DIVIDAS')
                                                        if dfdivida is not None:
                                                            if dfdivida['Possui_Divida']:
                                                                novonomearquivo = os.path.join(objeto.pastadownload,
                                                                pasta_com_divida,
                                                                caminhodestino)
                                                            else:
                                                                novonomearquivo = os.path.join(objeto.pastadownload,
                                                                    pasta_sem_divida,
                                                                    caminhodestino)
                                                        else:
                                                            novonomearquivo = os.path.join(
                                                                caminhodestino,
                                                                caminhodestino)

                                                        if os.path.isfile(caminho_arquivo):
                                                            aux.adicionarcabecalhopdf(caminho_arquivo,
                                                                                          novonomearquivo,
                                                                                          aux.left(codigocliente, 4)
                                                                                      )
                                                    else:
                                                        dadosiptu = [codigocliente, linha[NrIPTU], anoextracao, 'Verificar (Extrair Manualmente)', dfdivida['Possui_Divida']]

                                            else:
                                                dadosiptu = [codigocliente, linha[NrIPTU], anoextracao, mensagemerro.text, 'N/A']

                                        else:
                                            dadosiptu = [codigocliente, linha[NrIPTU], anoextracao, 'Problema Captcha', 'N/A']
                                    else:
                                        dadosiptu = [codigocliente, linha[NrIPTU], anoextracao, 'Inscrição inválida!', 'N/A']
                                        resolveu = True
                    else:
                        dadosiptu = [codigocliente, linha[NrIPTU], anoextracao, 'Inscrição inválida!', 'N/A']
                        resolveu = True
    caminho_arquivo = caminhodestino if not caminho_arquivo else caminho_arquivo
    if os.path.isfile(caminho_arquivo):
        listacodigo = []
        # listatipopag = []
        listaarquivo = []
        caminho = caminhodestino if os.path.isfile(caminhodestino) else caminho_arquivo
        listacodigo.append("'" + codigocliente + "'")
        listaarquivo.append("'" + caminho + "'")

        df = {'Codigo': listacodigo, 'Arquivo': listaarquivo}
        df = pd.DataFrame(df)

    return dadosiptu, df

    # finally:
    #     if site is not None:
    #         site.fecharsite()


def extrairbombeiros(objeto, linha, dataatual=''):
    """
    Extrai boletos de bombeiros e interage com o CAPTCHA.
    """
    site = None
    dadosbombeiros = []
    df = None
    resolveucaptcha = False

    if dataatual == '':
        dataatual = aux.hora('America/Sao_Paulo', 'DATA')

    codigocliente = linha[Codigo]
    caminhodestino = os.path.join(objeto.pastadownload, f"{codigocliente}_{linha[NrCBM]}.pdf")

    # Inicializa o site
    site = web.TratarSite(os.getenv('SITECBM'), os.getenv('NOMEPROFILECBM'))

    if not os.path.isfile(caminhodestino):
        while not resolveucaptcha:
            site.abrirnavegador()

            # Buscar os campos de inscrição e dígito verificador
            inscricao_result = site.verificarobjetoexiste('NAME', 'cbmerj', buscar_em_iframes=True)
            dv_result = site.verificarobjetoexiste('NAME', 'cbmerj_dv', buscar_em_iframes=True)

            if inscricao_result and dv_result:
                # Processar inscrição
                iframe_inscricao, inscricao = inscricao_result if isinstance(inscricao_result, tuple) else (None, inscricao_result)
                iframe_dv, dv = dv_result if isinstance(dv_result, tuple) else (None, dv_result)

                if iframe_inscricao:
                    site.irparaframe(iframe_inscricao)
                    inscricao.clear()
                    inscricao.send_keys(aux.left(linha[NrCBM], 7))

                # if iframe_dv:
                #     site.irparaframe(iframe_dv)
                    dv.clear()
                    dv.send_keys(aux.right(linha[NrCBM], 1))

                # Pegar o elemento captcha
                elementocaptcha = site.buscarobjetoemdf({'id': 'recaptcha-token', 'nodeName': 'INPUT', 'value': '^NaN'}, 'baseURI')

                if elementocaptcha:
                    print("CAPTCHA detectado. Resolva manualmente no navegador.")

                    # Pausa o script para que o usuário resolva o CAPTCHA manualmente
                    input("Pressione Enter aqui após resolver o CAPTCHA e enviar no site...")

                    # Pegar o botão de envio
                    botaoenvio_result = site.verificarobjetoexiste('ID', 'btnEnviar', buscar_em_iframes=True)
                    iframe_botaoenvio, botaoenvio = botaoenvio_result if isinstance(botaoenvio_result, tuple) else (None, botaoenvio_result)

                    # Trocar para o iframe do botão, se necessário
                    if iframe_botaoenvio:
                        site.irparaframe(iframe_botaoenvio)

                    # Clicar no botão de envio
                    if botaoenvio:
                        botaoenvio.click()
                        resolveucaptcha = True
                    else:
                        print("Botão de envio não encontrado.")
                else:
                    print("Elemento 'g-recaptcha-response' não encontrado.")


            if not resolveucaptcha:
                print("Tentativa de resolver o CAPTCHA falhou. Tentando novamente...")

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
