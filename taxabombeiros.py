import datetime
import web
import auxiliares as aux
import sensiveis as senha
import sys
import messagebox as msg


Codigo = 0
NrIPTU = 1
IndiceIPTU = 2

tempoesperadownload = 180


def taxabombeiro(objeto, linha):
    """
    : param caminhobanco: caminho do banco para realizar a pesquisa.
    : param resposta: opção selecionada de extração.
    : param visual: janela a ser manipulada.
    """
    resolveucaptcha = False

    try:
        gerarboleto = not objeto.somentevalores.get()

        # objeto.listachaves = ['Cod Cliente', 'Nº CBMERJ', 'Área Construída', 'Utilização', 'Faixa', 'Proprietário', 'Endereço',
        #                     'taxa[anos_em_debito]', 'taxa[Exercicio]', 'taxa[Parcela]', 'taxa[Vencimento]', 'taxa[Valor]',
        #                     'taxa[Mora]', 'taxa[Total]', 'Status']
        # objeto.listaexcel = []
        site = web.TratarSite(senha.siteCBM, senha.nomeprofileCBM)

        # for indice, linha in enumerate(objeto.resultado):

        if aux.hora('America/Sao_Paulo', 'HORA') < datetime.time(22, 00, 00) and objeto.texto != 'Este serviço encontra-se temporariamente indisponível.':
            codigocliente = linha[Codigo]
            # ==================== Parte Gráfica =======================================================
            objeto.visual.mudartexto('labelcodigocliente', 'Código Cliente: ' + codigocliente)
            cbm = str(linha['cbm'])
            objeto.visual.mudartexto('labelstatus', 'Extraindo boleto...')
            # ==================== Parte Gráfica =======================================================
            # Verifica se o CAPTCHA foi resolvido e se não é depois das 22:00 (o site não gera boleto depois desse horário) para
            # continuar tentando resolver o CAPTCHA (caso não tenha sido resolvido) em "looping"
            while not resolveucaptcha and aux.hora('America/Sao_Paulo', 'HORA') < datetime.time(22, 00, 00):
                # Variável que vai receber o texto de erro do site (caso exista)
                objeto.texto = ''
                # Verifica a hora para entrar no site, caso esteja fora do horário válido, nem inicia
                if aux.hora('America/Sao_Paulo', 'HORA') < datetime.time(22, 00, 00):
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
                        inscricao = site.verificarobjetoexiste('NAME', 'cbmerj')
                        # Campo de dígito verificador
                        dv = site.verificarobjetoexiste('NAME', 'cbmerj_dv')
                        # Elemento da chave do captcha
                        elementocaptcha = site.verificarobjetoexiste('ID', 'recaptcha-token')
                        # Elemento preenchido para a resposta do captcha
                        respostacaptcha = site.verificarobjetoexiste('ID', 'g-recaptcha-response')
                        # Testa se tem os dois campos supracitados
                        if inscricao is not None and dv is not None:
                            # "Limpa" o campo de inscrição
                            inscricao.clear()
                            # Preenche o campo de inscrição com os dados do banco de dados (sem o dígito verificador)
                            inscricao.send_keys(aux.left(linha['cbm'], 7))
                            # "Limpa" o campo do dígito verificador
                            dv.clear()
                            # Preenche o campo do dígito verificador com os dados do banco de dados
                            dv.send_keys(aux.right(linha['cbm'], 1))
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
                                        campoiptu.send_keys(linha['iptu'])
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
                                        if int(linhacobranca.get('taxa[Exercicio]')) < aux.hora('America/Sao_Paulo', 'DATA').year - 1:
                                            linhanova.update({'Status': 'Vencido'})
                                        else:
                                            linhanova.update({'Status': 'Imposto Ano Corrente'})

                                        # Insere a linha formada da junção das listas no Excel
                                        objeto.listaexcel.append(linhanova)
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
                                                                # Pasta Download inicial
                                                                pastadownloadinicial = aux.caminhospadroes(80)
                                                                # Espera o download
                                                                site.esperadownloads(pastadownloadinicial, 10)
                                                                # Verifica o último arquivo baixado
                                                                baixado = aux.ultimoarquivo(pastadownloadinicial, 'pdf')
                                                                # Verifica se o arquivo baixado vem do site
                                                                if 'modules' not in baixado and 'Taxa de Incêndio' not in baixado:
                                                                    # Caso não seja do site ele "limpa" a informação do último arquivo da pasta
                                                                    baixado = ''

                                                                # Verifica se o arquivo baixado vem do site para renomear, adicionar o cabeçalho
                                                                # e mover para a pasta de "download" definida
                                                                if len(baixado) > 0:
                                                                    # Define o nome do arquivo (adicionando o índice de ano quando necessário)
                                                                    if indiceano == 0:
                                                                        caminhodestino = objeto.pastadownload + '/' + listaanos[indicebotao - 1] + '_' + codigocliente + '_' + \
                                                                                         linha['cbm'] + '.pdf'
                                                                    else:
                                                                        caminhodestino = objeto.pastadownload + '/' + listaanos[indicebotao - 1] + '_' + codigocliente + '_' + \
                                                                                         linha['cbm'] + '_' + str(indiceano) + '.pdf'

                                                                    # Trata o caminho para ignorar caracteres especiais no endereço
                                                                    caminhodestino = aux.to_raw(caminhodestino)
                                                                    # Adiciona o código do cliente ao cabeçalho e move o arquivo para o caminho de destino
                                                                    aux.adicionarcabecalhopdf(baixado, caminhodestino, codigocliente)
                                                        # Verifica se tem abas abertas de boletos
                                                        while site.num_abas() > 1:
                                                            # Vai para última aba para não fechar a aba que tem os botões de gerar boleto
                                                            site.irparaaba(site.num_abas())
                                                            # Fecha as abas de boletos
                                                            site.fecharaba()
                                                            # Vai para última aba para não fechar a aba que tem os botões de gerar boleto
                                                            site.irparaaba(site.num_abas())
                                                    # Se a opção de boleto único for marcada ele não continua extraindo os boletos do parcelamento
                                                    if objeto.resposta == 1 and indiceano == 1 and listaanos[len(listaanos) - 1] == listaanos[indicebotao - 1]:
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
                                    dadoscbm = [codigocliente, aux.left(str(linha['cbm']), 7) + '-' + aux.right(str(linha['cbm']), 1), '', '', '', '', '', '', '', '', '',
                                                '', '', '', mensagemerro]
                                    # Adiciona o item com o cabeçalho
                                    if len(dadoscbm) == len(objeto.listachaves):
                                        objeto.listaexcel.append(dict(zip(objeto.listachaves, dadoscbm)))
                                    else:
                                        print(mensagemerro)
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
            if objeto.texto != 'Este serviço encontra-se temporariamente indisponível.':
                # Caso o erro não seja de serviço indisponível o horário é inválido
                msg.msgbox('Impossível gerar boletos depois das 22:00!', msg.MB_OK, 'Horário Inválido')
            else:
                # Mensagem de serviço indisponível
                msg.msgbox('Serviço fora do ar!', msg.MB_OK, 'Serviço com problemas')

    except Exception as e:
        with open("Log_" + aux.acertardataatual() + ".txt", "a") as myfile:
            myfile.write(str(e))
        msg.msgbox("Erro! Log salvo em: " + "Log_" + aux.acertardataatual() + ".txt", msg.MB_OK, 'Erro')
