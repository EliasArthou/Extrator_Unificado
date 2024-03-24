"""
Extração em geral
"""

import datetime
import os

import boletos
import web
import auxiliares as aux
import sensiveis as senha
import time
import messagebox as msg
import sys
import condominios
import Biptu
import pandas as pd


class Extrator:
    def __init__(self, visual):
        self.visual = visual
        self.extracao = self.visual.tipoextracao.get()
        self.site = None
        self.bd = None
        self.listaexcel = []
        self.texto = ''
        self.tempoinicio = time.time()
        self.tempofim = ''
        self.codigocliente = ''
        self.sql = ''
        self.indicecliente = ''
        self.resultado = []
        self.opcoesextracao = -1
        self.resposta = -1
        self.pastadownload = aux.caminhoprojeto() + '\\' + 'Downloads'
        self.listachaves = []
        self.listadados = []

    def controlaextracao(self):
        # try:
        self.visual.acertaconfjanela(False)

        if os.path.isfile(os.path.join(aux.caminhoprojeto(), 'Scai.WMB')):
            caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                                  tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Todos os Arquivos:', '*.*')],
                                                  caminhoini=aux.caminhoprojeto(), arquivoinicial='Scai.WMB')
        else:
            if os.path.isdir(aux.caminhoprojeto()):
                caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                                      tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Todos os Arquivos:', '*.*')],
                                                      caminhoini=aux.caminhoprojeto())
            else:
                caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                                      tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Todos os Arquivos:', '*.*')])

        if len(caminhobanco) == 0:
            msg.msgbox('Selecione o caminho do Banco de Dados!', msg.MB_OK, 'Erro Banco')
            self.visual.manipularradio(self.extracao, True)
            sys.exit()

        if not bool(self.visual.somentevalores.get()):
            self.pastadownload = aux.caminhoselecionado(3, 'Selecione o caminho dos arquivos:', caminhoini=self.pastadownload)

        self.resposta = str(self.visual.tipopagamento.get())
        if self.visual.tipoextracao.get() != 'Condomínios':
            self.indicecliente = aux.criarinputbox('Cliente de Corte', 'Iniciar a partir de um cliente? (0 fará de todos da lista)', valorinicial='0')
        else:
            self.indicecliente = '0'

            if self.indicecliente is not None:
                if not str(self.indicecliente).isdigit():
                    msg.msgbox('Digite um valor válido (precisa ser numérico)!', msg.MB_OK, 'Opção Inválida')
                    self.visual.manipularradio(self.extracao, True)
                    sys.exit()
            else:
                msg.msgbox('Digite o inicío desejado ou deixe 0 (Zero)!', msg.MB_OK, 'Opção Inválida!')
                self.visual.manipularradio(self.extracao, True)
                sys.exit()

        self.tempoinicio = time.time()

        self.visual.acertaconfjanela(True)

        self.visual.mudartexto('labelstatus', 'Executando pesquisa no banco...')

        match self.extracao:
            case 'Bombeiros':
                self.sql = senha.sqlcbm

            case 'Prefeitura':
                if bool(self.visual.faltantes) and self.visual.tiposervico.get() == 'IPTU':
                    self.sql = senha.sqliptufaltante
                else:
                    self.sql = senha.sqliptucompleto

            case 'Condomínios':
                if self.bd is None:
                    self.bd = aux.Banco(caminhobanco)

                if self.visual.iniciodomes:
                    self.bd.executarsql('DELETE * FROM BoletosCondominios')

                self.sql = aux.retornarlistaboletos()

            case _:
                msg.msgbox(f'Opção Inválida!', msg.MB_OK, 'Opção não reconhecida!')
        if self.bd is None:
            self.bd = aux.Banco(caminhobanco)

        if len(str(self.indicecliente)) > 0:
            self.indicecliente = str(self.indicecliente).zfill(4)
            if self.indicecliente == '0000' or self.visual.tiposervico.get() != 'IPTU':
                self.resultado = self.bd.consultar(self.sql)
            else:
                self.resultado = self.bd.consultar(self.sql.replace(';', ' ') + "WHERE Codigo >= '{codigo}' ORDER BY Codigo;".format(codigo=self.indicecliente))

            if self.resultado:
                self.resultado.sort(key=lambda x: x[0])

        self.bd.fecharbanco()

        match self.extracao:
            case 'Bombeiros':
                self.extrairboletosbombeiro()

            case 'Prefeitura':
                self.extrairboletos()

            case 'Condomínios':

                self.extraircondominio()

            case _:
                msg.msgbox(f'Extração não configurada!', msg.MB_OK, 'Extração não reconhecida!')

        # Verifica se o browser está aberto
        if self.site:
            # Fecha o browser
            self.site.fecharsite()

        self.criarlog()
        # finally:

        # Verifica se o browser está aberto
        if self.site is not None:
            # Fecha o browser
            self.site.fecharsite()

        self.tempofim = time.time()

        hours, rem = divmod(self.tempofim - self.tempoinicio, 3600)
        minutes, seconds = divmod(rem, 60)

        self.visual.manipularradio(self.extracao, True)
        self.visual.acertaconfjanela(False)

        msg.msgbox(f'O tempo decorrido foi de: {"{:0>2}:{:0>2}:{:05.2f}".format(int(hours), int(minutes), int(seconds))}', msg.MB_OK, 'Tempo Decorrido')

    def extrairboletosbombeiro(self):
        """
        : param caminhobanco: caminho do banco para realizar a pesquisa.
        : param resposta: opção selecionada de extração.
        : param visual: janela a ser manipulada.
        """

        # try:
        # gerarboleto = not self.visual.somentevalores.get()

        self.listachaves = ['Cod Cliente', 'Nº CBMERJ', 'Área Construída', 'Utilização', 'Faixa', 'Proprietário', 'Endereço',
                            'taxa[anos_em_debito]', 'taxa[Exercicio]', 'taxa[Parcela]', 'taxa[Vencimento]', 'taxa[Valor]',
                            'taxa[Mora]', 'taxa[Total]', 'Status']
        self.listaexcel = []
        # site = web.TratarSite(senha.siteCBM, senha.nomeprofileCBM)

        for indice, linha in enumerate(self.resultado):
            resolveucaptcha = False
            if aux.hora('America/Sao_Paulo', 'HORA') < datetime.time(22, 00, 00) and self.texto != 'Este serviço encontra-se temporariamente indisponível.':
                codigocliente = linha[Biptu.Codigo]
                # ==================== Parte Gráfica =======================================================
                self.visual.mudartexto('labelcodigocliente', 'Código Cliente: ' + codigocliente)
                cbm = str(linha['cbm'])
                cbm = str(cbm.strip()).zfill(8)
                cbm = '{}{}{}{}{}{}{}-{}'.format(*cbm)
                self.visual.mudartexto('labelinscricao', 'Inscrição: ' + cbm)
                self.visual.mudartexto('labelquantidade', 'Item ' + str(indice + 1) + ' de ' + str(len(self.resultado)) + '...')
                self.visual.mudartexto('labelstatus', 'Extraindo boleto...')
                # Atualiza a barra de progresso das transações (Views)
                self.visual.configurarbarra('barraextracao', len(self.resultado), indice + 1)
                time.sleep(0.1)
                # ==================== Parte Gráfica =======================================================
                # Verifica se o CAPTCHA foi resolvido e se não é depois das 22:00 (o site não gera boleto depois desse horário) para
                # continuar tentando resolver o CAPTCHA (caso não tenha sido resolvido) em "looping"
                while not resolveucaptcha and aux.hora('America/Sao_Paulo', 'HORA') < datetime.time(22, 00, 00):
                    # Variável que vai receber o texto de erro do site (caso exista)
                    self.texto = ''
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
                            inscricao = site.verificarobjetoexiste('NAME', 'num_cbmerj')
                            # Campo de dígito verificador
                            dv = site.verificarobjetoexiste('NAME', 'dv_cbmerj')
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
                                                self.listaexcel.append(linhanova)
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
                                                                        self.visual.mudartexto('labelstatus', 'Salvando Boleto...')
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
                                                                                caminhodestino = self.pastadownload + '/' + listaanos[indicebotao - 1] + '_' + codigocliente + '_' + \
                                                                                                 linha['cbm'] + '.pdf'
                                                                            else:
                                                                                caminhodestino = self.pastadownload + '/' + listaanos[indicebotao - 1] + '_' + codigocliente + '_' + \
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
                                                            if self.resposta == 1 and indiceano == 1 and listaanos[len(listaanos) - 1] == listaanos[indicebotao - 1]:
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
                                            if len(dadoscbm) == len(self.listachaves):
                                                self.listaexcel.append(dict(zip(self.listachaves, dadoscbm)))
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
                self.visual.acertaconfjanela(False)
                # Texto quando o serviço está indisponível
                if self.texto != 'Este serviço encontra-se temporariamente indisponível.':
                    # Caso o erro não seja de serviço indisponível o horário é inválido
                    msg.msgbox('Impossível gerar boletos depois das 22:00!', msg.MB_OK, 'Horário Inválido')
                else:
                    # Mensagem de serviço indisponível
                    msg.msgbox('Serviço fora do ar!', msg.MB_OK, 'Serviço com problemas')

                # Sai do Looping
                break

        # except Exception as e:
        #     with open("Log_" + aux.acertardataatual() + ".txt", "a") as myfile:
        #         myfile.write(str(e))
        #     msg.msgbox("Erro! Log salvo em: " + "Log_" + aux.acertardataatual() + ".txt", msg.MB_OK, 'Erro')

    def extrairboletos(self):
        """

        : param caminhobanco: caminho do banco para realizar a pesquisa.
        : param resposta: opção selecionada de extração.
        : param visual: janela a ser manipulada.
        """

        df = None
        dadosiptu = None
        linhatemp = []

        # try:

        self.listaexcel = []
        dataatual = aux.hora('America/Sao_Paulo', 'DATA')
        anoatual = dataatual.year

        for indice, linha in enumerate(self.resultado):
            codigocliente = linha[Biptu.Codigo]
            # ===================================== Parte Gráfica =======================================================
            self.visual.mudartexto('labelcodigocliente', 'Código Cliente: ' + codigocliente)
            iptu = str(linha[Biptu.NrIPTU])
            iptu = str(iptu.strip()).zfill(8)
            iptu = '{}.{}{}{}.{}{}{}-{}'.format(*iptu)
            self.visual.mudartexto('labelinscricao', 'Inscrição Imobiliária: ' + iptu)

            self.visual.mudartexto('labelquantidade', 'Item ' + str(indice + 1) + ' de ' + str(len(self.resultado)) + '...')
            self.visual.mudartexto('labelstatus', 'Extraindo boleto...')
            # Atualiza a barra de progresso das transações (Views)
            self.visual.configurarbarra('barraextracao', len(self.resultado), indice + 1)
            # ===================================== Parte Gráfica =======================================================
            match self.visual.tiposervico.get():
                case 'IPTU':
                    self.listachaves = ['Código Cliente', 'Inscrição', 'Guia do Exercício', 'Nr Guia', 'Valor', 'Contribuinte', 'Endereço', 'Status']
                    dadosiptu, df = Biptu.extrairboletos(self, linha)
                    if df is not None:
                        self.bd.adicionardf('Codigos IPTUs', df, 7)
                case 'Nada Consta':
                    self.listachaves = ['Código Cliente', 'Inscrição', 'Guia do Exercício', 'Status']
                    dadosiptu, df = Biptu.extrairnadaconsta(self, linha, dataatual)
                case 'Certidão Negativa':
                    dadosiptu, df = Biptu.extraircertidaonegativa(self, linha, dataatual)
                    # self.extrairboletosiptu()
                    print('Teste')

            if len(dadosiptu) > 0:
                self.listadados.extend(sublist for sublist in dadosiptu)
            else:
                linhatemp = linha[:2]
                linhatemp.append(anoatual)
                linhatemp.append('Já extraído!')

                self.listadados.append(linhatemp)

        # except Exception as e:
        #     # print(f"Erro:{str(e)}")
        #     msg.msgbox(str(e), msg.MB_OK, 'Erro')

    def extraircondominio(self):
        listaboletos = []
        self.visual.acertaconfjanela(False)

        self.listachaves = ['Código', 'Login', 'Senha', 'Administradora', 'Condomínio', 'Unidade', 'Resposta', 'Check de Arquivo', 'CheckErro', 'Nome Função', 'Problema Login']
        self.listaexcel = []

        self.resultado = self.resultado
        for indice, linha in enumerate(self.resultado):
            codigocliente = linha[condominios.identificador]
            # ===================================== Parte Gráfica =======================================================
            self.visual.mudartexto('labelcodigocliente', 'Código Cliente: ' + codigocliente)
            self.visual.mudartexto('labelinscricao', '')
            self.visual.mudartexto('labelquantidade', 'Item ' + str(indice + 1) + ' de ' + str(len(self.resultado)) + '...')
            self.visual.mudartexto('labelstatus', 'Extraindo boleto...')
            # Atualiza a barra de progresso das transações (Views)
            self.visual.configurarbarra('barraextracao', len(self.resultado), indice + 1)
            time.sleep(0.1)
            # ===================================== Parte Gráfica =======================================================
            if len(linha[condominios.Usuario].strip()) > 0 and len(linha[condominios.Senha].strip()) > 0:
                multiplas = aux.encontrar_administradora(linha[condominios.Administradora])
                if multiplas is None:
                    listatemp = getattr(condominios, senha.retornaradministradora('nomereal', linha[condominios.Administradora], 'nomereduzido').lower())(self, linha)
                    if listatemp is not None and len(listatemp) > 0:
                        listaboletos = listaboletos + listatemp
                    # listaboletos.append(getattr(condominios, senha.retornaradministradora('nomereal', linha[condominios.Administradora], 'nomereduzido').lower())(self, linha))
                else:
                    listatemp = getattr(condominios, multiplas['Site'])(self, linha)
                    if listatemp is not None and len(listatemp) > 0:
                        listaboletos = listaboletos + listatemp
                listatemp = None

                # listaboletos.append(getattr(condominios, multiplas['Site'])(self, linha))

            else:
                linha[condominios.Resposta] = 'Usuário e/ou senha não preenchido!'
                linha[condominios.ProblemaLogin] = True
                if len(linha[condominios.Resposta]) == 0:
                    linha[condominios.Resposta] = condominios.mensagemerropadrao
                self.visual.mudartexto('labelstatus', linha[condominios.Resposta])

        listatemp = aux.get_files_not_in_list(listaboletos, self.pastadownload)
        for item in listatemp:
            itemtemp = boletos.barcodereader(item)
            if itemtemp is not None:
                listaboletos.append(itemtemp)

        if listaboletos is not None:
            # Convertendo a lista de dicionários em um DataFrame
            df = pd.DataFrame(listaboletos)

            self.bd.adicionardf('BoletosCondominios', df, 1)
            if os.path.isfile(aux.caminhoprojeto()+'\\saida.xlsx'):
                # Salvando o DataFrame em um arquivo Excel
                df.to_excel('saida.xlsx', index=False)

    def criarlistadicionarios(self):
        # Verifica se tem dados e cabeçalhos nas respectivas linhas e se as mesmas têm a mesma quantidade de colunas
        if len(self.listachaves) > 0 and len(self.resultado) > 0 and all(len(sublist) == len(self.listachaves) for sublist in self.resultado):
            lista_de_dicionarios = []
            for valores_correspondentes in self.resultado:
                novo_dicionario = {chave: valor for chave, valor in zip(self.listachaves, valores_correspondentes)}
                lista_de_dicionarios.append(novo_dicionario)
            return lista_de_dicionarios
        else:
            return None

    def criarlog(self):
        self.listaexcel = self.criarlistadicionarios()
        if self.listaexcel is not None:
            if len(self.listaexcel) > 0:
                # Atualiza o "status" na tela de usuário
                self.visual.mudartexto('labelstatus', 'Salvando lista...')
                # Escreve no arquivo a lista e salva o Excel
                nomearquivo = 'Log_' + aux.acertardataatual() + '.xlsx'
                aux.escreverlistaexcelog(nomearquivo, self.listaexcel)
