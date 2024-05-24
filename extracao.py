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
import taxabombeiros as bombeiro
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
                                                  tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Arquivos Excel', '*.xlsx'), ('Todos os Arquivos:', '*.*')],
                                                  caminhoini=aux.caminhoprojeto(), arquivoinicial='Scai.WMB')
        else:
            if os.path.isdir(aux.caminhoprojeto()):
                caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                                      tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Arquivos Excel', '*.xlsx'), ('Todos os Arquivos:', '*.*')],
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
        if self.visual.tipoextracao.get() != 'Condomínios' and caminhobanco[-4:].lower() != 'xlsx':
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
                if bool(self.visual.faltantes):  # and self.visual.tiposervico.get() == 'IPTU':
                    self.sql = senha.sqliptufaltante
                else:
                    self.sql = senha.sqliptucompleto

            case 'Condomínios':
                if self.bd is None:
                    self.bd = aux.Banco(caminhobanco)

                if self.visual.iniciodomes.get():
                    self.bd.executarsql('DELETE * FROM BoletosCondominios')

                self.sql = aux.retornarlistaboletos()

            case _:
                msg.msgbox(f'Opção Inválida!', msg.MB_OK, 'Opção não reconhecida!')
        if self.bd is None:
            self.bd = aux.Banco(caminhobanco)

        if len(str(self.indicecliente)) > 0:
            if caminhobanco[-4:].lower() != 'xlsx':
                self.indicecliente = str(self.indicecliente).zfill(4)
                if self.indicecliente == '0000' or self.visual.tipoextracao.get() != 'Prefeitura':  # or self.visual.tiposervico.get() != 'IPTU':
                    self.resultado = self.bd.consultar(self.sql)
                else:
                    self.resultado = self.bd.consultar(self.sql.replace(';', ' ') + "WHERE Codigo >= '{codigo}' ORDER BY Codigo;".format(codigo=self.indicecliente))
            else:
                # Lendo o arquivo Excel e transformando-o em um DataFrame
                self.resultado = self.bd.ler_excel(caminhobanco)

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
                cbm = str(linha[Biptu.NrCBM])
                cbm = str(cbm.strip()).zfill(8)
                cbm = '{}{}{}{}{}{}{}-{}'.format(*cbm)
                self.visual.mudartexto('labelinscricao', 'Inscrição: ' + cbm)
                self.visual.mudartexto('labelquantidade', 'Item ' + str(indice + 1) + ' de ' + str(len(self.resultado)) + '...')
                self.visual.mudartexto('labelstatus', 'Extraindo boleto...')
                # Atualiza a barra de progresso das transações (Views)
                self.visual.configurarbarra('barraextracao', len(self.resultado), indice + 1)
                time.sleep(0.1)
                self.texto = ''
                # Verifica a hora para entrar no site, caso esteja fora do horário válido, nem inicia
                if aux.hora('America/Sao_Paulo', 'HORA') < datetime.time(22, 00, 00):
                    dadosiptu, df = Biptu.extrairbombeiros(self, linha, aux.hora('America/Sao_Paulo', 'DATA'))

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
                        # Remover as aspas simples de todos os valores do DataFrame
                        df = df.applymap(lambda x: x.replace("'", "") if isinstance(x, str) else x)
                        # Remover vírgulas dos valores e substituir ponto por vírgula
                        df['Valor Total'] = df['Valor Total'].str.replace('.', '').str.replace(',', '.')
                        # Converter a coluna 'Valor Total' para tipo float para representar valores monetários
                        df['Valor Total'] = df['Valor Total'].astype('float')
                        # Remover as aspas dos valores
                        df['Parcela'] = df['Parcela'].str.replace("'", "")
                        df['Parcela'] = df['Parcela'].astype(int)
                        self.bd.adicionardf('[Codigos IPTUs]', df, 7)
                case 'Nada Consta':
                    self.listachaves = ['Código Cliente', 'Inscrição', 'Guia do Exercício', 'Status']
                    dadosiptu, df = Biptu.extrairnadaconsta(self, linha, dataatual)
                case 'Certidão Negativa':
                    dadosiptu, df = Biptu.extraircertidaonegativa(self, linha, dataatual)

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
