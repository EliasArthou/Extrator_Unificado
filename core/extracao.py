"""
Extração em geral
"""

import datetime
import os

from core import boletos
from core import web
from auxiliares import utils as aux
from auxiliares import sensiveis as senha
import time
from auxiliares import messagebox as msg
import sys
from extratores.Condominios import condominios
from extratores.Prefeitura import Biptu
from extratores.Bombeiros import bombeiros_hibrido as Bh
import pandas as pd
from dotenv import load_dotenv


# Carrega as variáveis do arquivo .env
load_dotenv()

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
        self.datames = ''

    def controlaextracao(self):
        # try:
        self.visual.acertaconfjanela(False)

        if os.path.isfile(os.path.join(aux.caminhoprojeto(), 'Scai.WMB')):
            caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                                  tipoarquivos=[('Banco ' + os.getenv('EMPRESA'), '*.WMB'), ('Arquivos Excel', '*.xlsx'), ('Todos os Arquivos:', '*.*')],
                                                  caminhoini=aux.caminhoprojeto(), arquivoinicial='Scai.WMB')
        else:
            if os.path.isdir(aux.caminhoprojeto()):
                caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                                      tipoarquivos=[('Banco ' + os.getenv('EMPRESA'), '*.WMB'), ('Arquivos Excel', '*.xlsx'), ('Todos os Arquivos:', '*.*')],
                                                      caminhoini=aux.caminhoprojeto())
            else:
                caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                                      tipoarquivos=[('Banco ' + os.getenv('EMPRESA'), '*.WMB'), ('Todos os Arquivos:', '*.*')])

        if len(caminhobanco) == 0:
            msg.msgbox('Selecione o caminho do Banco de Dados!', msg.MB_OK, 'Erro Banco')
            self.visual.manipularradio(self.extracao, True)
            sys.exit()

        if not bool(self.visual.somentevalores.get()):
            self.pastadownload = aux.caminhoselecionado(3, 'Selecione o caminho dos arquivos:', caminhoini=self.pastadownload)

        # Padroniza subpasta por tipo de extração (Downloads/Bombeiros, Downloads/Prefeitura, Downloads/Condomínios)
        _subpasta_por_tipo = {
            'Bombeiros': 'Bombeiros',
            'Prefeitura': 'Prefeitura',
            'Condomínios': 'Condomínios',
        }.get(self.extracao, '')
        if _subpasta_por_tipo and not self.pastadownload.endswith(_subpasta_por_tipo):
            self.pastadownload = os.path.join(self.pastadownload, _subpasta_por_tipo)
            os.makedirs(self.pastadownload, exist_ok=True)

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
                self.sql = senha.SQLCBM

            case 'Prefeitura':
                if bool(self.visual.faltantes) and self.visual.tiposervico.get() == 'IPTU':
                    self.sql = senha.SQLIPTUFALTANTE
                else:
                    self.sql = senha.SQLIPTUCOMPLETO

            case 'Condomínios':
                if self.bd is None:
                    self.bd = aux.Banco(caminhobanco)

                if self.visual.iniciodomes.get():
                    self.bd.executarsql('DELETE * FROM BoletosCondominios')

                data = datetime.datetime.now()
                mes_ano = data.strftime("%m/%Y")
                self.visual.acertaconfjanela(False)
                self.datames = aux.criarinputbox('Mês dos Boletos', 'Data de extração dos boletos: (mm/aaaa)', valorinicial=mes_ano)
                self.visual.acertaconfjanela(True)
                self.sql = aux.retornarlistaboletos(anomes=self.datames)

            case _:
                msg.msgbox(f'Opção Inválida!', msg.MB_OK, 'Opção não reconhecida!')
        if self.bd is None:
            self.bd = aux.Banco(caminhobanco)

        if len(str(self.indicecliente)) > 0:
            if caminhobanco[-4:].lower() != 'xlsx':
                self.indicecliente = str(self.indicecliente).zfill(4)
                if self.indicecliente == '0000' or self.extracao != 'Prefeitura':
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
        Extrai boletos de bombeiros usando Playwright + AntiCaptcha.

        Migrado de Biptu.extrairbombeiros (Selenium, captcha manual via input())
        para bombeiros_hibrido.extrairbombeiros_hibrido (Playwright + AntiCaptcha
        automatico + fallback IPTU).

        Pass 3 (regras migradas do taxabombeiros legado):
          - Cota unica vs Parcelada: lida de self.resposta (1=Cota Unica), passada
            via parametro cota_unica do extrairbombeiros_hibrido.
          - Status temporal Vencido / Imposto Ano Corrente: calculado por debito
            em _extrair_debitos_bombeiros (campo status_temporal no dict).
          - Cidades especiais (Macae/SG/Campos): detectadas via mensagem
            "realize a consulta atraves do numero da inscricao" -> fallback IPTU.
        """
        from playwright.sync_api import sync_playwright

        self.listachaves = ['Cod Cliente', 'Nº CBMERJ', 'Área Construída', 'Utilização', 'Faixa', 'Proprietário', 'Endereço',
                            'taxa[anos_em_debito]', 'taxa[Exercicio]', 'taxa[Parcela]', 'taxa[Vencimento]', 'taxa[Valor]',
                            'taxa[Mora]', 'taxa[Total]', 'Status']
        self.listaexcel = []

        # Perfil Chrome dedicado pro CBM (mesma variavel que web.py usava)
        profile_name = os.getenv("NOMEPROFILECBM", "cbm_profile")
        profile_path = os.path.join(aux.caminhoprojeto(), "Profile", profile_name)
        os.makedirs(profile_path, exist_ok=True)

        with sync_playwright() as p:
            context = p.chromium.launch_persistent_context(
                user_data_dir=profile_path,
                headless=True,
                accept_downloads=True,
            )
            page = context.pages[0] if context.pages else context.new_page()

            try:
                for indice, linha in enumerate(self.resultado):
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

                    # Cota Unica vs Parcelada: vem do radio da janela
                    # (tipopagamento=1 = Cota Unica, 2 = Parcelado)
                    cota_unica = str(getattr(self, 'resposta', '')) == '1'

                    dadosiptu, df = Bh.extrairbombeiros_hibrido(
                        page, self, linha,
                        aux.hora('America/Sao_Paulo', 'DATA'),
                        cota_unica=cota_unica,
                    )
            finally:
                try:
                    context.close()
                except Exception:
                    pass

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

        self.listachaves = ['Código', 'Login', 'Senha', 'Administradora', 'Condomínio', 'Unidade', 'Login Múltiplo', 'Resposta', 'Check de Arquivo', 'CheckErro', 'Nome Função', 'Problema Login']
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
