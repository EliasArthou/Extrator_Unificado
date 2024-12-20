"""
Todas as funções de suporte
"""

import os
import sys
import time
import pyodbc
import glob
import shutil
import sensiveis as senha
import requests
import pandas as pd
import win32gui
import boletos
from sqlalchemy import create_engine
import openpyxl
import ctypes
from fuzzywuzzy import fuzz

# Constantes para SetThreadExecutionState
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

caminho = ''


class Banco:
    def __init__(self, caminho):
        self.conxn = None
        self.cursor = None
        self.engine = None
        self.erro = ''
        self.constr = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + caminho + ';Pwd=' + senha.senhabanco
        self.abrirconexao()

    def abrirconexao(self):
        try:
            if self.conxn is not None:
                try:
                    self.cursor.execute("SELECT 1")
                    return
                except pyodbc.Error:
                    self.conxn = None

            if len(self.constr) > 0:
                self.conxn = pyodbc.connect(self.constr)
                self.cursor = self.conxn.cursor()
                self.cursor.execute("SELECT 1")
                self.cursor = self.conxn.cursor()
                self.engine = create_engine(f"mssql+pyodbc:///?odbc_connect={self.constr}")

        except pyodbc.Error as e:
            self.erro = str(e)

    def ler_excel(self, caminho_arquivo):
        # Abre o arquivo Excel
        workbook = openpyxl.load_workbook(caminho_arquivo)

        # Seleciona a primeira planilha
        sheet = workbook.active

        # Inicializa uma lista para armazenar os dados do Excel
        dados_excel = []

        # Itera sobre as linhas da planilha
        for row in sheet.iter_rows(min_row=2, values_only=True):
            # Converte os valores None para strings vazias
            row = ["" if value is None else value for value in row]
            # Adiciona cada linha como uma lista à lista de dados_excel
            dados_excel.append(list(row))

        # Retorna os dados lidos do Excel como uma lista de listas
        return dados_excel

    def consultar(self, sql):
        self.abrirconexao()
        self.cursor.execute(sql)
        resultado = self.cursor.fetchall()
        resultado = [
            list(str(value).strip() if isinstance(value, str) else value for value in row)
            for row in resultado
        ]
        return resultado

    def obter_nome_coluna_por_indice(self, tabela, indice):
        """
        Obtém o nome da coluna da tabela do Access com base no índice.
        """
        # Obter informações sobre a estrutura da tabela
        estrutura_tabela = self.cursor.columns(table=tabela)

        # Percorrer as informações da estrutura da tabela para encontrar o nome da coluna com base no índice
        for coluna_info in estrutura_tabela:
            if coluna_info.ordinal == indice + 1:
                return coluna_info.column_name

        # Se o índice fornecido estiver fora do intervalo, retorne None
        return None

    def adicionardf(self, tabela, df, indicelimpeza=-1):
        try:
            self.abrirconexao()

            # Exclui os registros existentes na tabela com base no índice indicado
            if 0 <= indicelimpeza < len(df.columns):
                nome_coluna_df = df.columns[indicelimpeza]
                colunatabela = self.obter_nome_coluna_por_indice(tabela, indicelimpeza)

                if colunatabela is not None:
                    # Consulta para excluir os registros existentes na tabela que têm chaves primárias correspondentes aos registros no dataframe
                    # consulta_delete = f"DELETE FROM {tabela} WHERE {colunatabela} IN ({','.join(['?'] * len(df))})"
                    # self.cursor.execute(consulta_delete, tuple(df[nome_coluna_df]))
                    # A df[nome_coluna_df] contêm os valores que quero excluir
                    valores_unicos = tuple(df[nome_coluna_df].unique())  # Isso garante que cada valor seja único
                    placeholders = ','.join(['?'] * len(valores_unicos))  # Cria um placeholder para cada valor único

                    consulta_delete = f"DELETE FROM {tabela} WHERE {colunatabela} IN ({placeholders})"
                    self.cursor.execute(consulta_delete, valores_unicos)

            # Salvar o dataframe na tabela do banco de dados
            for index, row in df.iterrows():
                values = tuple(row)
                placeholders = ','.join(['?'] * len(values))
                insert_sql = f"INSERT INTO {tabela} VALUES ({placeholders})"
                try:
                    self.cursor.execute(insert_sql, values)
                    self.conxn.commit()
                except Exception as e:
                    print(f"Erro ao inserir a linha {index}: {row}")
                    print(f"Detalhe do erro: {e}")
                    self.conxn.rollback()

                    # df.to_sql(tabela, self.conxn, if_exists='append', index=False)

        except pyodbc.Error as e:
            print("Erro ao inserir DataFrame na tabela:", e)

    def executarsql(self, sql):
        try:
            self.abrirconexao()
            self.cursor.execute(sql)
            rows_affected = self.cursor.rowcount
            self.conxn.commit()
            return rows_affected

        finally:
            self.conxn.close()
            self.conxn = None

    def fecharbanco(self):
        if self.cursor:
            self.cursor.close()
        if self.conxn:
            self.conxn.close()


def numero_para_romano(numero):
    valores = [
        1000, 900, 500, 400,
        100, 90, 50, 40,
        10, 9, 5, 4,
        1
    ]
    simbolos = [
        "M", "CM", "D", "CD",
        "C", "XC", "L", "XL",
        "X", "IX", "V", "IV",
        "I"
    ]
    romano = ''
    i = 0
    while numero > 0:
        for _ in range(numero // valores[i]):
            romano += simbolos[i]
            numero -= valores[i]
        i += 1
    return romano


# class Banco:
#     """
#     Criado para se conectar e realizar operações no banco de dados
#     """
#
#     def __init__(self, caminho):
#         self.conxn = None
#         self.cursor = None
#         self.erro = ''
#         self.engine = None
#         self.constr = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + caminho + ';Pwd=' + senha.senhabanco
#         # Crie a string de conexão
#         self.params = quote_plus(self.constr)
#
#         self.abrirconexao()
#
#     def abrirconexao(self):
#         try:
#             # Verifica se a conexão já existe e está aberta
#             if self.conxn is not None:
#                 try:
#                     # Tenta executar um comando simples para verificar a conexão
#                     self.cursor.execute("SELECT 1")
#                     return  # Conexão está ativa, então não precisa abrir outra
#                 except pyodbc.Error:
#                     # A conexão existente falhou, então tentaremos reconectar
#                     self.conxn = None  # Resetar a conexão para tentar novamente
#
#             if len(self.constr) > 0:
#                 self.engine = create_engine("mssql+pyodbc:///?odbc_connect=%s" % self.params)
#                 self.conxn = pyodbc.connect(self.constr)
#                 self.cursor = self.conxn.cursor()
#                 self.cursor.execute("SELECT 1")
#                 # self.cursor.close()
#                 self.cursor = self.conxn.cursor()
#
#         except pyodbc.Error as e:
#             self.erro = str(e)
#
#     def consultar(self, sql):
#         """
#         :param sql: código sql a ser executado (uma consulta SQL).
#         :return: o resultado da consulta em uma lista.
#         """
#         self.abrirconexao()
#
#         self.cursor.execute(sql)
#         resultado = self.cursor.fetchall()
#         resultado = [
#             list(str(value).strip() if isinstance(value, str) else value for value in row)
#             for row in resultado
#         ]
#         return resultado
#
#     def adicionardf(self, tabela, df, indicelimpeza=-1):
#         # try:
#         self.abrirconexao()
#
#         # Se indicelimpeza for fornecido, limpe os registros antes de adicionar novos
#         # if indicelimpeza != -1:
#         #     for linha in df.values:
#         #         barras = str(linha[indicelimpeza])
#         #         # Constrói a instrução SQL DELETE
#         #         strSQL = "DELETE FROM [{0}] WHERE Barras = ?".format(tabela)
#         #
#         #         # Executa a consulta DELETE com a variável 'barras' como parâmetro
#         #         self.cursor.execute(strSQL, (barras,))
#         #     self.conxn.commit()
#         # Consulta para excluir os registros existentes na tabela que têm chaves primárias correspondentes aos registros no dataframe
#         consulta_delete = f"DELETE FROM {tabela} WHERE barras IN (%s)" % ','.join(['%s'] * len(df))
#         with self.engine.connect() as conn:
#             conn.execute(consulta_delete, tuple(df['barras']))
#
#         # Salvar o dataframe na tabela do banco de dados
#         df.to_sql(tabela, self.engine, if_exists='append', index=False, chunksize=1000)
#
#         # Prepara a instrução SQL de inserção
#         # num_colunas = len(df.columns)
#         # placeholders = ','.join(['?'] * num_colunas)
#         # strSQL = "INSERT INTO [{0}] VALUES ({1})".format(tabela, placeholders)
#         #
#         # # Executa a inserção em bloco
#         # registros = [tuple(map(str, linha)) for linha in df.values]
#         # self.cursor.executemany(strSQL, registros)
#         # self.conxn.commit()
#
#         # except pyodbc.Error as e:
#         #     self.erro = str(e)
#         #
#         # finally:
#         #     self.conxn.close()
#
#         # for linha in df.values:
#         #     my_list = [str(x) for x in linha]
#         #     if len(my_list) > 0:
#         #         if indicelimpeza != -1:
#         #             self.abrirconexao()
#         #             strSQL = "DELETE * FROM [%s] WHERE Barras = %s" % (tabela, my_list[indicelimpeza])
#         #             self.cursor.execute(strSQL)
#         #             self.conxn.commit()
#         #
#         #         strSQL = "INSERT INTO [" + tabela + "] VALUES (%s)" % ', '.join(my_list)
#         #         self.cursor.execute(strSQL)
#         #         self.conxn.commit()
#
#         # df.to_csv('df.csv', sep=';', encoding='utf-8', index=False)
#
#         # RUN QUERY
#         # strSQL = "INSERT INTO [%s] SELECT * FROM [text;HDR=Yes;FMT=Delimited(;);Database=D:\Projetos\Extrair Imposto].[df.csv]" % tabela
#
#         # self.cursor.execute(strSQL)
#         # self.conxn.commit()
#
#         # self.conxn.close()  # CLOSE CONNECTION
#         # os.remove('df.csv')
#
#     def executarsql(self, sql):
#         try:
#             self.abrirconexao()
#             self.cursor.execute(sql)
#             rows_affected = self.cursor.rowcount
#             self.conxn.commit()  # Certifique-se de fazer o commit após a execução
#             return rows_affected
#
#         finally:
#             self.conxn.close()
#             self.conxn = None
#
#     def fecharbanco(self):
#         """
#         Fecha a conexão com o banco de dados
#         """
#         self.cursor.close()


def caminhoprojeto(subpasta=''):
    """
    : param subpasta: adiciona o caminho da subpasta dada como entrada (caso preenchido).
    : return: o caminho do projeto ao qual a rotina está inserida.
    """
    import errno

    try:
        caminho = ''
        if getattr(sys, 'frozen', False):
            caminho = os.path.dirname(sys.executable)
        else:
            caminho = os.path.dirname(os.path.abspath(__file__))

        if len(caminho) > 0:
            if len(subpasta) > 0:
                if not os.path.isdir(os.path.join(caminho, subpasta)):
                    os.mkdir(os.path.join(caminho, subpasta))
                caminho = os.path.join(caminho, subpasta)

        if os.path.isdir(caminho):
            return caminho
        else:
            return ''

    except OSError as exc:
        if exc.errno != errno.EEXIST:
            raise
        pass


def ultimoarquivo(caminho, extensao):
    """
    : param caminho: diretório onde pesquisa será realizada.
    : param extensao: extensão do arquivo que está sendo buscado o último alterado.
    : return: retorna o caminho completo do último arquivo atualizado da extensão e caminho dado como entrada.
    """
    lista_arquivos = glob.glob(caminho + '/*.' + extensao)
    ultimoatualizado = max(lista_arquivos, key=os.path.getmtime)
    if len(ultimoatualizado) > 0:
        ultimoatualizado = os.path.abspath(ultimoatualizado)

    return ultimoatualizado


def renomeararquivo(nomeantigo, novonome, codcliente=''):
    """
    : param nomeantigo: nome do arquivo a ser renomeado (endereço completo).
    : param novonome: nome novo do arquivo (endereço completo).
    """
    if codcliente == '':
        if os.path.isfile(to_raw(novonome)):
            os.remove(to_raw(novonome))
        time.sleep(0.5)
        mover_arquivo(nomeantigo, novonome)
        # os.rename(to_raw(nomeantigo), to_raw(novonome))
    else:
        adicionarcabecalhopdf(nomeantigo, novonome, codcliente)


def to_raw(string):
    """
    : param "string": texto a ser tratado.
    : return: "string" com o prefixo f para criar "strings" literais formatadas e r usado para tornar a "string" numa
    "string" literal bruta e ignorar caracteres especiais 'dentro' dela como o '\', por exemplo.
    """
    return fr"{string}"


def quantidade_cores(caminho):
    """

    : param caminho: endereço da imagem salva.
    : return: a quantidade de cores presente na matriz de cores da imagem.
    """
    from PIL import Image
    # from matplotlib.pyplot import imshow

    if os.path.isfile(caminho):
        image = Image.open(caminho)
        # Corta a imagem para tirar as bordas inúteis
        image = image.crop((1, 0, 130, 24))
        # imshow(image)
        # Pega as cores e a quantidade de pixels (pelo que entendi) da mesma cor e coloca em uma lista
        # Calcula o número de pixels na tela (quant pixels eixo "x" * quant pixels eixo "y") para retornar
        # o máximo de cores possíveis (pior caso cada pixel de uma cor diferente), se tiver mais cores do que o definido
        # ao chamar o getcolors retorna None
        cores = image.convert('RGB').getcolors(image.size[0] * image.size[1])
        if cores is not None:
            # Retorna a quantidade de elementos da lista de cores
            return len(cores)
        else:
            return 0


def retornorvalorlista(lista, indice):
    """

    : param lista: lista a ser realizada a busca.
    : param indice: item a ser retornado.
    : return: item do índice dado como entrada.
    """
    import numpy as np

    lst = np.array(lista)
    return lst[indice]


def escreverlistaexcelog(caminho, lista):
    """

    : param caminho: caminho do arquivo a ser escrito.
    : param lista: lista a ser adicionada no arquivo do caminho dado.
    """

    df = pd.DataFrame(lista)

    writer = pd.ExcelWriter(caminho, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Lista', index=False)
    worksheet = writer.sheets['Lista']
    worksheet.autofit()
    writer._save()


def acertardataatual():
    """

    :return: retorna data e hora do sistema formatado.
    """
    from datetime import datetime

    textodata = datetime.now()
    return str(textodata.strftime('%Y_%m_%d_%H_%M_%S'))


def caminhospadroes(caminho):
    """

    : param caminho: código dos caminhos padrões (dúvidas, confira lista abaixo).
    : return: retorna o caminho padrão selecionado (str)
    """
    import ctypes.wintypes

    # CSIDL	                        Decimal	Hex	    Shell	Description
    # CSIDL_ADMINTOOLS	            48	    0x30	5.0	    The file system directory that is used to store administrative tools for an individual user.
    # CSIDL_ALTSTARTUP	            29	    0x1D	 	    The file system directory that corresponds to the user's nonlocalized Startup program group.
    # CSIDL_APPDATA	                26	    0x1A	4.71	The file system directory that serves as a common repository for application-specific data.
    # CSIDL_BITBUCKET	            10	    0x0A	 	    The virtual folder containing the objects in the user's Recycle Bin.
    # CSIDL_CDBURN_AREA	            59	    0x3B	6.0	    The file system directory acting as a staging area for files waiting to be written to CD.
    # CSIDL_COMMON_ADMINTOOLS	    47	    0x2F	5.0	    The file system directory containing administrative tools for all users of the computer.
    # CSIDL_COMMON_ALTSTARTUP	    30	    0x1E	        NT-based only	The file system directory that corresponds to the nonlocalized Startup program group for all users.
    # CSIDL_COMMON_APPDATA	        35	    0x23	5.0	    The file system directory containing application data for all users.
    # CSIDL_COMMON_DESKTOPDIRECTORY	25	    0x19	        NT-based only	The file system directory that contains files and folders that appear on the desktop for all users.
    # CSIDL_COMMON_DOCUMENTS	    46	    0x2E	 	    The file system directory that contains documents that are common to all users.
    # CSIDL_COMMON_FAVORITES	    31	    0x1F	        NT-based only	The file system directory that serves as a common repository for favorite items common to all users.
    # CSIDL_COMMON_MUSIC	        53	    0x35	6.0	    The file system directory that serves as a repository for music files common to all users.
    # CSIDL_COMMON_PICTURES	        54	    0x36	6.0	    The file system directory that serves as a repository for image files common to all users.
    # CSIDL_COMMON_PROGRAMS	        23	    0x17	        NT-based only	The file system directory that contains the directories for the common program groups that appear on the Start menu for all users.
    # CSIDL_COMMON_STARTMENU	    22	    0x16	        NT-based only	The file system directory that contains the programs and folders that appear on the Start menu for all users.
    # CSIDL_COMMON_STARTUP	        24	    0x18	        NT-based only	The file system directory that contains the programs that appear in the Startup folder for all users.
    # CSIDL_COMMON_TEMPLATES	    45	    0x2D	        NT-based only	The file system directory that contains the templates that are available to all users.
    # CSIDL_COMMON_VIDEO	        55	    0x37	6.0	    The file system directory that serves as a repository for video files common to all users.
    # CSIDL_COMPUTERSNEARME	        61	    0x3D	6.0	    The folder representing other machines in your workgroup.
    # CSIDL_CONNECTIONS	            49	    0x31	6.0	    The virtual folder representing Network Connections, containing network and dial-up connections.
    # CSIDL_CONTROLS	            3	    0x03	 	    The virtual folder containing icons for the Control Panel applications.
    # CSIDL_COOKIES	                33	    0x21	 	    The file system directory that serves as a common repository for Internet cookies.
    # CSIDL_DESKTOP	                0	    0x00	 	    The virtual folder representing the Windows desktop, the root of the shell namespace.
    # CSIDL_DESKTOPDIRECTORY	    16	    0x10	 	    The file system directory used to physically store file objects on the desktop.
    # CSIDL_DRIVES	                17	    0x11	 	    The virtual folder representing My Computer, containing everything on the local computer: storage devices, printers, and Control Panel. The folder may also contain mapped network drives.
    # CSIDL_FAVORITES	            6	    0x06	 	    The file system directory that serves as a common repository for the user's favorite items.
    # CSIDL_FONTS	                20	    0x14	 	    A virtual folder containing fonts.
    # CSIDL_HISTORY	                34	    0x22	 	    The file system directory that serves as a common repository for Internet history items.
    # CSIDL_INTERNET	            1	    0x01	 	    A viritual folder for Internet Explorer.
    # CSIDL_INTERNET_CACHE	        32	    0x20	4.72	The file system directory that serves as a common repository for temporary Internet files.
    # CSIDL_LOCAL_APPDATA	        28	    0x1C	5.0	    The file system directory that serves as a data repository for local (nonroaming) applications.
    # CSIDL_MYDOCUMENTS	            5	    0x05	6.0	    The virtual folder representing the "My Documents" desktop item.
    # CSIDL_MYMUSIC	                13	    0x0D	6.0	    The file system directory that serves as a common repository for music files.
    # CSIDL_MYPICTURES	            39	    0x27	5.0	    The file system directory that serves as a common repository for image files.
    # CSIDL_MYVIDEO	                14	    0x0E	6.0	    The file system directory that serves as a common repository for video files.
    # CSIDL_NETHOOD	                19	    0x13	 	    A file system directory containing the link objects that may exist in the "My Network Places" virtual folder.
    # CSIDL_NETWORK	                18	    0x12	 	    A virtual folder representing Network Neighborhood, the root of the network namespace hierarchy.
    # CSIDL_PERSONAL	            5	    0x05	 	    The file system directory used to physically store a user's common repository of documents. (From shell version 6.0 onwards, CSIDL_PERSONAL is equivalent to CSIDL_MYDOCUMENTS, which is a virtual folder.)
    # CSIDL_PHOTOALBUMS	            69	    0x45	Vista	The virtual folder used to store photo albums.
    # CSIDL_PLAYLISTS	            63	    0x3F	Vista	The virtual folder used to store play albums.
    # CSIDL_PRINTERS	            4	    0x04	 	    The virtual folder containing installed printers.
    # CSIDL_PRINTHOOD	            27	    0x1B	 	    The file system directory that contains the link objects that can exist in the Printers virtual folder.
    # CSIDL_PROFILE	                40	    0x28	5.0	    The user's profile folder.
    # CSIDL_PROGRAM_FILES	        38	    0x26	5.0	    The Program Files folder.
    # CSIDL_PROGRAM_FILESX86	    42	    0x2A	5.0	    The Program Files folder for 32-bit programs on 64-bit systems.
    # CSIDL_PROGRAM_FILES_COMMON	43	    0x2B	5.0	    A folder for components that are shared across applications.
    # CSIDL_PROGRAM_FILES_COMMONX86	44	    0x2C	5.0	    A folder for 32-bit components that are shared across applications on 64-bit systems.
    # CSIDL_PROGRAMS	            2	    0x02	 	    The file system directory that contains the user's program groups (which are themselves file system directories).
    # CSIDL_RECENT	                8	    0x08	 	    The file system directory that contains shortcuts to the user's most recently used documents.
    # CSIDL_RESOURCES	            56	    0x38	6.0	    The file system directory that contains resource data.
    # CSIDL_RESOURCES_LOCALIZED	    57	    0x39	6.0	    The file system directory that contains localized resource data.
    # CSIDL_SAMPLE_MUSIC	        64	    0x40	Vista	The file system directory that contains sample music.
    # CSIDL_SAMPLE_PLAYLISTS	    65	    0x41	Vista	The file system directory that contains sample playlists.
    # CSIDL_SAMPLE_PICTURES	        66	    0x42	Vista	The file system directory that contains sample pictures.
    # CSIDL_SAMPLE_VIDEOS	        67	    0x43	Vista	The file system directory that contains sample videos.
    # CSIDL_SENDTO	                9	    0x09	 	    The file system directory that contains Send To menu items.
    # CSIDL_STARTMENU	            11	    0x0B	 	    The file system directory containing Start menu items.
    # CSIDL_STARTUP	                7	    0x07	 	    The file system directory that corresponds to the user's Startup program group.
    # CSIDL_SYSTEM	                37	    0x25	5.0	    The Windows System folder.
    # CSIDL_SYSTEMX86	            41	    0x29	5.0	    The Windows 32-bit System folder on 64-bit systems.
    # CSIDL_TEMPLATES	            21	    0x15	 	    The file system directory that serves as a common repository for document templates.
    # CSIDL_WINDOWS	                36	    0x24	5.0	    The Windows directory or SYSROOT.

    if caminho != 80:
        csidl_personal = caminho  # Caminho padrão
        shgfp_type_current = 0  # Para não pegar a pasta padrão e sim a definida como documentos

        buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
        ctypes.windll.Shell32.SHGetFolderPathW(None, csidl_personal, None, shgfp_type_current, buf)

        return buf.value
    else:
        if os.name == 'nt':
            import winreg
            sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
            downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
                location = winreg.QueryValueEx(key, downloads_guid)[0]
            return location
        else:
            return os.path.join(os.path.expanduser('~'), 'downloads')


def caminhoselecionado(tipojanela=1, titulojanela='Selecione o caminho/arquivo:',
                       tipoarquivos=('Todos os Arquivos', '*.*'), caminhoini=caminhospadroes(5), arquivoinicial=''):
    """

    : param tipojanela: 1 — Seleciona Arquivo (Padrão); 2 — Seleciona caminho para salvar arquivo; 3 — Seleciona diretório.
    : param titulojanela: cabeçalho exibido na caixa de diálogo.
    : param tipoarquivos: extensão dos arquivos permitidos da seleção.
    : param caminhoini: caminho inicial.
    : param arquivoinicial: arquivo inicial.
    : return:
    """
    import tkinter as tk
    from tkinter import filedialog

    'Cria a janela raiz'
    root = tk.Tk()
    root.withdraw()

    if tipojanela == 1:
        retorno = filedialog.askopenfilename(title=titulojanela,
                                             initialdir=caminhoini, filetypes=tipoarquivos, initialfile=arquivoinicial)
        if retorno is None:  # asksaveasfile return `None` if dialog closed with "cancel".
            return None

    elif tipojanela == 2:
        name = filedialog.asksaveasfile(mode='w', defaultextension='.txt',
                                        filetypes=tipoarquivos,
                                        initialdir=caminhoini,
                                        title=titulojanela, initialfile=arquivoinicial)
        if name is None:  # asksaveasfile return `None` if dialog closed with "cancel".
            return None
        text2save = str(name.name)
        name.write('')
        retorno = text2save

    elif tipojanela == 3:
        name = filedialog.askdirectory(initialdir=caminhoini, title=titulojanela)
        if name is None:  # askdirectory return `None` if dialog closed with "cancel".
            return None
        text2save = name
        retorno = text2save

    else:
        return

    return retorno


def criarinputbox(titulo, mensagem, substituircaracter='', valorinicial=''):
    """
    : param valorinicial: valor pré-preenchido na caixa de texto.
    : param titulo: cabeçalho da caixa de recebimento de dados do usuário (inputbox).
    : param mensagem: mensagem (normalmente descritiva a entrada) para orientar o usuário.
    : param substituircaracter: caso seja um campo de senha informar o parâmetro para que a digitação não fique visível.
    : return: janela com as opções escolhidas.
    """
    import tkinter as tk
    from tkinter import simpledialog

    root = tk.Tk()
    root.attributes("-topmost", True)
    root.withdraw()

    # the input dialog
    user_inp = simpledialog.askstring(title=titulo, prompt=mensagem, initialvalue=valorinicial, show=substituircaracter)
    if user_inp is None:
        user_inp = 0

    root.attributes("-topmost", True)

    return user_inp


def left(s, amount):
    """

    : param s: "string".
    : param amount: quantidade de caracteres.
    :return: retorna à esquerda de uma "string"
    """
    return s[:amount]


def right(s, amount):
    """

    : param s: "string".
    : param amount: quantidade de caracteres.
    : return: retorna à direita de uma "string"
    """
    return s[-amount:]


def mid(s, offset, amount):
    """

    : param s: "string".
    : param "offset": início do "corte" da "string".
    : param amount: quantidade de caracteres.
    : return:
    """
    return s[(offset - 1):(offset - 1) + amount]


def reset_eof_of_pdf_return_stream(pdf_stream_in: list):
    actual_line = 0
    # find the line position of the EOF
    for i, x in enumerate(pdf_stream_in[::-1]):
        if b'%%EOF' in x:
            actual_line = len(pdf_stream_in) - i
            # print(f'EOF found at line position {-i} = actual {actual_line}, with value {x}')
            break

    # return the list up to that point
    return pdf_stream_in[:actual_line]


def extrairtextopdf(caminho, tipos):
    import PyPDF2
    import re

    texto = ''
    df = None

    with open(caminho, 'rb') as arquivo:
        reader = PyPDF2.PdfReader(arquivo)
        for pagina in range(len(reader.pages)):
            p = reader.pages[pagina]
            texto += p.extract_text()

    match tipos.upper():
        case 'BOLETOS':
            lista = re.findall(r'([\d]{1}.[\d]{3}.[\d]{3}-[\d]{1})\n[\d]{2}', texto)
            listalimpa = []
            for indice, linha in enumerate(lista):
                if indice % 2:
                    linha = linha.replace('.', '')
                    linha = linha.replace('-', '')
                    linha = linha.replace('\n', '')
                    listalimpa.append("'" + linha + "'")
            df = pd.DataFrame(listalimpa, columns=['Inscricao'])

            lista = re.findall(r'COMPETÊNCIA\n([\d]{4})\n[\d]{2}', texto)
            listalimpa = []
            for indice, linha in enumerate(lista):
                if indice % 2:
                    listalimpa.append("'" + linha.replace('\n', '') + "'")
            df['Competencia'] = listalimpa

            lista = re.findall(r'([\d]{2}/[\d]{2}/[\d]{4})\n[\d]{2}', texto)
            listalimpa = []
            for indice, linha in enumerate(lista):
                if indice % 2:
                    listalimpa.append("'" + linha.replace('\n', '') + "'")
            df['Vencimentos'] = listalimpa

            lista = re.findall(r'CONTRIBUINTE\n(.*?)\d{2}\.', texto, re.DOTALL)
            listalimpa = []
            for indice, linha in enumerate(lista):
                if indice % 2:
                    listalimpa.append("'" + linha.replace('\n', '') + "'")
            df['Contribuinte'] = listalimpa

            # lista = re.findall(r'AUTENTICAÇÃO AUTOMÁTICA[\D]PARA USO DO BANCO[\D]\n([\d]{11}.[\d] [\d]{11}.[\d] [\d]{11}.[\d] [\d]{11}.[\d])[\D]', texto)
            lista = re.findall(r'AUTENTICAÇÃO AUTOMÁTICA[\D]PARA USO DO BANCO[\D]\n([\d]{11}.[\d] [\d]{11}.[\d] [\d]{11}.[\d] [\d]{11}.[\d])\n', texto)
            listalimpa = []
            for indice, linha in enumerate(lista):
                linha = linha.replace('.', '')
                linha = linha.replace(' ', '')
                listalimpa.append("'" + linha + "'")
            df['Codigo de Barras'] = listalimpa

            lista = re.findall(r'GUIA/COTA\n([\d]{2}/[\d]{2})', texto)
            listalimpa = []
            for indice, linha in enumerate(lista):
                if indice % 2:
                    listalimpa.append("'" + linha.replace('\n', '') + "'")
            df['Guia'] = listalimpa

            lista = re.findall(r'VALOR TOTAL\n(\d+(?:\.\d{3})*,\d{2})', texto)
            listalimpa = []
            for indice, linha in enumerate(lista):
                if indice % 2:
                    listalimpa.append("'" + linha + "'")
            df['Valor Total'] = listalimpa

            lista = re.findall(r'GUIA/COTA\n[\d]{2}/([\d]{2})', texto)
            listalimpa = []
            for indice, linha in enumerate(lista):
                if indice % 2:
                    listalimpa.append("'" + str(int(linha)) + "'")
            df['Parcela'] = listalimpa

        case 'NADA CONSTA':
            lista = re.findall(r'INSCRIÇÃO: ([\d]{1}.[\d]{3}.[\d]{3}-[\d]{1})', texto)
            listalimpa = [f"'{item.strip()}'" for item in lista]
            df = pd.DataFrame(listalimpa, columns=['Inscricao'])

            lista = re.findall(r'COTA.*?- IPTU ([\d]{4})', texto)

            listalimpa = [f"'{item}'" for item in lista]
            df['Competencia'] = listalimpa

            lista = re.findall(r'VENCIMENTO: ([\d]{2}/[\d]{2}/[\d]{4})', texto)
            listalimpa = [f"'{item.strip()}'" for item in lista]
            df['Vencimentos'] = listalimpa

            lista = re.search(r'[\d].[\d]{3}.[\d]{3}-[\d] [\d]{2}([\d\D]*)\n', texto)
            df['Contribuinte'] = "'" + lista.group(1) + "'"

            lista = re.findall(r'([\d]{11}.[\d] [\d]{11}.[\d] [\d]{11}.[\d] [\d]{11}.[\d])', texto)
            if not lista:
                return None

            listalimpa = []
            for indice, linha in enumerate(lista):
                linha = linha.replace('.', '')
                linha = linha.replace(' ', '')
                listalimpa.append("'" + linha + "'")
            df['Codigo de Barras'] = listalimpa

            lista = re.findall(r'COTA ([\d]{2}|ÚNICA) - IPTU ', texto)
            listalimpa = [f"'{item.strip()}'" for item in lista]
            df['Guia'] = listalimpa

            lista = re.findall(r'VALOR.*R\$: (\d+(?:\.\d+)?,\d+)', texto)

            listalimpa = [f"'{item}'" for item in lista]
            df['Valor Total'] = listalimpa

    return df


def hora(timezone, pedaco=''):
    import datetime

    url = 'http://worldtimeapi.org/api/timezone/' + timezone
    resposta = requests.get(url)
    try:
        match pedaco.upper():
            case 'DATA':
                return datetime.datetime.fromisoformat(resposta.json()['datetime']).date()

            case 'HORA':
                return datetime.datetime.fromisoformat(resposta.json()['datetime']).time()

            case _:
                return datetime.datetime.fromisoformat(resposta.json()['datetime'])

    except Exception as e:
        print(str(e))
        match pedaco.upper():
            case 'DATA':
                return datetime.datetime.fromisoformat(resposta.json()['datetime']).date()

            case 'HORA':
                return datetime.datetime.fromisoformat(resposta.json()['datetime']).time()

            case _:
                return datetime.datetime.fromisoformat(resposta.json()['datetime'])


def timezones_disponiveis():
    url = 'http://worldtimeapi.org/api/timezone/'
    resposta = requests.get(url)
    timezones = resposta.json()
    return timezones


def adicionarcabecalhopdf(arquivo, arquivodestino, cabecalho, centralizado=False, codigobarras=True, posicao_x=0, posicao_y=0):
    import fitz

    tempoespera = 0
    # Abre o arquivo PDF que quer adicionar o cabeçalho
    if os.path.isfile(arquivo):
        with fitz.open(arquivo) as pdf:
            # Verifica se está protegido por senha
            if pdf.is_encrypted:
                # Caso esteja protegido por senha chama outra função com outra biblioteca para gerar uma cópia sem senha
                # (não funciona para retirar a senha de leitura), a função retorna se conseguiu gerar a cópia desbloqueada
                # ou não (se consegue ela já descarta o original com senha)
                desbloqueado = removersenhapdf(arquivo)

            # Verifica se o arquivo não era criptografado ou se era e a função "removersenhapdf" retirou a senha
            if not pdf.is_encrypted or desbloqueado:
                # Carrega a fonte em memória
                fonte = fitz.Font(fontfile=os.path.join(caminhoprojeto(), 'arial-bold.ttf'))

                # "Varre" as páginas do PDF
                for pg in pdf:
                    # Verificar se é a primeira página
                    if pg.number == 0:
                        # Linha para ignorar as formatações da página e não interferir na escrita
                        pg.wrap_contents()
                        # Cálculo de meio da página
                        largura_pagina = pg.mediabox_size.y
                        # Verifica o tamanho do texto considerando a fonte informada na variável "fonte" no início da função
                        largura_texto = fonte.text_length(cabecalho, 15)
                        if not centralizado:
                            deslocamento_direita = 30
                        else:
                            deslocamento_direita = -30
                        if posicao_x == 0:
                            posicao_x = (largura_pagina - largura_texto) / 2 + deslocamento_direita
                        if posicao_y == 0:
                            posicao_y = 20
                        # Insere o cabeçalho no meio da página
                        pg.insert_text((posicao_x, posicao_y), cabecalho, fontsize=10, color=(0, 0, 0), fontfile=fonte)
                        # Para o FOR porque só quero o cabeçalho na primeira página
                        break
            # Salvo o arquivo com o cabeçalho
            pdf.save(arquivodestino)

        # while not (os.path.isfile(arquivodestino) or tempoespera <= 30):
        #     time.sleep(1)
        #     tempoespera += 1
        #
        # tempoespera = 0

        # Verifica se o arquivo foi salvo
        if os.path.isfile(arquivodestino):
            # Apaga o arquivo original
            os.remove(arquivo)

            if codigobarras:
                # Pega as informações do boleto (linha digitável, vencimento, valor)
                dadosboleto = boletos.barcodereader(arquivodestino)
                if dadosboleto is not None:
                    return dadosboleto


def removersenhapdf(arquivobloqueado):
    import pikepdf

    desbloqueado = False

    if os.path.isfile(arquivobloqueado):
        caminho, extensao = os.path.splitext(arquivobloqueado)
        # Abre o arquivo e cria uma cópia sem senha
        with pikepdf.open(arquivobloqueado) as pdf:
            pdf.save(caminho + 'Desbloqueado' + extensao)

        # Verifica se o arquivo desbloqueado foi gerado
        if os.path.isfile(caminho + 'Desbloqueado' + extensao):
            with pikepdf.open(caminho + 'Desbloqueado' + extensao) as pdf:
                desbloqueado = not pdf.is_encrypted

            if desbloqueado:
                os.remove(arquivobloqueado)
                os.rename(caminho + 'Desbloqueado' + extensao, arquivobloqueado)
            else:
                os.remove(caminho + 'Desbloqueado' + extensao)

    return desbloqueado


def encontrar_administradora(administradora, campo='Administradora'):
    lista = senha.listamultiplas
    for dicionario in lista:
        if dicionario[campo] == administradora:
            return dicionario
    return None


def salvarcomo(caminhoarquivo, tempoespera):
    import win32con
    import win32api

    possible_titles = ["Salvar arquivo PDF como", "Salvar Saída de Impressão como"]

    # Loop sobre os títulos possíveis
    for title in possible_titles:
        # Encontra a janela de salvar como
        # dlg = win32gui.FindWindow('#32770', title)
        dlg = wait_for_window(title, tempoespera)

        if dlg != 0:
            # Encontra o botão Salvar
            button = win32gui.FindWindowEx(dlg, 0, "Button", "&Salvar")

            # Encontra a caixa de edição do nome do arquivo
            edit = win32gui.FindWindowEx(dlg, 0, "Edit", None)

            # Define o nome do arquivo e o envia para a caixa de edição
            win32gui.SendMessage(edit, win32con.WM_SETTEXT, None, caminhoarquivo)

            # Clica no botão Salvar
            win32gui.SendMessage(dlg, win32con.WM_COMMAND, 1, button)

            # Espera o processo de salvamento terminar
            while wait_for_window(title, tempoespera):
                win32api.Sleep(100)

            return True
        else:
            return False


def wait_for_window(title, max_wait_seconds):
    """Espera até que uma janela com o título especificado seja encontrada ou o tempo máximo de espera seja atingido."""
    start_time = time.time()
    while True:
        dlg = win32gui.FindWindow('#32770', title)
        if dlg != 0:
            return dlg
        if time.time() - start_time >= max_wait_seconds:
            return 0
        time.sleep(0.1)


def mover_arquivo(origem, destino):
    shutil.copyfile(origem, destino)
    os.remove(origem)


def retornastrinflistaadm(campo):
    nomes_administradoras = [admin[campo] for admin in senha.listaadministradora]
    string_nomes_administradoras = ', '.join(["'{}'".format(item) for item in nomes_administradoras])
    return string_nomes_administradoras


def retornarlistaboletos(listaadministradora=None):
    if 'WHERE' not in senha.sqlcondominios:
        if listaadministradora is not None:
            string_nomes_administradoras = ', '.join(["'{}'".format(item) for item in listaadministradora])

            sql = (senha.sqlcondominios + ' WHERE NomeAdm in (%s) ORDER BY Codigo;' % string_nomes_administradoras)
        else:
            sql = (senha.sqlcondominios + ' WHERE NomeAdm in (%s) ORDER BY Codigo;' % retornastrinflistaadm('nomereal'))
    else:
        if listaadministradora is not None:

            string_nomes_administradoras = ', '.join(["'{}'".format(item) for item in listaadministradora])

            sql = (senha.sqlcondominios + ' AND NomeAdm in (%s) ORDER BY Codigo;' % string_nomes_administradoras)
        else:
            # string_nomes_administradoras = "'ABRJ ADMINISTRADORA DE BENS'"
            sql = (senha.sqlcondominios + ' AND NomeAdm in (%s) ORDER BY Codigo;' % retornastrinflistaadm('nomereal'))
            # sql = (senha.sqlcondominios + ' AND NomeAdm in (%s) ORDER BY Codigo;' % string_nomes_administradoras)
    return sql


def enviarSMS(mensagem):
    from twilio.rest import Client

    try:
        client = Client(senha.account_sid, senha.auth_token)
        message = client.messages.create(
            to=senha.my_number,
            from_=senha.twilio_number,
            body=mensagem
        )

        print(message.sid)

        return True

    except Exception as e:
        print(str(e))

        return False


def listartodosarquivoscaminho(caminho, extensao=''):
    arquivos = []

    # Percorre recursivamente todas as pastas e subpastas do diretório especificado
    for root, dirs, files in os.walk(caminho):
        # Para cada arquivo encontrado, verifica se é um arquivo PDF e adiciona à lista
        for file in files:
            if file.endswith(extensao) or len(extensao) == 0:
                arquivos.append(os.path.join(root, file))

    if arquivos:
        return arquivos
    else:
        return None


def buscar_item(lista, item):
    item_norm = os.path.normpath(item)
    for sublista in lista:
        if item_norm in sublista:
            try:
                index = sublista.index(item_norm)
                return sublista[index - 1]
            except IndexError:
                return ''
    return ''


def extrair_arquivos(caminho_origem, temp_dir):
    # extrai todos os arquivos zip do diretório para uma pasta temporária

    import zipfile
    import shutil

    listazips = []
    arquivo_destino = ''
    for root, dirs, files in os.walk(caminho_origem):
        for filename in files:
            if filename.endswith('.zip'):
                caminho = os.path.join(root, filename)
                with zipfile.ZipFile(caminho, 'r') as zip_file:
                    for arquivo_nome in zip_file.namelist():
                        if arquivo_nome.lower().endswith('.pdf'):
                            with zip_file.open(arquivo_nome) as arquivo_zip:
                                caminhoarquivo = os.path.join(temp_dir, os.path.basename(arquivo_nome))
                                with open(caminhoarquivo, 'wb') as arquivo_destino:
                                    shutil.copyfileobj(arquivo_zip, arquivo_destino)
                                if os.path.isfile(caminhoarquivo):
                                    listazips.append([caminho, caminhoarquivo])

    if listazips:
        return listazips
    else:
        return None


def get_files_not_in_list(list_of_files, directory_path):
    directory_path = directory_path.replace('/', '\\')
    if os.path.isdir(directory_path):
        # Obter todos os arquivos pdf na pasta
        all_pdf_files = glob.glob(os.path.join(directory_path, '*.pdf'))

        # Retornar os arquivos que estão na pasta, mas não estão na lista
        return [file for file in all_pdf_files if file not in [item['Nome do Arquivo'] for item in list_of_files]]
    else:
        return []


def prevent_sleep():
    # Impede o sistema de entrar em suspensão ou desligar a tela
    ctypes.windll.kernel32.SetThreadExecutionState(
        ES_CONTINUOUS |
        ES_SYSTEM_REQUIRED |
        ES_DISPLAY_REQUIRED
    )


def allow_sleep():
    # Permite que o sistema entre em suspensão novamente
    ctypes.windll.kernel32.SetThreadExecutionState(
        ES_CONTINUOUS
    )
