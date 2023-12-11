from pyzbar.pyzbar import decode
from pdf2image import convert_from_path
import os
import auxiliares as aux
import datetime


def barcodereader(pdf_path, pdffile, cabecalho, arquivozip='', qualidade=300):

    # try:
    tentativas = 3
    barras = []
    valor = 0
    vencimento = ''

    completepath = os.path.join(pdf_path, pdffile)
    if len(arquivozip) == 0:
        completepathuser = os.path.join(pdffile)
    else:
        completepathuser = os.path.join(arquivozip, os.path.basename(pdffile))

    quantidadepixels = calculate_pixels(pdffile, pdffile, qualidade)
    if quantidadepixels <= 178956970:
        caminhpoppler = aux.caminhoprojeto() + r'\poppler-23.01.0\library\bin'
        pages = convert_from_path(completepath, qualidade, poppler_path=caminhpoppler)

        for pagina in pages:
            infocodigobarras = decode(pagina)
            if infocodigobarras:
                infocodigobarras = list(filter(lambda x: x.type == 'I25', infocodigobarras))

            if infocodigobarras:
                infocodigobarras = decode(pagina)
                if infocodigobarras:
                    infocodigobarras = list(filter(lambda x: x.type == 'I25', infocodigobarras))

                codigobarras = infocodigobarras[0].data.decode('ASCII')
                linhadigitavel = linha_digitavel(codigobarras)
                if linhadigitavel:
                    valor, vencimento = extrai_info_boleto(linhadigitavel)
                dados = [aux.left(os.path.basename(pdffile), 4), infocodigobarras[0].data.decode('ASCII'), infocodigobarras[0].type, completepathuser.replace('/', '\\'),
                         linha_digitavel(infocodigobarras[0].data.decode('ASCII')), valor, vencimento]
                barras = dict(zip(cabecalho, dados))

        if barras:
            return barras
        elif tentativas > 0:
            nova_qualidade = qualidade * 2  # aumenta a qualidade em 50%
            return barcodereader(pdf_path, pdffile, cabecalho, arquivozip, nova_qualidade)
        else:
            return False
    else:
        return False

    # except Exception as e:
    #     print(e, pdffile)
    #     return False


def calculate_pixels(pdf_path, pdffile, qualidade=300):
    import os
    import math
    import PyPDF2

    try:
        completepath = os.path.join(pdf_path, pdffile)

        # Abre o arquivo PDF e obtém o número de páginas e o tamanho de cada página em pontos
        with open(completepath, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            num_pages = len(pdf_reader.pages)
            page_sizes = [pdf_reader.pages[page_num].mediabox for page_num in range(num_pages)]

        # Calcula o número de pixels para cada página em uma determinada resolução
        total_pixels = 0
        for size in page_sizes:
            width, height = size.width, size.height
            pixels = math.ceil(width * qualidade / 72) * math.ceil(height * qualidade / 72)
            total_pixels += pixels

        return total_pixels

    except Exception as e:
        print(e, pdffile)
        return False


def linha_digitavel(linha):
    def modulo10(num):
        soma = 0
        peso = 2
        for c in reversed(num):
            parcial = int(c) * peso
            if parcial > 9:
                s = str(parcial)
                parcial = int(s[0]) + int(s[1])
            soma += parcial
            if peso == 2:
                peso = 1
            else:
                peso = 2

        resto10 = soma % 10
        if resto10 == 0:
            modulo10 = 0
        else:
            modulo10 = 10 - resto10

        return modulo10

    def monta_campo(campo):
        campo_dv = "%s%s" % (campo, modulo10(campo))
        return "%s.%s" % (campo_dv[0:5], campo_dv[5:])

    return ' '.join([monta_campo(linha[0:4] + linha[19:24]), monta_campo(linha[24:34]), monta_campo(linha[34:44]), linha[4], linha[5:19]])


def listarcodigobarras(visual, caminho, temp_dir, listazips):
    import time

    # lista todos os arquivos pdf no diretório original e na pasta temporária
    pdfs = aux.listartodosarquivoscaminho(caminho, '.pdf')
    if pdfs:
        listatemp = aux.listartodosarquivoscaminho(temp_dir, '.pdf')
        if listatemp:
            pdfs += listatemp

    else:
        listatemp = aux.listartodosarquivoscaminho(temp_dir, '.pdf')
        if listatemp:
            pdfs = listatemp


    dados = []
    cabecalho = ['Cliente', 'Código de Barras', 'Tipo Código de Barras', 'Nome do Arquivo', 'Linha Digitável', 'Valor', 'Data Vencimento']
    if pdfs:
        for indice, boletos in enumerate(pdfs):
            # ===================================== Parte Gráfica =======================================================
            visual.mudartexto('labelcodigocliente', 'Arquivo: ' + os.path.basename(boletos))
            visual.mudartexto('labelquantidade', 'Item ' + str(indice + 1) + ' de ' + str(len(pdfs)) + '...')
            visual.mudartexto('labelstatus', 'Extraindo Código de Barras...')
            # Atualiza a barra de progresso das transações (Views)
            visual.configurarbarra('barraextracao', len(pdfs), indice + 1)
            time.sleep(0.1)
            # ===================================== Parte Gráfica =======================================================
            if listazips:
                arquivozip = aux.buscar_item(listazips, boletos)
            else:
                arquivozip = ''
            teste = barcodereader(caminho, boletos, cabecalho, arquivozip)
            if len(arquivozip) == 0:
                completepathuser = os.path.join(boletos)
            else:
                completepathuser = os.path.join(arquivozip, os.path.basename(boletos))
            if type(teste) is bool:
                dadostemp = [aux.left(os.path.basename(boletos), 4), '', '', completepathuser.replace('/', '\\'), '']
                teste = dict(zip(cabecalho, dadostemp))

            if type(teste) is dict:
                dados.append(teste)
            else:
                print('ERRO: ' + teste)

    if dados:
        return dados
    else:
        return None


def importar_boletos(visual):
    import messagebox
    import tempfile

    salvouarquivo = False
    arquivo_caminho_origem = aux.caminhoselecionado(3, 'Pasta dos Boletos')
    local_destino = aux.caminhoselecionado(3, 'Pasta do Resultado')

    if len(arquivo_caminho_origem) > 0:
        with tempfile.TemporaryDirectory() as temp_dir:
            listazips = aux.extrair_arquivos(arquivo_caminho_origem, temp_dir)
            listaexcel = listarcodigobarras(visual, arquivo_caminho_origem, temp_dir, listazips)

        if listaexcel:
            visual.mudartexto('labelstatus', 'Salvando Arquivo...')
            if len(local_destino) > 0:
                nomearquivo = os.path.join(local_destino, 'Log_' + aux.acertardataatual() + '.xlsx')
            else:
                nomearquivo = 'Log_' + aux.acertardataatual() + '.xlsx'

            if len(nomearquivo) > 0:
                aux.escreverlistaexcelog(nomearquivo, listaexcel)

            if os.path.isfile(nomearquivo):
                salvouarquivo = True

            if salvouarquivo:
                messagebox.msgbox('Arquivo Salvo com Sucesso!', messagebox.MB_OK, 'Arquivo Salvo')
            else:
                messagebox.msgbox('Problema ao salvar o arquivo!', messagebox.MB_OK, 'Erro Salvamento')
        else:
            messagebox.msgbox('Nenhum Código de Barras encontrado no caminho selecionado!', messagebox.MB_OK, 'Erro Caminho')

    else:
        messagebox.msgbox('Pasta Não Selecionada!', messagebox.MB_OK, 'Erro Caminho')

    visual.configurarbarra('barraextracao', 200, 0)
    visual.mudartexto('labelquantidade', '')
    visual.mudartexto('labelstatus', '')
    visual.mudartexto('labelcodigocliente', 'Arquivo:            ')


def extrai_info_boleto(linha_digitavel):
    # Extrai o valor do boleto a partir da linha digitável
    fator_vencimento = int(linha_digitavel[40:44])
    valor_documento = float(linha_digitavel[46:len(linha_digitavel)]) / 100
    data_vencimento = (datetime.date(1997, 10, 7) + datetime.timedelta(days=fator_vencimento))

    return valor_documento, data_vencimento
