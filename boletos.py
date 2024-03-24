from pyzbar.pyzbar import decode
from pdf2image import convert_from_path
import os
import auxiliares as aux
import datetime


def barcodereader(completepath, qualidade=300):
    # Mensagem de erro
    mensagem_erro = "Impossível ler código de barras"

    # Inicialização de variáveis
    tentativas = 3
    barras = []
    valor = 0
    vencimento = ''
    cabecalho = ['Cliente', 'Código de Barras', 'Tipo Código de Barras', 'Nome do Arquivo', 'Linha Digitável', 'Valor', 'Data Vencimento']

    # Função para calcular a quantidade de pixels (não incluída neste código)
    quantidadepixels = calculate_pixels(completepath, completepath, qualidade)

    if quantidadepixels <= 178956970:
        # Caminho para a biblioteca Poppler (ajuste conforme necessário)
        caminhpoppler = aux.caminhoprojeto() + r'\poppler-23.11.0\library\bin'

        # Conversão do PDF em imagens de páginas
        pages = convert_from_path(completepath, qualidade, poppler_path=caminhpoppler)

        for pagina in pages:
            infocodigobarras = decode(pagina)
            if infocodigobarras:
                infocodigobarras = list(filter(lambda x: x.type == 'I25', infocodigobarras))

                if infocodigobarras:
                    codigobarras = infocodigobarras[0].data.decode('ASCII')
                    linhadigitavel = linha_digitavel(codigobarras)
                    if linhadigitavel:
                        valor, vencimento = extrai_info_boleto(linhadigitavel)
                    dados = [aux.left(os.path.basename(completepath), 4), infocodigobarras[0].data.decode('ASCII'), infocodigobarras[0].type, completepath.replace('/', '\\'),
                             linhadigitavel, valor, vencimento]
                    barras = dict(zip(cabecalho, dados))
                    break

        if barras:
            return barras
        else:
            dados_erro = [aux.left(os.path.basename(completepath), 4), mensagem_erro, mensagem_erro, completepath.replace('/', '\\'),
                          mensagem_erro, mensagem_erro, mensagem_erro]
            return dict(zip(cabecalho, dados_erro))
    else:
        dados_erro = [aux.left(os.path.basename(completepath), 4), mensagem_erro, mensagem_erro, completepath.replace('/', '\\'),
                      mensagem_erro, mensagem_erro, mensagem_erro]
        return dict(zip(cabecalho, dados_erro))


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


def extrai_info_boleto(linha_digitavel):
    # Extrai o valor do boleto a partir da linha digitável
    fator_vencimento = int(linha_digitavel[40:44])
    valor_documento = float(linha_digitavel[46:len(linha_digitavel)]) / 100
    data_vencimento = (datetime.date(1997, 10, 7) + datetime.timedelta(days=fator_vencimento)).strftime("%d/%m/%Y")

    return valor_documento, data_vencimento
