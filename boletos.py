from pyzbar.pyzbar import decode
from pdf2image import convert_from_path
import os
import auxiliares as aux
import datetime
import shutil
import logging
from dotenv import load_dotenv

# Carrega as variáveis do arquivo .env
load_dotenv()


def barcodereader(completepath, qualidade=300, renomear=True):
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
                        if vencimento and renomear:
                            mover_arquivo_condominio(completepath, vencimento)

                    # Obtém o nome do arquivo
                    basename = os.path.basename(completepath)

                    cliente = basename[:4]
                    dados = [cliente, infocodigobarras[0].data.decode('ASCII'), infocodigobarras[0].type, completepath.replace('/', '\\'),
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
    # Extrai o fator de vencimento e o valor do boleto
    fator_vencimento = int(linha_digitavel[40:44])
    valor_documento = float(linha_digitavel[46:]) / 100

    quant_dias = 50
    # Data base inicial (antiga)
    database_inicial = datetime.date(1997, 10, 7)

    # Data atual
    hoje = datetime.date.today()

    # Calcula quantos blocos de 10.000 dias completos se passaram até (hoje + "quant_dias" dias)
    blocosdiascorridos = (((hoje + datetime.timedelta(days=quant_dias)) - database_inicial).days) // 10000

    # Calcula as datas base
    data_base_nova = database_inicial + datetime.timedelta(days=blocosdiascorridos * 10000)
    data_base_antiga = database_inicial + datetime.timedelta(days=(blocosdiascorridos - 1) * 10000)

    # Define a data limite para comparação (hoje + "quant_dias" dias)
    limite_data = hoje + datetime.timedelta(days=quant_dias)

    # Calcula as possíveis datas de vencimento
    data_vencimento_antiga = data_base_antiga + datetime.timedelta(days=fator_vencimento)
    data_vencimento_nova = data_base_nova + datetime.timedelta(days=fator_vencimento)

    # Calcula a diferença (em dias) entre cada vencimento e a data limite
    diferenca_antiga = abs((data_vencimento_antiga - limite_data).days)
    diferenca_nova = abs((data_vencimento_nova - limite_data).days)
    if diferenca_nova > 950:
        diferenca_nova = diferenca_nova - 1000
        fator_vencimento -= 1000
        data_vencimento_nova = data_base_nova + datetime.timedelta(days=fator_vencimento)


    # Seleciona a data de vencimento mais próxima do limite; se empatar, escolhe a nova
    if diferenca_antiga < diferenca_nova or fator_vencimento > 9970:
        data_vencimento = data_vencimento_antiga
    else:
        data_vencimento = data_vencimento_nova

    return valor_documento, data_vencimento.strftime("%d/%m/%Y")


def renomear_arquivo(caminho_atual, vencimento):
    """
    Renomeia o arquivo com o formato AAAAMM_nomeoriginal.pdf, se o nome do arquivo ainda não começar com AAAAMM_.

    Args:
    caminho_atual (str): Caminho atual do arquivo.
    vencimento (str): Data de vencimento no formato DD/MM/AAAA.
    """
    try:
        vencimento_datetime = datetime.datetime.strptime(vencimento, "%d/%m/%Y")
        ano_mes = vencimento_datetime.strftime("%Y%m")
        nome_arquivo = os.path.basename(caminho_atual)

        # Verificar se o nome do arquivo já começa com AAAAMM_
        if not nome_arquivo.startswith(ano_mes + "_"):
            novo_nome = ano_mes + "_" + nome_arquivo
            novo_caminho = os.path.join(os.path.dirname(caminho_atual), novo_nome)
            os.rename(caminho_atual, novo_caminho)
    except Exception as e:
        print(f"Erro ao renomear o arquivo: {str(e)}")


def mover_arquivo_condominio(caminho_atual, vencimento):
    """
    Move o arquivo para uma pasta no formato YYYY_MM se o ano e mês do vencimento forem menores que o ano e mês atuais.
    Se a pasta não existir, ela será criada. Caso já exista um arquivo com o mesmo nome, ele será renomeado.

    Args:
    caminho_atual (str): Caminho atual do arquivo.
    vencimento (str): Data de vencimento no formato DD/MM/AAAA.

    Returns:
    str: Novo caminho do arquivo, caso movido; ou o caminho original, se não houver movimentação.
    """
    try:
        # Verifica se a data está no formato esperado
        try:
            vencimento_datetime = datetime.datetime.strptime(vencimento, "%d/%m/%Y")
        except ValueError:
            raise ValueError("Formato de data inválido. Use DD/MM/AAAA.")

        ano_mes_vencimento = vencimento_datetime.strftime("%Y_%m")
        ano_mes_atual = datetime.datetime.now().strftime("%Y_%m")

        # Verifica se o ano e mês do vencimento são menores que o atual
        if ano_mes_vencimento < ano_mes_atual:
            diretorio_destino = os.path.join(os.path.dirname(caminho_atual), ano_mes_vencimento)

            # Cria a pasta se não existir
            if not os.path.exists(diretorio_destino):
                os.makedirs(diretorio_destino)

            # Define o caminho completo do destino
            nome_arquivo = os.path.basename(caminho_atual)
            caminho_destino = os.path.join(diretorio_destino, nome_arquivo)

            # Verifica se o arquivo já existe e renomeia se necessário
            base, extensao = os.path.splitext(caminho_destino)
            contador = 1
            while os.path.exists(caminho_destino):
                caminho_destino = f"{base}_{contador}{extensao}"
                contador += 1

            # Move o arquivo para a pasta
            shutil.move(caminho_atual, caminho_destino)
            print(f"Arquivo movido para: {caminho_destino}")
            return caminho_destino  # Retorna o novo caminho do arquivo
        else:
            print("O arquivo permanece no local atual.")
            return caminho_atual  # Retorna o caminho original

    except Exception as e:
        logging.error(f"Erro ao mover o arquivo {caminho_atual}: {str(e)}")
        return caminho_atual  # Retorna o caminho original em caso de erro