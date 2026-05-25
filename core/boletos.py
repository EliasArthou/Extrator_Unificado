from pyzbar.pyzbar import decode, ZBarSymbol
from pdf2image import convert_from_path
import os
from auxiliares import utils as aux
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

        # Workaround: pdfinfo.exe do poppler não consegue abrir paths com
        # caracteres não-ASCII (ex: "Condomínios" com 'í'). Converte pra
        # short path 8.3 (ex: "CONDOM~1") que é sempre ASCII puro.
        completepath_para_poppler = completepath
        try:
            import win32api as _win32api
            completepath_para_poppler = _win32api.GetShortPathName(completepath)
        except Exception:
            pass  # se win32api não disponível, tenta o caminho normal

        # Conversão do PDF em imagens de páginas
        pages = convert_from_path(completepath_para_poppler, qualidade,
                                   poppler_path=caminhpoppler)

        for pagina in pages:
            # Restringe symbols aos relevantes pra boletos BR: I25 (padrao),
            # QRCODE (boletos novos com QR PIX), PDF417 (alguns boletos antigos)
            # e CODE128 (etiquetas comerciais). Evita warnings ruidosos do decoder
            # DataBar (zbar tenta varios decoders em paralelo por default).
            # Bonus: leitura mais rapida que decode() sem restricao.
            infocodigobarras = decode(pagina, symbols=[
                ZBarSymbol.I25,
                ZBarSymbol.QRCODE,
                ZBarSymbol.PDF417,
                ZBarSymbol.CODE128,
            ])
            if infocodigobarras:
                # Aceita I25 (boleto BR padrao) e QRCODE (boleto novo com QR PIX)
                infocodigobarras = list(filter(
                    lambda x: x.type in ('I25', 'QRCODE'), infocodigobarras
                ))
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
    """
    Extrai vencimento e valor de uma linha digitável de 47 dígitos.

    Layout padrão FEBRABAN da linha digitável (sem pontuação):
      pos  0- 9 : campo 1 (banco + moeda + 5 chars do free + DV1)
      pos 10-20 : campo 2 (10 chars do free + DV2)
      pos 21-31 : campo 3 (10 chars do free + DV3)
      pos 32    : DV geral
      pos 33-36 : fator de vencimento (4 chars)
      pos 37-46 : valor (10 chars, sem vírgula — dividido por 100)

    O código anterior usava [40:44] e [46:], posições erradas que retornavam
    fator/valor incorretos.

    Vencimento:
      Em fev/2025 a FEBRABAN resetou o fator: o dia 22/02/2025 passou a ser
      fator 1000 (em vez de 9978 pela base antiga 07/10/1997). Pra decidir
      qual base usar, escolhe a que coloca o vencimento mais próximo de hoje.
    """
    # ─ Limpa formatação (linha_digitavel() retorna com espaços e pontos:
    #   "23792.76609 90000.186644 49003.745301 1 14530000216011") ───────────────
    linha_limpa = linha_digitavel.replace('.', '').replace(' ', '').replace('-', '')

    # ─ Extração ────────────────────────────────────────────────────────────────
    fator_vencimento = int(linha_limpa[33:37])
    valor_documento  = float(linha_limpa[37:47]) / 100

    # ─ Bases de cálculo ────────────────────────────────────────────────────────
    base_antiga = datetime.date(1997, 10, 7)            # fator 0 = 07/10/1997
    base_nova   = datetime.date(2025, 2, 22)            # fator 1000 = 22/02/2025
    hoje        = datetime.date.today()

    data_antiga = base_antiga + datetime.timedelta(days=fator_vencimento)
    data_nova   = base_nova   + datetime.timedelta(days=fator_vencimento - 1000)

    # ─ Escolhe a base que dá vencimento mais próximo de hoje ──────────────────
    diff_antiga = abs((data_antiga - hoje).days)
    diff_nova   = abs((data_nova   - hoje).days)

    data_vencimento = data_nova if diff_nova <= diff_antiga else data_antiga

    return valor_documento, data_vencimento.strftime("%d/%m/%Y")


# Versão antiga preservada pra rollback se a nova quebrar algum boleto:
# def extrai_info_boleto(linha_digitavel):
#     fator_vencimento = int(linha_digitavel[40:44])   # ← posição ERRADA
#     valor_documento  = float(linha_digitavel[46:]) / 100   # ← posição ERRADA
#     quant_dias = 50
#     database_inicial = datetime.date(1997, 10, 7)
#     hoje = datetime.date.today()
#     blocosdiascorridos = (((hoje + datetime.timedelta(days=quant_dias)) - database_inicial).days) // 10000
#     data_base_nova = database_inicial + datetime.timedelta(days=blocosdiascorridos * 10000)
#     data_base_antiga = database_inicial + datetime.timedelta(days=(blocosdiascorridos - 1) * 10000)
#     limite_data = hoje + datetime.timedelta(days=quant_dias)
#     data_vencimento_antiga = data_base_antiga + datetime.timedelta(days=fator_vencimento)
#     data_vencimento_nova = data_base_nova + datetime.timedelta(days=fator_vencimento)
#     diferenca_antiga = abs((data_vencimento_antiga - limite_data).days)
#     diferenca_nova = abs((data_vencimento_nova - limite_data).days)
#     if diferenca_nova > 950:
#         diferenca_nova = diferenca_nova - 1000
#         fator_vencimento -= 1000
#         data_vencimento_nova = data_base_nova + datetime.timedelta(days=fator_vencimento)
#     if diferenca_antiga < diferenca_nova or fator_vencimento > 9970:
#         data_vencimento = data_vencimento_antiga
#     else:
#         data_vencimento = data_vencimento_nova
#     return valor_documento, data_vencimento.strftime("%d/%m/%Y")


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


# ═══════════════════════════════════════════════════════════════════════════════
#  Pós-processamento unificado de boleto recém-baixado
# ═══════════════════════════════════════════════════════════════════════════════

_MESES_PT_BOLETOS = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
                     "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


def _resolver_destino_com_dedup(destino_inicial, barcode_novo):
    """
    Resolve o caminho de destino considerando colisão de nome na subpasta.

    Regras (combinadas com o usuário):
      1. Nome igual + barcode igual → descarta o novo (mantém o existente).
      2. Nome igual + barcode diferente → renomeia sequencial (_1, _2, ...).
      3. Nome igual + barcode do existente ilegível → renomeia sequencial
         (perder boleto é pior que ter duplicata).

    Parâmetros
    ----------
    destino_inicial : pathlib.Path — caminho de destino sem dedup
    barcode_novo    : str          — código de barras do PDF novo (pode ser '')

    Retorno
    -------
    (destino_final : Path, descartar : bool)
        - descartar=True  → o caller deve apagar o PDF novo e usar o existente.
        - descartar=False → o caller move o PDF novo para `destino_final`.
    """
    if not destino_inicial.exists():
        return destino_inicial, False

    # Lê barcode do arquivo já existente na subpasta
    barcode_existente = ""
    try:
        info_exist = barcodereader(str(destino_inicial), renomear=False)
        if info_exist:
            barcode_existente = info_exist.get('Código de Barras') or ''
    except Exception as e:
        logging.warning(
            f"Não foi possível ler barcode de {destino_inicial.name} "
            f"para dedup: {e}"
        )
        barcode_existente = ""

    # Caso 1: barcodes batem → descarta o novo
    if barcode_novo and barcode_existente and barcode_novo == barcode_existente:
        return destino_inicial, True

    # Caso 2/3: diferente ou ilegível → renomeia sequencial
    k = 1
    candidato = destino_inicial.parent / (
        f"{destino_inicial.stem}_{k}{destino_inicial.suffix}"
    )
    while candidato.exists():
        k += 1
        candidato = destino_inicial.parent / (
            f"{destino_inicial.stem}_{k}{destino_inicial.suffix}"
        )
    return candidato, False


def processar_boleto_baixado(caminho_pdf, codigo_cliente, pasta_downloads,
                             pasta_destino_forcada=None, info_barcode=None):
    """
    Pós-processa um PDF de boleto recém-baixado:

      1. Lê código de barras (pyzbar via barcodereader, sem efeito colateral
         de move interno — passando renomear=False).
      2. Move para a subpasta de destino:
         (a) Se `pasta_destino_forcada` for fornecida → move pra essa pasta
             (ignora o vencimento do barcode). Usado pelo extracao_api.py
             para forçar todos os PDFs de um ciclo para a mesma pasta YYYY_MM_<Mes>,
             evitando que barcodes mal lidos enviem o PDF pra pasta errada.
         (b) Senão, calcula YYYY_MM_<Mes> a partir do vencimento extraído
             e move pra subpasta ao lado do arquivo. Sanity check: se vencimento
             estiver mais de 6 meses no futuro, NÃO move (provável bug de barcode).
         O nome do arquivo é PRESERVADO (sem prefixo de data).
      3. Calcula Pasta_Mes para uso em groupby no relatório Excel.
      4. Retorna dict com chaves esperadas pelos call sites em
         extracao_api.py e fluxo_pw_new.py.

    Se o barcode não puder ser lido, ainda assim move pra `pasta_destino_forcada`
    (se fornecida), e retorna dict com campos de barcode vazios.

    Parâmetros
    ----------
    caminho_pdf            : str | Path — caminho atual do PDF (já com cabeçalho)
    codigo_cliente         : str         — identificador (primeiros 4 chars = cliente)
    pasta_downloads        : str | Path — raiz de downloads (usado só como fallback)
    pasta_destino_forcada  : str | Path | None — se fornecido, força o destino do
                                                 move pra esta pasta independente
                                                 do vencimento extraído

    Retorno
    -------
    dict | None  — None só se o arquivo não existe; caso contrário um dict
                   com as chaves Cliente, Cod_Barras, Tipo_Cod_Barras,
                   Nome_Arquivo, Linha_Digitavel, Valor, Vencimento,
                   Pasta_Mes, Caminho_Completo.
    """
    from pathlib import Path as _Path

    caminho = _Path(caminho_pdf)
    if not caminho.exists():
        return None

    mensagem_erro = "Impossível ler código de barras"

    # ── 1. Lê barcode (ou usa info já fornecida pelo chamador) ────────────────
    # Se info_barcode foi passada, usa ela — usado pelo extracao_api.py que
    # lê o barcode ANTES de aplicar o cabeçalho (porque pypdf reescreve o PDF
    # de uma forma que pyzbar não consegue mais decodificar o I25).
    cod_barras = tipo = linha_dig = vencimento = ""
    valor = 0
    if info_barcode is not None:
        info = info_barcode
    else:
        try:
            info = barcodereader(str(caminho), renomear=False)
        except Exception as e:
            logging.error(f"barcodereader falhou em {caminho}: {e}")
            info = {}

    cod_barras = (info.get('Código de Barras') or '') if info else ''
    tipo       = (info.get('Tipo Código de Barras') or '') if info else ''
    linha_dig  = (info.get('Linha Digitável') or '') if info else ''
    valor      = (info.get('Valor', 0) or 0) if info else 0
    vencimento = (info.get('Data Vencimento') or '') if info else ''

    barcode_ok = (cod_barras and cod_barras != mensagem_erro
                  and vencimento and vencimento != mensagem_erro)

    caminho_final = caminho
    pasta_mes_nome = ""

    # ── 2. Decide o destino e move ────────────────────────────────────────────
    if pasta_destino_forcada is not None:
        # CAMINHO A: destino vem de fora (ex: ciclo da CLI). Ignora vencimento.
        try:
            pasta_destino = _Path(pasta_destino_forcada)
            pasta_destino.mkdir(parents=True, exist_ok=True)
            pasta_mes_nome = pasta_destino.name

            destino_inicial = pasta_destino / caminho_final.name
            # Não compara consigo mesmo (caso o PDF ja' esteja no destino)
            if caminho_final.resolve() == destino_inicial.resolve():
                pass  # ja' está no lugar certo
            else:
                destino_final, descartar = _resolver_destino_com_dedup(
                    destino_inicial, cod_barras
                )
                if descartar:
                    logging.info(
                        f"Duplicata descartada: {caminho_final.name} ja' existe "
                        f"em {pasta_destino.name} com mesmo barcode"
                    )
                    try:
                        caminho_final.unlink()
                    except Exception as e:
                        logging.error(f"Erro ao apagar duplicata {caminho_final}: {e}")
                    caminho_final = destino_inicial
                else:
                    shutil.move(str(caminho_final), str(destino_final))
                    caminho_final = destino_final
        except Exception as e:
            logging.error(f"Erro ao mover {caminho_final} para destino forçado: {e}")

    elif barcode_ok:
        # CAMINHO B: usa vencimento do barcode + sanity check.
        try:
            venc_dt = datetime.datetime.strptime(vencimento, "%d/%m/%Y")
            hoje = datetime.date.today()
            meses_no_futuro = (venc_dt.date().year - hoje.year) * 12 + \
                              (venc_dt.date().month - hoje.month)

            if meses_no_futuro > 6:
                logging.warning(
                    f"Vencimento suspeito ({vencimento}) em {caminho.name} — "
                    f"{meses_no_futuro} meses no futuro. PDF mantido em {caminho.parent}."
                )
            else:
                pasta_mes_nome = (
                    f"{venc_dt.year}_{venc_dt.month:02d}_"
                    f"{_MESES_PT_BOLETOS[venc_dt.month - 1]}"
                )
                pasta_destino = caminho_final.parent / pasta_mes_nome
                pasta_destino.mkdir(parents=True, exist_ok=True)
                destino_inicial = pasta_destino / caminho_final.name
                destino_final, descartar = _resolver_destino_com_dedup(
                    destino_inicial, cod_barras
                )
                if descartar:
                    logging.info(
                        f"Duplicata descartada: {caminho_final.name} ja' existe "
                        f"em {pasta_destino.name} com mesmo barcode"
                    )
                    try:
                        caminho_final.unlink()
                    except Exception as e:
                        logging.error(f"Erro ao apagar duplicata {caminho_final}: {e}")
                    caminho_final = destino_inicial
                else:
                    shutil.move(str(caminho_final), str(destino_final))
                    caminho_final = destino_final
        except Exception as e:
            logging.error(f"Erro ao mover {caminho_final}: {e}")

    # ── 3. Cliente: prioriza codigo_cliente, senão deriva do nome ─────────────
    cliente = (codigo_cliente[:4] if codigo_cliente
               else caminho_final.name[:4])

    return {
        "Cliente":          cliente,
        "Cod_Barras":       cod_barras if barcode_ok else "",
        "Tipo_Cod_Barras":  tipo if barcode_ok else "",
        "Nome_Arquivo":     caminho_final.name,
        "Linha_Digitavel":  linha_dig if barcode_ok else "",
        "Valor":            valor if barcode_ok else 0,
        "Vencimento":       vencimento if barcode_ok else "",
        "Pasta_Mes":        pasta_mes_nome,
        "Caminho_Completo": str(caminho_final),
    }