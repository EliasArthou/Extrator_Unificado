"""
Extrator IPTU Híbrido: Playwright navega + intercepta resposta do boleto.

Resolve o problema de sincronismo do Playwright puro:
- Playwright faz toda a navegação (reCAPTCHA, inscrição, guia, checkbox, gerar)
- Em vez de expect_download + popup, intercepta a resposta HTTP da SegundaTela
- Extrai o PDF em base64 diretamente da resposta interceptada
- Sem expect_download, sem popup de confirmação, sem race conditions
"""

import os
import re
import base64
import requests
import pandas as pd
import auxiliares as aux
from datetime import date, datetime
from bs4 import BeautifulSoup
from anticaptchaofficial.imagecaptcha import imagecaptcha
from anticaptchaofficial.recaptchav2proxyless import recaptchaV2Proxyless
from playwright.sync_api import Page, Response
from dotenv import load_dotenv
from rich.console import Console
import logging
import urllib3

load_dotenv()

# Suprime warnings de conexão do requests/urllib3 (ex: ConnectionResetError ao reusar sessão)
logging.getLogger("urllib3").setLevel(logging.ERROR)
urllib3.disable_warnings()

console = Console(force_terminal=True, color_system="truecolor")

Codigo, NrIPTU = 0, 1
NrCBM = 1    # coluna CBM na query SQLCBM (Codigo=0, CBM=1, IPTU=2, Cidade=3)
NrIPTU_CBM = 2  # coluna IPTU na query SQLCBM (diferente do NrIPTU do IPTU Rio)


def _extrair_pdf_base64(html: str) -> str | None:
    """Extrai o PDF em base64 do HTML da SegundaTela.aspx."""
    match = re.search(r'data:application/pdf;base64,([A-Za-z0-9+/=\s]+)', html)
    if match:
        return match.group(1).replace('\n', '').replace('\r', '').replace(' ', '')
    return None


# ─────────────────────────────────────────────
# PLAYWRIGHT: navegação completa
# ─────────────────────────────────────────────

def _navegar_ate_dados(page: Page, nriptu: str) -> tuple[bool, str]:
    """
    Playwright: preenche inscrição + PROSSEGUIR (reCAPTCHA).
    Retorna: (sucesso, mensagem_erro)
    """
    ERRO_SESSAO = 'Exception: Message TelaSelecao was received'

    for tentativa in range(3):
        try:
            if page.get_by_text("FALHA NA REQUISIÇÃO").is_visible():
                page.get_by_role("link", name="clique aqui").click()
                page.wait_for_load_state("networkidle")

            if os.getenv('SITEIPTU') not in page.url:
                page.goto(os.getenv('SITEIPTU'), wait_until="networkidle", timeout=60000)

            page.wait_for_load_state("networkidle")

            input_insc = page.locator("#ctl00_ePortalContent_inscricao_input")
            input_insc.wait_for(state="visible")
            input_insc.fill(nriptu)

            try:
                with page.expect_navigation(wait_until="networkidle", timeout=30000):
                    page.locator("input[value='PROSSEGUIR']").click()
            except:
                pass

            msg = page.locator("#ctl00_ePortalContent_MSG")
            if msg.is_visible(timeout=2000):
                txt = msg.inner_text().strip()
                if ERRO_SESSAO in txt:
                    page.goto(os.getenv('SITEIPTU'))
                    continue
                return False, txt

            if "TelaSelecao.aspx" in page.url and not page.get_by_text("DADOS DO IMÓVEL").is_visible():
                continue

            page.wait_for_selector("text='DADOS DO IMÓVEL'", state="visible", timeout=15000)
            return True, ""

        except Exception as e:
            if tentativa == 2:
                return False, f"Erro Playwright: {str(e)}"
            try:
                page.goto(os.getenv('SITEIPTU'))
            except:
                pass

    return False, "Falha após 3 tentativas"


def _texto_locator(page: Page, seletor: str) -> str:
    """Retorna o texto de um locator, ou '' se não encontrar."""
    try:
        loc = page.locator(seletor)
        if loc.count() > 0:
            return loc.first.inner_text().strip()
    except:
        pass
    return ''


def _gerar_boleto_interceptado(page: Page, tipo_pag: str, nrguia: int = 0) -> tuple[str | None, str]:
    """
    Playwright: marca checkbox, gera boleto, intercepta resposta HTTP.
    NOTA: a guia já deve ter sido clicada antes de chamar esta função.

    Retorna: (pdf_base64, mensagem_erro)
    """
    try:
        # 1. Marca o checkbox da data
        sel_id = "[id*='cbCotaUnica']" if tipo_pag == '1' else f"[id*='Chk_00{int(tipo_pag) - 1}']"
        page.locator(sel_id).check()

        # 2. Intercepta a resposta do POST de gerar boleto
        resposta_html = None

        def interceptar_resposta(response: Response):
            nonlocal resposta_html
            if "SegundaTela" in response.url and response.status == 200:
                try:
                    resposta_html = response.text()
                except:
                    pass

        page.on("response", interceptar_resposta)

        try:
            # 3. Clica em GERAR BOLETO
            page.locator("#ctl00_ePortalContent_btnDarmIndiv").click()

            # Espera o popup de confirmação e clica "Sim"/"Ok"
            popup_ok = page.locator("#popup_ok")
            try:
                popup_ok.wait_for(state="visible", timeout=3000)
                popup_ok.click()
            except:
                pass

            # Espera a página carregar (SegundaTela com o PDF)
            page.wait_for_load_state("networkidle", timeout=60000)

            # Se não interceptou, tenta pegar do HTML da página atual
            if resposta_html is None:
                resposta_html = page.content()

        finally:
            page.remove_listener("response", interceptar_resposta)

        # 4. Extrai o PDF base64 da resposta interceptada
        if resposta_html:
            pdf_b64 = _extrair_pdf_base64(resposta_html)
            if pdf_b64:
                return pdf_b64, ""
            else:
                # Debug: salva resposta
                debug_path = os.path.join(os.path.dirname(__file__) or ".", "debug_response.html")
                with open(debug_path, "w", encoding="utf-8") as f:
                    f.write(resposta_html)
                return None, f"PDF não encontrado na resposta. Debug salvo em {debug_path}"

        return None, "Nenhuma resposta capturada do servidor"

    except Exception as e:
        return None, f"Erro ao gerar boleto: {str(e)}"


# ─────────────────────────────────────────────
# FUNÇÃO PRINCIPAL (substitui extrairIPTU_pw)
# ─────────────────────────────────────────────

def extrairIPTU_hibrido(page: Page, objeto, linha, nrguia=0):
    """
    Extração híbrida: Playwright navega + intercepta PDF.
    Drop-in replacement para extrairIPTU_pw.

    Retorna: (lista_dados | lista_de_listas, DataFrame | None)
    """
    dadosiptu = []
    df = None

    cod, nriptu = str(linha[Codigo]), str(linha[NrIPTU]).strip()
    gerarboleto = not bool(objeto.visual.somentevalores.get())
    salvardadospdf = objeto.visual.codigosdebarra.get()
    tipo_pag = str(objeto.visual.tipopagamento.get())

    # ── Nome do arquivo com sufixo para múltiplas guias ──
    if nrguia == 0:
        caminhodestino = os.path.join(objeto.pastadownload, f"{cod}_{nriptu}.pdf")
    else:
        caminhodestino = os.path.join(objeto.pastadownload, f"{cod}_{nriptu}_{nrguia + 1}.pdf")

    try:
        # ── Se o arquivo já existe e precisa gerar, pula o download ──
        if os.path.isfile(caminhodestino) and gerarboleto:
            # Arquivo já existe — pula direto pra extração de dados do PDF
            pass
        else:
            # ── Playwright: inscrição + reCAPTCHA ──
            sucesso, erro = _navegar_ate_dados(page, nriptu)
            if not sucesso:
                return [cod, nriptu, '', '', '', '', '', erro], None

            # ── Extrai dados da tela (contribuinte, endereço, exercício) ──
            contribuinte = _texto_locator(page, "#ctl00_ePortalContent_TELA_CONTRIBUINTE")
            endereco = _texto_locator(page, "#ctl00_ePortalContent_TELA_ENDERECO")
            exercicio = _texto_locator(page, "#ctl00_ePortalContent_GuiaExercicio")

            # ── Clica na guia ──
            guias = page.locator('[title*="visualizar / imprimir"]')
            if guias.count() == 0:
                return [cod, nriptu, '', '', '', '', '', 'Verificar (Extrair Manualmente)'], None

            guias.nth(nrguia).click()
            page.wait_for_load_state("networkidle", timeout=15000)

            # ── Extrai valores por cota ──
            match tipo_pag:
                case '1':
                    valor_cota = _texto_locator(page, "#ctl00_ePortalContent_TELA_VALOR_COTA_UNICA")
                    dadosiptu = [cod, nriptu, exercicio, 1, valor_cota, contribuinte, endereco]
                case '2' | '3' | '4':
                    if tipo_pag == '2':
                        name_valor = 'Valor_Prim'
                    elif tipo_pag == '3':
                        name_valor = 'Valor_Segu'
                    else:
                        name_valor = 'Valor_Terc'

                    valores = page.locator(f"[name='{name_valor}']")
                    count = valores.count()
                    for idx in range(count):
                        texto_valor = valores.nth(idx).inner_text().strip()
                        dadosiptu.append([cod, nriptu, exercicio, str(idx + 1), texto_valor,
                                          contribuinte, endereco, 'Ok'])
                case _:
                    dadosiptu = [cod, nriptu, exercicio, '', '', contribuinte, endereco]

            # ── Gera o boleto (se necessário) ──
            if gerarboleto:
                pdf_b64, erro_boleto = _gerar_boleto_interceptado(page, tipo_pag, nrguia)

                if pdf_b64:
                    os.makedirs(os.path.dirname(caminhodestino) or ".", exist_ok=True)
                    with open(caminhodestino, "wb") as f:
                        f.write(base64.b64decode(pdf_b64))

                    try:
                        aux.adicionarcabecalhopdf(caminhodestino, caminhodestino,
                                                  cod, codigobarras=False)
                    except Exception:
                        pass
                else:
                    if isinstance(dadosiptu, list) and len(dadosiptu) > 0:
                        if isinstance(dadosiptu[0], list):
                            # lista de listas (cotas parceladas)
                            for item in dadosiptu:
                                if len(item) > 7:
                                    item[7] = erro_boleto
                        else:
                            dadosiptu.append(erro_boleto)

        # ── Extrai dados do PDF se existir ──
        if (df is None or len(dadosiptu) == 0) and os.path.isfile(caminhodestino) and salvardadospdf:
            df = aux.extrairtextopdf(caminhodestino, 'Boletos')
            if df is not None:
                total_rows = df[df.columns[0]].count()
                listacodigo = ["'" + cod + "'" for _ in range(total_rows)]
                listatipopag = ["'PARCELADO'" for _ in range(total_rows)]
                listaarquivo = ["'" + os.path.basename(caminhodestino) + "'" for _ in range(total_rows)]

                df.insert(loc=0, column='Codigo', value=listacodigo)
                df.insert(loc=4, column='TpoPagto', value=listatipopag)
                df.insert(loc=5, column='Arquivo', value=listaarquivo)

        # ── RESET: volta para TelaSelecao ──
        try:
            page.goto(os.getenv('SITEIPTU'), wait_until="networkidle", timeout=15000)
        except:
            pass

        return dadosiptu, df

    except Exception as e:
        try:
            page.goto(os.getenv('SITEIPTU'))
        except:
            pass
        return [cod, nriptu, '', '', '', '', '', f"Erro: {str(e)}"], None


# ─────────────────────────────────────────────
# NADA CONSTA (híbrido — mesmo portal iportal)
# ─────────────────────────────────────────────

def _navegar_nadaconsta(page: Page, nriptu: str) -> tuple[bool, str, str]:
    """
    Playwright: preenche inscrição + PROSSEGUIR no portal Nada Consta.
    Retorna: (sucesso, mensagem_erro, ano_exercicio)
    """
    url_nc = os.getenv('SITENADACONSTA')

    for tentativa in range(3):
        try:
            if url_nc not in page.url:
                page.goto(url_nc, wait_until="networkidle", timeout=60000)

            page.wait_for_load_state("networkidle")

            # Preenche inscrição
            input_insc = page.locator("#Inscricao")
            input_insc.wait_for(state="visible")
            input_insc.fill('')
            input_insc.fill(nriptu)

            # Pega o exercício selecionado
            exercicio_el = page.locator("#Exercicio")
            anoextracao = ''
            try:
                if exercicio_el.count() > 0:
                    anoextracao = exercicio_el.input_value().strip()
            except:
                pass

            # Clica PROSSEGUIR (tem reCAPTCHA v3)
            try:
                with page.expect_navigation(wait_until="networkidle", timeout=30000):
                    page.locator("#Avancar").click()
            except:
                pass

            # Verifica erro
            msg = page.locator("#ctl00_ePortalContent_MSG")
            try:
                if msg.is_visible(timeout=2000):
                    txt = msg.inner_text().strip()
                    if txt:
                        return False, txt, anoextracao
            except:
                pass

            return True, "", anoextracao

        except Exception as e:
            if tentativa == 2:
                return False, f"Erro Playwright: {str(e)}", ''
            try:
                page.goto(url_nc)
            except:
                pass

    return False, "Falha após 3 tentativas", ''


def extrairnadaconsta_hibrido(page: Page, objeto, linha, dataatual=''):
    """
    Extração híbrida do Nada Consta: Playwright navega + intercepta PDF base64.
    Drop-in replacement para extrairnadaconsta.

    Retorna: (lista_dados, DataFrame | None)
    """
    dadosiptu = []
    df = None

    cod = str(linha[Codigo])
    nriptu = str(linha[NrIPTU]).strip()
    gerarboleto = not bool(objeto.visual.somentevalores.get())

    if dataatual == '':
        dataatual = aux.hora('America/Sao_Paulo', 'DATA')
    anoextracao = str(dataatual.year)

    caminhodestino = os.path.join(objeto.pastadownload, f"{cod}_{nriptu}.pdf")

    try:
        # ── Se o arquivo já existe, pula ──
        if not os.path.isfile(caminhodestino) or not gerarboleto:
            # ── Playwright: inscrição + reCAPTCHA ──
            sucesso, erro, ano_site = _navegar_nadaconsta(page, nriptu)
            if ano_site:
                anoextracao = ano_site

            if not sucesso:
                dadosiptu = [cod, nriptu, anoextracao, erro]
            else:
                # ── Intercepta PDF da resposta ──
                html_conteudo = page.content()
                pdf_b64 = _extrair_pdf_base64(html_conteudo)

                if pdf_b64:
                    # Pode haver múltiplos PDFs (links "aqui")
                    os.makedirs(os.path.dirname(caminhodestino) or ".", exist_ok=True)
                    with open(caminhodestino, "wb") as f:
                        f.write(base64.b64decode(pdf_b64))

                    try:
                        aux.adicionarcabecalhopdf(caminhodestino, caminhodestino,
                                                  cod, codigobarras=False)
                    except Exception:
                        pass
                else:
                    dadosiptu = [cod, nriptu, anoextracao, 'Verificar (Extrair Manualmente)']

        # ── Extrai dados do PDF se existir ──
        if os.path.isfile(caminhodestino):
            df = aux.extrairtextopdf(caminhodestino, 'NADA CONSTA')
            if df is not None:
                df.insert(loc=0, column='Codigo', value="'" + cod + "'")
                df.insert(loc=5, column='Arquivo', value="'" + caminhodestino + "'")
                if len(dadosiptu) == len(linha) + 1:
                    data_atual = date(dataatual.year, dataatual.month, 1)
                    filtros = df[df['Vencimentos'].apply(
                        lambda x: datetime.strptime(str(x).replace("'", ""), "%d/%m/%Y").date() < data_atual
                    )]
                    if filtros is not None and len(filtros) > 0:
                        dadosiptu.append('Imóvel com dívida!')
                    else:
                        dadosiptu.append('Imóvel com parcelas em dia!')
            else:
                dadosiptu = [cod, nriptu, anoextracao, 'Sem valores a pagar!']

        if df is not None:
            if len(linha) == len(dadosiptu):
                dadosiptu.append('')

        # ── RESET ──
        try:
            page.goto(os.getenv('SITENADACONSTA'), wait_until="networkidle", timeout=15000)
        except:
            pass

        return dadosiptu, df

    except Exception as e:
        try:
            page.goto(os.getenv('SITENADACONSTA'))
        except:
            pass
        return [cod, nriptu, anoextracao, f"Erro: {str(e)}"], None


# ─────────────────────────────────────────────
# CERTIDÃO FISCAL (puro requests — site ASP clássico)
# ─────────────────────────────────────────────

def _corrigir_encoding_html(resp: requests.Response) -> str:
    """
    Corrige encoding do HTML retornado pelo site ASP clássico.
    O site usa Windows-1252/Latin-1, mas requests pode interpretar errado.
    Remove charset antigo e força UTF-8.
    """
    # Decodifica os bytes crus com windows-1252 (padrão de sites ASP antigos)
    for enc in ['windows-1252', 'latin-1', 'utf-8']:
        try:
            html = resp.content.decode(enc)
            break
        except (UnicodeDecodeError, LookupError):
            continue
    else:
        html = resp.content.decode('latin-1', errors='replace')

    # Remove qualquer meta charset/content-type existente pra não conflitar
    html = re.sub(
        r'<meta[^>]*charset[^>]*>', '', html, flags=re.IGNORECASE
    )
    html = re.sub(
        r'<meta[^>]*Content-Type[^>]*>', '', html, flags=re.IGNORECASE
    )

    # Insere meta charset UTF-8 no início do head
    if re.search(r'<head', html, re.IGNORECASE):
        html = re.sub(
            r'(<head[^>]*>)', r'\1\n<meta charset="UTF-8">',
            html, count=1, flags=re.IGNORECASE
        )
    else:
        html = '<html><head><meta charset="UTF-8"></head>' + html

    return html


def _html_para_pdf(caminho_html: str, caminho_pdf: str, base_url: str = ''):
    """Converte HTML local pra PDF usando Playwright headless (encoding correto)."""
    from playwright.sync_api import sync_playwright
    from pathlib import Path

    # Insere <base href> pra que imagens com caminhos relativos funcionem
    if base_url:
        with open(caminho_html, 'rb') as f:
            conteudo = f.read()

        # Adiciona base href logo após <head> (ou no início)
        base_tag = f'<base href="{base_url}">'.encode()
        if b'<head>' in conteudo.lower():
            conteudo = re.sub(
                rb'(<head[^>]*>)', rb'\1' + base_tag,
                conteudo, count=1, flags=re.IGNORECASE
            )
        elif b'<html>' in conteudo.lower():
            conteudo = re.sub(
                rb'(<html[^>]*>)', rb'\1<head>' + base_tag + rb'</head>',
                conteudo, count=1, flags=re.IGNORECASE
            )
        else:
            conteudo = b'<html><head>' + base_tag + b'</head>' + conteudo

        with open(caminho_html, 'wb') as f:
            f.write(conteudo)

    url_local = Path(caminho_html).as_uri()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        try:
            page.goto(url_local, wait_until="networkidle", timeout=30000)
            page.pdf(path=caminho_pdf, format="A4", print_background=True)
            print(f"[OK] Certidão salva em: {caminho_pdf}")
        finally:
            browser.close()

    # Remove HTML temporário
    if os.path.isfile(caminho_html):
        os.remove(caminho_html)


def _resolver_captcha_imagem(session: requests.Session, url_captcha: str,
                              pasta_download: str, url_referer: str = '') -> tuple[str | None, imagecaptcha | None]:
    """
    Baixa a imagem do CAPTCHA e resolve via AntiCaptcha.
    Retorna: (texto_resolvido, solver) — o solver é retornado pra poder reportar erro depois.
    """
    caminho_captcha = os.path.join(pasta_download, 'captcha.png')

    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'image/webp,image/apng,image/*,*/*;q=0.8',
        }
        if url_referer:
            headers['Referer'] = url_referer

        resp = session.get(url_captcha, headers=headers, timeout=30)
        if resp.status_code == 200 and len(resp.content) > 100:
            # Converte pra PNG (AntiCaptcha só aceita JPEG/PNG/GIF)
            from PIL import Image
            from io import BytesIO
            img = Image.open(BytesIO(resp.content))
            img.save(caminho_captcha, 'PNG')

            print(f"[CAPTCHA] Imagem baixada: {len(resp.content)} bytes, convertida pra PNG")

            solver = imagecaptcha()
            solver.set_verbose(0)
            solver.set_key(os.getenv('CHAVEANTICAPTCHA'))

            texto = solver.solve_and_return_solution(caminho_captcha)

            # Limpa arquivo temporário
            if os.path.isfile(caminho_captcha):
                os.remove(caminho_captcha)

            if texto and len(str(texto)) > 0:
                return texto, solver
        else:
            print(f"[CAPTCHA] Imagem muito pequena ou erro: status={resp.status_code}, tamanho={len(resp.content)} bytes")
    except Exception as e:
        print(f"[CAPTCHA] Erro ao resolver: {e}")

    return None, None


def extraircertidaonegativa_requests(objeto, linha, dataatual=''):
    """
    Extração da Certidão Fiscal por requests puro (sem browser).
    O site www2.rio.rj.gov.br/smf/iptucertfiscal/ usa CAPTCHA de imagem simples.

    Retorna: (lista_dados, DataFrame | None)
    """
    dadosiptu = []
    df = None
    dfdivida = None

    if dataatual == '':
        dataatual = aux.hora('America/Sao_Paulo', 'DATA')
    anoextracao = str(dataatual.year) if dataatual else ''

    cod = str(linha[Codigo])
    nriptu = str(linha[NrIPTU]).strip()
    caminhodestino = f'certidao_{cod}_{nriptu}.pdf'

    objeto.visual.mudartexto('labelstatus', 'Extraindo certidão...')

    pasta_com_divida = os.path.join(objeto.pastadownload, "Com_Divida")
    pasta_sem_divida = os.path.join(objeto.pastadownload, "Sem_Divida")
    os.makedirs(pasta_com_divida, exist_ok=True)
    os.makedirs(pasta_sem_divida, exist_ok=True)

    # Verifica se já existe
    caminho_arquivo = aux.is_valid_file_in_path(caminhodestino, objeto.pastadownload)

    if not caminho_arquivo:
        if nriptu == '':
            dadosiptu = [cod, nriptu, anoextracao, 'Inscrição inválida!', 'N/A']
        else:
            url_base = os.getenv('SITECERTIDAOENFITEUTICA')
            url_captcha = url_base + 'include/captcha.asp'
            limitetentativas = 30
            resolveu = False

            session = requests.Session()
            session.headers.update({
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
            })

            for tentativa in range(limitetentativas):
                if resolveu:
                    break

                try:
                    # 1. Carrega página inicial pra pegar cookie de sessão
                    resp_inicial = session.get(url_base, timeout=30)
                    if resp_inicial.status_code != 200 or 'Request Rejected' in resp_inicial.text:
                        dadosiptu = [cod, nriptu, anoextracao, 'Problema carregamento site', 'N/A']
                        break

                    # Verifica erro na página (font size=2 color=red)
                    soup_inicial = BeautifulSoup(resp_inicial.text, 'html.parser')
                    erro_pagina = soup_inicial.find(
                        lambda tag: tag.name == 'font' and tag.get('size') == '2' and tag.get('color') == 'red'
                    )
                    if erro_pagina and erro_pagina.text.strip():
                        dadosiptu = [cod, nriptu, anoextracao, erro_pagina.text.strip(), 'N/A']
                        break

                    # 2. Resolve CAPTCHA
                    texto_captcha, solver_captcha = _resolver_captcha_imagem(
                        session, url_captcha, objeto.pastadownload, url_referer=url_base)
                    if not texto_captcha:
                        dadosiptu = [cod, nriptu, anoextracao, 'Problema Captcha', 'N/A']
                        continue

                    # 3. POST com inscrição + captcha
                    dados_post = {
                        'inscricao': nriptu,
                        'texto_imagem': texto_captcha,
                        'btConsultar': 'Consultar'
                    }
                    resp = session.post(url_base, data=dados_post, timeout=60)

                    if resp.status_code != 200:
                        continue

                    # Verifica resposta
                    html_resp = _corrigir_encoding_html(resp)

                    if 'Request Rejected' in html_resp:
                        dadosiptu = [cod, nriptu, anoextracao, 'Problema carregamento site', 'N/A']
                        break

                    # Captcha errado? Reporta ao AntiCaptcha pra estorno
                    if 'Código digitado não confere' in html_resp:
                        if solver_captcha:
                            solver_captcha.report_incorrect_image_captcha()
                        continue

                    # Inscrição inválida?
                    if 'INSCRIÇÃO IMOBILIÁRIA INVÁLIDA' in html_resp.upper():
                        dadosiptu = [cod, nriptu, anoextracao, 'Inscrição inválida!', 'N/A']
                        resolveu = True
                        break

                    # Verifica msg de erro (ctl00_ePortalContent_MSG)
                    soup_resp = BeautifulSoup(html_resp, 'html.parser')
                    msg_erro = soup_resp.find(id='ctl00_ePortalContent_MSG')
                    if msg_erro and msg_erro.text.strip():
                        dadosiptu = [cod, nriptu, anoextracao, msg_erro.text.strip(), 'N/A']
                        resolveu = True
                        break

                    # Sucesso — salva a página como PDF
                    resolveu = True
                    caminho_arquivo = os.path.join(objeto.pastadownload, caminhodestino)

                    # Salva HTML cru e converte pra PDF via Playwright
                    try:
                        caminho_html_temp = os.path.abspath(
                            caminho_arquivo.replace('.pdf', '.html'))
                        with open(caminho_html_temp, 'wb') as f:
                            f.write(resp.content)

                        _html_para_pdf(caminho_html_temp, caminho_arquivo,
                                       base_url=url_base)

                    except Exception as e:
                        dadosiptu = [cod, nriptu, anoextracao, f'Erro ao salvar: {str(e)}', 'N/A']
                        break

                    if os.path.isfile(caminho_arquivo):
                        dfdivida = aux.extrairtextopdf(caminho_arquivo, 'DIVIDAS')
                        if dfdivida is not None:
                            if dfdivida['Possui_Divida']:
                                novonomearquivo = os.path.join(pasta_com_divida, caminhodestino)
                            else:
                                novonomearquivo = os.path.join(pasta_sem_divida, caminhodestino)
                        else:
                            novonomearquivo = os.path.join(objeto.pastadownload, caminhodestino)

                        try:
                            aux.adicionarcabecalhopdf(caminho_arquivo, novonomearquivo,
                                                      aux.left(cod, 4))
                        except Exception:
                            pass
                    else:
                        possui_divida = dfdivida['Possui_Divida'] if dfdivida else 'N/A'
                        dadosiptu = [cod, nriptu, anoextracao, 'Verificar (Extrair Manualmente)',
                                     possui_divida]

                except Exception as e:
                    if tentativa == limitetentativas - 1:
                        dadosiptu = [cod, nriptu, anoextracao, f'Erro: {str(e)}', 'N/A']

            session.close()

    # ── Extrai dados do PDF se existir ──
    caminho_arquivo = caminhodestino if not caminho_arquivo else caminho_arquivo
    if os.path.isfile(caminho_arquivo):
        caminho = caminhodestino if os.path.isfile(caminhodestino) else caminho_arquivo
        df = pd.DataFrame({
            'Codigo': ["'" + cod + "'"],
            'Arquivo': ["'" + caminho + "'"]
        })

    return dadosiptu, df


# ─────────────────────────────────────────────
# USO STANDALONE (para teste rápido)
# ─────────────────────────────────────────────

def testar_extracao(inscricao: str = "31469687", tipo_pagamento: str = "2",
                    pasta_download: str = "Downloads", headless: bool = False):
    """Testa a extração de um boleto individual."""
    from playwright.sync_api import sync_playwright

    caminho_perfil = os.path.join(
        os.path.dirname(__file__), "Profile",
        os.getenv('NOMEPROFILEIPTU', 'iptu_profile')
    )

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=caminho_perfil,
            headless=headless,
            accept_downloads=True,
            args=["--disable-blink-features=AutomationControlled", "--start-maximized"]
        )
        page = context.pages[0] if context.pages else context.new_page()

        try:
            inscricao_limpa = inscricao.replace(".", "").replace("-", "")

            # Playwright: reCAPTCHA + dados
            print(f"[PW] Consultando inscrição {inscricao}...")
            page.goto(os.getenv('SITEIPTU'), wait_until="networkidle", timeout=60000)
            sucesso, erro = _navegar_ate_dados(page, inscricao_limpa)

            if not sucesso:
                print(f"[ERRO] {erro}")
                return

            contribuinte = _texto_locator(page, "#ctl00_ePortalContent_TELA_CONTRIBUINTE")
            endereco = _texto_locator(page, "#ctl00_ePortalContent_TELA_ENDERECO")
            print(f"[PW] Contribuinte: {contribuinte}")
            print(f"[PW] Endereço: {endereco}")

            # Clica na primeira guia antes de gerar boleto
            guias = page.locator('[title*="visualizar / imprimir"]')
            if guias.count() == 0:
                print("[ERRO] Nenhuma guia disponível")
                return
            guias.first.click()
            page.wait_for_load_state("networkidle", timeout=15000)

            # Playwright: checkbox + boleto (interceptando resposta)
            print(f"[PW] Gerando boleto (tipo_pagamento={tipo_pagamento})...")
            pdf_b64, erro_boleto = _gerar_boleto_interceptado(page, tipo_pagamento)

            if pdf_b64:
                os.makedirs(pasta_download, exist_ok=True)
                caminho = os.path.join(pasta_download, f"{inscricao_limpa}.pdf")
                with open(caminho, "wb") as f:
                    f.write(base64.b64decode(pdf_b64))
                tamanho_kb = os.path.getsize(caminho) / 1024
                print(f"[OK] PDF salvo em: {caminho} ({tamanho_kb:.1f} KB)")
            else:
                print(f"[ERRO] {erro_boleto}")

        finally:
            context.close()


def testar_nadaconsta(inscricao: str = "31469687", pasta_download: str = "Downloads",
                      headless: bool = False):
    """Testa a extração do Nada Consta."""
    from playwright.sync_api import sync_playwright

    caminho_perfil = os.path.join(
        os.path.dirname(__file__), "Profile",
        os.getenv('NOMEPROFILEIPTU', 'iptu_profile')
    )

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=caminho_perfil,
            headless=headless,
            accept_downloads=True,
            args=["--disable-blink-features=AutomationControlled", "--start-maximized"]
        )
        page = context.pages[0] if context.pages else context.new_page()

        try:
            inscricao_limpa = inscricao.replace(".", "").replace("-", "")
            url_nc = os.getenv('SITENADACONSTA')

            print(f"[PW] Consultando Nada Consta para inscrição {inscricao}...")
            page.goto(url_nc, wait_until="networkidle", timeout=60000)

            sucesso, erro, anoextracao = _navegar_nadaconsta(page, inscricao_limpa)

            if not sucesso:
                print(f"[ERRO] {erro}")
                return

            print(f"[PW] Exercício: {anoextracao}")

            # Tenta extrair PDF base64 da resposta
            html_conteudo = page.content()
            pdf_b64 = _extrair_pdf_base64(html_conteudo)

            if pdf_b64:
                os.makedirs(pasta_download, exist_ok=True)
                caminho = os.path.join(pasta_download, f"nadaconsta_{inscricao_limpa}.pdf")
                with open(caminho, "wb") as f:
                    f.write(base64.b64decode(pdf_b64))
                tamanho_kb = os.path.getsize(caminho) / 1024
                print(f"[OK] PDF salvo em: {caminho} ({tamanho_kb:.1f} KB)")
            else:
                print("[ERRO] PDF não encontrado na resposta")

        finally:
            context.close()


def testar_certidao(inscricao: str = "31469687", pasta_download: str = "Downloads"):
    """Testa a extração da Certidão Fiscal por requests."""
    inscricao_limpa = inscricao.replace(".", "").replace("-", "")
    url_base = os.getenv('SITECERTIDAOENFITEUTICA')
    url_captcha = url_base + 'include/captcha.asp'

    print(f"[REQ] Consultando Certidão Fiscal para inscrição {inscricao}...")
    session = requests.Session()
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
    })

    try:
        os.makedirs(pasta_download, exist_ok=True)
        max_tentativas = 10

        for tentativa in range(1, max_tentativas + 1):
            # 1. Carrega página pra pegar sessão
            resp = session.get(url_base, timeout=30)
            print(f"[REQ] Tentativa {tentativa}/{max_tentativas} — Página inicial: status {resp.status_code}")

            if 'Request Rejected' in resp.text:
                print("[ERRO] Request Rejected pelo servidor")
                return

            # 2. Resolve CAPTCHA
            texto_captcha, solver_captcha = _resolver_captcha_imagem(
                session, url_captcha, pasta_download, url_referer=url_base)
            if not texto_captcha:
                print("[ERRO] Não conseguiu resolver CAPTCHA, tentando de novo...")
                continue
            print(f"[REQ] CAPTCHA resolvido: {texto_captcha}")

            # 3. POST
            dados_post = {
                'inscricao': inscricao_limpa,
                'texto_imagem': texto_captcha,
                'btConsultar': 'Consultar'
            }
            resp = session.post(url_base, data=dados_post, timeout=60)
            print(f"[REQ] POST resposta: status {resp.status_code}, tamanho {len(resp.text)}")

            if 'Código digitado não confere' in resp.text:
                if solver_captcha:
                    solver_captcha.report_incorrect_image_captcha()
                    print(f"[AVISO] CAPTCHA incorreto — reportado ao AntiCaptcha pra estorno")
                else:
                    print(f"[AVISO] CAPTCHA incorreto")
                continue

            if 'INSCRIÇÃO IMOBILIÁRIA INVÁLIDA' in resp.text.upper():
                print("[ERRO] Inscrição inválida")
                return

            # 4. Sucesso — salva HTML cru e converte pra PDF via Playwright
            caminho_pdf = os.path.join(pasta_download, f"certidao_{inscricao_limpa}.pdf")
            caminho_html = os.path.abspath(caminho_pdf.replace('.pdf', '.html'))

            # Salva bytes crus (o HTML tem seu próprio charset, o browser entende)
            with open(caminho_html, 'wb') as f:
                f.write(resp.content)

            # Usa Playwright headless pra abrir o HTML e imprimir como PDF
            _html_para_pdf(caminho_html, caminho_pdf, base_url=url_base)
            return

        print(f"[ERRO] Falhou após {max_tentativas} tentativas")

    finally:
        session.close()


# ─────────────────────────────────────────────
# BOMBEIROS — Híbrido (Playwright + AntiCaptcha reCAPTCHA v2)
# ─────────────────────────────────────────────

SITEKEY_BOMBEIROS = '6LdZ1EQkAAAAAAKRMr1Dhld-8WiyW5Qt0HgCXyfa'
URL_IFRAME_BOMBEIROS = 'https://www.funesbom.rj.gov.br/sistema/imovel'

# Índices das colunas no resultado da query SQLCBM
# Codigo=0, CBM=1, IPTU=2, Cidade=3
NrCidade = 3   # coluna cidade (nome da cidade)


def _resolver_codigo_municipio(page: Page, cidade: str) -> str:
    """Busca o código do município direto no select da página FUNESBOM.
    Faz match exato, depois case-insensitive, depois parcial.
    Default: '064' (Rio de Janeiro)."""
    if not cidade:
        return '064'
    cidade = cidade.strip().replace('"', '\\"').replace("'", "\\'")

    resultado = page.evaluate(f"""() => {{
        const sel = document.querySelector('select[name="municipio"]');
        if (!sel) return null;
        const nome = "{cidade}";
        const nomeUpper = nome.toUpperCase();
        const opts = Array.from(sel.options).filter(o => o.value);

        // Busca exata
        for (const o of opts) {{
            if (o.text === nome) return o.value;
        }}
        // Case-insensitive
        for (const o of opts) {{
            if (o.text.toUpperCase() === nomeUpper) return o.value;
        }}
        // Parcial (contém)
        for (const o of opts) {{
            if (o.text.toUpperCase().includes(nomeUpper) || nomeUpper.includes(o.text.toUpperCase())) return o.value;
        }}
        return null;
    }}""")

    if resultado:
        return resultado
    console.print(f"[yellow][AVISO][/] Município '{cidade}' não encontrado no combo FUNESBOM, usando Rio de Janeiro (064)")
    return '064'


# ─────────────────────────────────────────────
# CAPTCHA — Resolvedor unificado (imagem + reCAPTCHA v2)
# ─────────────────────────────────────────────

def resolver_captcha(tipo: str = 'auto', *,
                     # -- parâmetros para imagem (tipo='imagem') --
                     session: requests.Session | None = None,
                     url_captcha: str = '',
                     pasta_download: str = '',
                     url_referer: str = '',
                     # -- parâmetros para recaptcha v2 (tipo='recaptchav2') --
                     page_url: str = '',
                     sitekey: str = '',
                     # -- parâmetros para auto-detecção (tipo='auto') --
                     page: Page | None = None,
                     ) -> tuple[str | None, object | None]:
    """
    Resolvedor unificado de CAPTCHA via AntiCaptcha.

    Tipos suportados:
        'imagem'       — CAPTCHA de imagem (baixa e resolve via imagecaptcha)
        'recaptchav2'  — reCAPTCHA v2 checkbox (resolve via recaptchaV2Proxyless)
        'auto'         — Detecta automaticamente a partir da page do Playwright

    Retorna: (solução, solver)
        - imagem: (texto, solver_imagecaptcha) — solver pra report_incorrect se precisar
        - recaptchav2: (token, None) — token g-recaptcha-response
        - None em caso de erro
    """
    if tipo == 'auto':
        tipo = _detectar_tipo_captcha(page)
        if not tipo:
            console.print("[yellow][CAPTCHA][/] Nenhum CAPTCHA detectado na página")
            return None, None
        console.print(f"[cyan][CAPTCHA][/] Tipo detectado: {tipo}")

    if tipo == 'imagem':
        return _resolver_captcha_imagem(session, url_captcha, pasta_download, url_referer)
    elif tipo == 'recaptchav2':
        token = _resolver_recaptcha_v2(page_url, sitekey)
        return token, None
    else:
        console.print(f"[red][CAPTCHA][/] Tipo não suportado: {tipo}")
        return None, None


def _detectar_tipo_captcha(page: Page | None) -> str | None:
    """
    Detecta o tipo de CAPTCHA presente na página Playwright.
    Retorna: 'recaptchav2', 'recaptchav3', 'imagem' ou None.
    """
    if page is None:
        return None

    try:
        info = page.evaluate("""() => {
            // reCAPTCHA v2 — checkbox visível ou iframe anchor
            const hasV2Iframe = !!document.querySelector('iframe[src*="recaptcha"][src*="anchor"]');
            const hasV2Checkbox = !!document.querySelector('.g-recaptcha');
            const hasV2Textarea = !!document.querySelector('textarea[name="g-recaptcha-response"]');

            // reCAPTCHA v3 — invisível, geralmente só script com render=sitekey
            const scripts = Array.from(document.querySelectorAll('script[src*="recaptcha"]'));
            const hasV3Render = scripts.some(s => s.src.includes('render=') && !s.src.includes('render=explicit'));

            // CAPTCHA de imagem — procura img com src contendo captcha
            const hasImageCaptcha = !!document.querySelector('img[src*="captcha"]');

            return { hasV2Iframe, hasV2Checkbox, hasV2Textarea, hasV3Render, hasImageCaptcha };
        }""")

        if info.get('hasV2Iframe') or info.get('hasV2Checkbox') or info.get('hasV2Textarea'):
            return 'recaptchav2'
        if info.get('hasV3Render'):
            return 'recaptchav3'
        if info.get('hasImageCaptcha'):
            return 'imagem'
        return None
    except Exception:
        return None


def _detectar_sitekey_v2(page: Page) -> str | None:
    """Extrai o sitekey do reCAPTCHA v2 da página."""
    try:
        return page.evaluate("""() => {
            // Tenta pegar do atributo data-sitekey
            const div = document.querySelector('[data-sitekey]');
            if (div) return div.dataset.sitekey;

            // Tenta pegar do iframe src
            const iframe = document.querySelector('iframe[src*="recaptcha"]');
            if (iframe) {
                try {
                    const url = new URL(iframe.src);
                    const k = url.searchParams.get('k');
                    if (k) return k;
                } catch(e) {}
            }
            return null;
        }""")
    except Exception:
        return None


def _resolver_recaptcha_v2(page_url: str, sitekey: str) -> str | None:
    """
    Resolve reCAPTCHA v2 usando AntiCaptcha (recaptchaV2Proxyless).
    Mostra progresso com timer inline.
    Retorna o token g-recaptcha-response ou None.
    """
    import time as _time
    import threading

    chave_api = os.getenv('CHAVEANTICAPTCHA')
    if not chave_api:
        console.print("[red][ERRO][/] CHAVEANTICAPTCHA não configurada no .env")
        return None

    if not sitekey:
        console.print("[red][ERRO][/] sitekey não fornecida para reCAPTCHA v2")
        return None

    solver = recaptchaV2Proxyless()
    solver.set_verbose(0)
    solver.set_key(chave_api)
    solver.set_website_url(page_url)
    solver.set_website_key(sitekey)

    # Thread separada pra mostrar timer enquanto resolve
    resultado = [None]
    erro = [None]
    finalizado = threading.Event()

    def _resolver():
        # Supressão do ConnectionResetError feita pelo _FilteredStream no modo lote
        try:
            resultado[0] = solver.solve_and_return_solution()
            if not resultado[0]:
                erro[0] = solver.error_code
        except Exception as e:
            erro[0] = str(e)
        finally:
            finalizado.set()

    t = threading.Thread(target=_resolver, daemon=True)
    inicio = _time.time()
    t.start()

    def _fmt(seg: float) -> str:
        m, s = divmod(int(seg), 60)
        return f"{m:02d}:{s:02d}"

    console.print(f"[cyan][CAPTCHA][/] Resolvendo reCAPTCHA v2...")
    while not finalizado.wait(timeout=5):
        pass  # Mantém a thread ativa pra o Live continuar atualizando

    elapsed = _time.time() - inicio

    if resultado[0]:
        console.print(f"[green][CAPTCHA][/] Resolvido em {_fmt(elapsed)} ({len(resultado[0])} chars)")
        return resultado[0]
    else:
        console.print(f"[red][CAPTCHA][/] Falha em {_fmt(elapsed)}: {erro[0]}")
        return None


def _submeter_busca_bombeiros(page: Page, campos: dict,
                              destino: str = 'imovel/busca',
                              endpoint: str = '/sistema/imovel/busca') -> tuple[bool, str]:
    """
    Submete a busca no site dos Bombeiros via fetch.
    campos: dict com os campos do form (cbmerj/cbmerj_dv ou inscricao/municipio).
    destino: valor do campo hidden 'destino' (varia por aba).
    endpoint: URL do POST (varia por aba).
    Retorna: (sucesso, mensagem_erro)
    """
    # Resolve reCAPTCHA v2 via AntiCaptcha
    sitekey = _detectar_sitekey_v2(page) or SITEKEY_BOMBEIROS
    token, _ = resolver_captcha('recaptchav2', page_url=URL_IFRAME_BOMBEIROS, sitekey=sitekey)
    if not token:
        return False, "Falha ao resolver reCAPTCHA v2"

    # Pega o _token CSRF do form
    csrf_token = page.evaluate("""() => {
        const input = document.querySelector('input[name="_token"]');
        return input ? input.value : '';
    }""")

    # Monta os campos do form
    campos_js = '\n'.join([f"formData.append('{k}', '{str(v).replace(chr(39), chr(92)+chr(39))}');" for k, v in campos.items()])

    resultado = page.evaluate(f"""async () => {{
        const formData = new URLSearchParams();
        formData.append('_token', '{csrf_token}');
        formData.append('destino', '{destino}');
        {campos_js}
        formData.append('g-recaptcha-response', '{token}');

        const resp = await fetch('{endpoint}', {{
            method: 'POST',
            headers: {{'Content-Type': 'application/x-www-form-urlencoded'}},
            body: formData.toString(),
            redirect: 'follow',
        }});

        const html = await resp.text();
        return {{ status: resp.status, url: resp.url, html: html }};
    }}""")

    if resultado:
        status = resultado.get('status', 0)
        html_resp = resultado.get('html', '')
        url_resp = resultado.get('url', '')
        tam_post = len(html_resp)
        tam_post_fmt = f"{tam_post / 1024:.1f} KB" if tam_post >= 1024 else f"{tam_post} B"
        if tam_post < 500:
            console.print(f"[yellow][AVISO][/] POST resposta: status {status}, tamanho: {tam_post_fmt} — resposta muito pequena, possível erro")
        else:
            console.print(f"[cyan][PW][/] POST resposta: status {status}, tamanho: {tam_post_fmt}")

        if 'Dados Cadastrais' in html_resp or 'Vencimentos' in html_resp:
            # NÃO usar page.set_content() — perde a sessão/cookies do site.
            # Guarda o HTML na page via window._htmlBombeiros para uso posterior.
            page.evaluate("(html) => { window._htmlBombeiros = html; }", html_resp)
            return True, ""
        elif 'não encontrado' in html_resp.lower():
            return False, "Imóvel não encontrado"
        else:
            trecho = html_resp[:500].replace('\n', ' ')
            console.print(f"[dim][DEBUG][/] Resposta: {trecho}")
            return False, f"Resposta inesperada (status {status})"
    else:
        return False, "Sem resposta do servidor"


def _navegar_bombeiros(page: Page, nrcbm: str, nriptu: str = '',
                       cidade: str = '') -> tuple[bool, str]:
    """
    Playwright: navega ao iframe de busca dos Bombeiros.
    Ordem de tentativa:
      1. CBMERJ (se tiver)
      2. Inscrição predial / IPTU (se tiver)
    Se não tiver nenhum dos dois, retorna erro sem navegar.
    (Busca por CPF/CNPJ não implementada — requer autenticação gov.br)
    """
    nrcbm_limpo = nrcbm.replace('-', '').replace('.', '').replace(' ', '')
    nriptu_limpo = nriptu.replace('.', '').replace('-', '').replace(' ', '')

    tem_cbm = len(nrcbm_limpo) >= 2
    tem_iptu = len(nriptu_limpo) >= 4

    # Se não tem nenhuma chave, nem tenta navegar
    if not tem_cbm and not tem_iptu:
        return False, "Sem CBMERJ e sem IPTU — nada para buscar"

    try:
        page.goto(URL_IFRAME_BOMBEIROS, wait_until="networkidle", timeout=60000)
        page.wait_for_selector("input[name='cbmerj']", state="visible", timeout=15000)

        # ── Tentativa 1: por CBMERJ ──
        if tem_cbm:
            inscricao = nrcbm_limpo[:-1]
            dv = nrcbm_limpo[-1]

            console.print(f"[cyan][PW][/] Buscando por CBMERJ: {inscricao}-{dv}")
            page.fill("input[name='cbmerj']", inscricao)
            page.fill("input[name='cbmerj_dv']", dv)

            sucesso, erro = _submeter_busca_bombeiros(page, {
                'cbmerj': inscricao,
                'cbmerj_dv': dv,
            })

            if sucesso:
                return True, ""

            # Não encontrou — tenta inscrição predial
            if tem_iptu:
                console.print(f"[yellow][PW][/] CBMERJ não encontrado, tentando por inscrição predial...")
                page.goto(URL_IFRAME_BOMBEIROS, wait_until="networkidle", timeout=60000)
                page.wait_for_selector("input[name='cbmerj']", state="visible", timeout=15000)
            else:
                return False, erro

        # ── Tentativa 2: por inscrição predial (IPTU) ──
        if tem_iptu:
            # Clica na aba "Inba Predial" para ativar o select de município
            page.evaluate("() => { const btn = document.querySelector('#btnIP'); if (btn) btn.click(); }")
            page.wait_for_timeout(500)
            cod_municipio = _resolver_codigo_municipio(page, cidade)
            console.print(f"[cyan][PW][/] Buscando por inscrição predial: {nriptu_limpo} | Município: {cidade or 'Rio de Janeiro'} ({cod_municipio})")

            sucesso, erro = _submeter_busca_bombeiros(
                page,
                campos={
                    'inscricao': nriptu_limpo,
                    'municipio': cod_municipio,
                },
                destino='imovel/busca-inscricao-predial',
                endpoint='/sistema/imovel/busca-inscricao-predial',
            )

            if sucesso:
                return True, ""
            return False, erro

        return False, "Nenhuma busca teve sucesso"

    except Exception as e:
        return False, f"Erro na navegação: {str(e)}"


def _extrair_debitos_bombeiros(page: Page) -> list[dict]:
    """
    Parseia a tabela de vencimentos da página de resultado.
    Retorna lista de dicts com dados de cada exercício com DÉBITO.
    """
    debitos = []
    # HTML pode estar em window._htmlBombeiros (via fetch) ou no page.content() (navegação real)
    html = page.evaluate("() => window._htmlBombeiros || null") or page.content()
    soup = BeautifulSoup(html, 'html.parser')

    # Procura tabela de vencimentos (busca flexível)
    tabela = None
    for table in soup.find_all('table'):
        headers = [th.get_text(strip=True).lower() for th in table.find_all('th')]
        if 'exercício' in headers or 'exercicio' in headers or 'vencimento' in headers:
            tabela = table
            break

    if not tabela:
        # Tenta encontrar por texto "Vencimentos" no HTML
        h_venc = soup.find(string=re.compile(r'Vencimentos', re.IGNORECASE))
        if h_venc:
            proximo = h_venc.find_next('table')
            if proximo:
                tabela = proximo

    if not tabela:
        console.print("[yellow][AVISO][/] Nenhuma tabela de vencimentos encontrada")
        return debitos

    for linha in tabela.find_all('tr')[1:]:  # pula header
        colunas = linha.find_all('td')
        if len(colunas) < 6:
            continue

        situacao_el = colunas[5].find('span') or colunas[5]
        situacao = situacao_el.get_text(strip=True).upper()

        # Busca form de boleto na linha inteira
        form = linha.find('form')
        hidden_fields = {}
        if form:
            for inp in form.find_all('input', {'type': 'hidden'}):
                nome = inp.get('name', '')
                valor = inp.get('value', '')
                if nome:
                    hidden_fields[nome] = valor

        exercicio_raw = colunas[0].get_text(strip=True)
        exercicio_limpo = re.sub(r'\D', '', exercicio_raw)  # Remove tudo que não for dígito
        item = {
            'exercicio': exercicio_limpo or exercicio_raw,  # Fallback pro original se ficar vazio
            'vencimento': colunas[1].get_text(strip=True),
            'taxa': colunas[2].get_text(strip=True),
            'multa': colunas[3].get_text(strip=True),
            'situacao': situacao,
            'form_fields': hidden_fields,
            'tem_boleto': bool(form),
        }
        # Só inclui exercícios com débito (ignora QUITADO, DESCARTADO, etc.)
        if 'DÉBITO' in situacao or 'DEBITO' in situacao:
            debitos.append(item)

    com_debito = [d for d in debitos if 'DÉBITO' in d['situacao'] or 'DEBITO' in d['situacao']]
    return com_debito


def _baixar_boleto_bombeiros(page: Page, form_fields: dict) -> bytes | None:
    """
    Gera o boleto via Playwright interceptando a resposta do POST.
    Retorna bytes do PDF ou None.
    """
    url_boleto = 'https://www.funesbom.rj.gov.br/sistema/pagamentos/gerar/boletoSemRegistro'

    try:
        # Monta form dinâmico via JS e submete interceptando a resposta
        # Ao invés de clicar no botão (que abre nova aba), fazemos fetch
        fields_safe = {k: str(v).replace('"', '\\"') for k, v in form_fields.items()}
        fields_js = ', '.join([f'"{k}": "{v}"' for k, v in fields_safe.items()])

        resultado = page.evaluate(f"""async () => {{
            const formData = new URLSearchParams({{{fields_js}}});
            const resp = await fetch("{url_boleto}", {{
                method: 'POST',
                headers: {{
                    'Content-Type': 'application/x-www-form-urlencoded',
                }},
                body: formData.toString(),
            }});

            if (!resp.ok) {{
                return {{ error: 'HTTP ' + resp.status, data: null }};
            }}

            const contentType = resp.headers.get('content-type') || '';
            const blob = await resp.blob();

            // Converte pra base64
            return new Promise((resolve) => {{
                const reader = new FileReader();
                reader.onloadend = () => {{
                    const base64 = reader.result.split(',')[1];
                    resolve({{ error: null, data: base64, contentType: contentType, size: blob.size }});
                }};
                reader.readAsDataURL(blob);
            }});
        }}""")

        if resultado and resultado.get('data'):
            pdf_bytes = base64.b64decode(resultado['data'])
            content_type = resultado.get('contentType', '')
            tamanho = resultado.get('size', len(pdf_bytes))
            tam_fmt = f"{tamanho / 1024:.1f} KB" if tamanho >= 1024 else f"{tamanho} B"
            if tamanho < 5000:
                console.print(f"[yellow][AVISO][/] Boleto recebido: {tam_fmt}, tipo: {content_type} — tamanho suspeito, pode não ser PDF válido")
            else:
                console.print(f"[green][OK][/] Boleto recebido: {tam_fmt}, tipo: {content_type}")

            # Verifica se é realmente PDF
            if pdf_bytes[:4] == b'%PDF':
                return pdf_bytes
            else:
                # Pode ser HTML de erro
                try:
                    texto = pdf_bytes.decode('utf-8', errors='replace')[:500]
                    console.print(f"[yellow][AVISO][/] Resposta não é PDF: {texto}")
                except:
                    pass
                return None
        else:
            erro = resultado.get('error', 'desconhecido') if resultado else 'sem resposta'
            console.print(f"[red][ERRO][/] Falha ao baixar boleto: {erro}")
            return None

    except Exception as e:
        console.print(f"[red][ERRO][/] Exceção ao baixar boleto: {str(e)}")
        return None


def extrairbombeiros_hibrido(page: Page, objeto, linha, dataatual=''):
    """
    Extrai boleto de bombeiros usando Playwright + AntiCaptcha reCAPTCHA v2.
    Retorna: (dados_list, df) — dados_list: [codigo, nrcbm, ano, status], df: DataFrame ou None
    """
    if dataatual == '':
        dataatual = aux.hora('America/Sao_Paulo', 'DATA')

    cod = str(linha[Codigo]).strip()
    nrcbm = str(linha[NrCBM]).strip()
    nrcbm_limpo = nrcbm.replace('-', '').replace('.', '').replace(' ', '')
    # Coluna IPTU (índice 2) — usada como fallback na busca
    nriptu = str(linha[NrIPTU_CBM]).strip() if len(linha) > NrIPTU_CBM else ''
    # Coluna cidade (índice 3) — nome da cidade para o select de município
    cidade = str(linha[NrCidade]).strip() if len(linha) > NrCidade else ''
    anoextracao = str(date.today().year)

    # Se cod vazio (teste avulso sem banco), usa o CBM limpo como identificador
    tem_cod_cliente = bool(cod)
    if not cod:
        cod = nrcbm_limpo

    os.makedirs(objeto.pastadownload, exist_ok=True)

    # Verifica se já existe algum arquivo desse cliente na pasta (skip por cliente)
    prefixo = f"{cod}_{nrcbm_limpo}_" if tem_cod_cliente else f"{nrcbm_limpo}_"
    arquivos_existentes = [
        f for f in os.listdir(objeto.pastadownload)
        if f.startswith(prefixo) and f.endswith('.pdf')
    ]
    if arquivos_existentes:
        console.print(f"[yellow][SKIP][/] Cliente {cod} já tem {len(arquivos_existentes)} arquivo(s) na pasta — pulando")
        return [cod, nrcbm, anoextracao, f'SKIP ({len(arquivos_existentes)} arquivo(s) existente(s))'], None

    # Navegação (tenta CBMERJ, fallback pra inscrição predial/IPTU)
    sucesso, erro = _navegar_bombeiros(page, nrcbm, nriptu=nriptu, cidade=cidade)
    if not sucesso:
        return [cod, nrcbm, anoextracao, erro], None

    # Extrai débitos
    debitos = _extrair_debitos_bombeiros(page)
    if not debitos:
        return [cod, nrcbm, anoextracao, 'Sem débitos'], None

    console.print(f"[cyan][PW][/] {len(debitos)} débito(s) encontrado(s)")
    for d in debitos:
        console.print(f"  Exercício {d['exercicio']} | Venc: {d['vencimento']} | "
                       f"Taxa: {d['taxa']} | Multa: {d['multa']} | Sit: {d['situacao']}")

    # Baixa boleto de cada exercício com débito
    df_total = None
    baixados = 0
    for debito in debitos:
        ano = debito['exercicio']
        # Nome: {cod}_{cbm}_{ano} (lote) ou {cbm}_{ano} (teste avulso)
        if tem_cod_cliente:
            nome_arquivo = f"{cod}_{nrcbm_limpo}_{ano}.pdf"
        else:
            nome_arquivo = f"{nrcbm_limpo}_{ano}.pdf"
        caminhodestino = os.path.join(objeto.pastadownload, nome_arquivo)

        if os.path.isfile(caminhodestino):
            console.print(f"[yellow][SKIP][/] Já existe: {os.path.basename(caminhodestino)}")
            baixados += 1
            continue

        if not debito['form_fields']:
            console.print(f"[yellow][AVISO][/] Exercício {ano} sem formulário de boleto")
            continue

        pdf_bytes = _baixar_boleto_bombeiros(page, debito['form_fields'])
        if pdf_bytes:
            with open(caminhodestino, 'wb') as f:
                f.write(pdf_bytes)
            tamanho_kb = len(pdf_bytes) / 1024

            # Adiciona cabeçalho (mesmo padrão do fluxo-pw-new: origem = destino)
            try:
                aux.adicionarcabecalhopdf_topo_adaptativo(caminhodestino, caminhodestino, cod)
                console.print(f"[green][OK][/] Boleto {ano}: {os.path.basename(caminhodestino)} "
                              f"({tamanho_kb:.1f} KB) [dim]— cabeçalho: {cod}[/]")
            except Exception as e:
                console.print(f"[green][OK][/] Boleto {ano}: {os.path.basename(caminhodestino)} "
                              f"({tamanho_kb:.1f} KB) [yellow](sem cabeçalho: {e})[/]")

            baixados += 1

            # Extrai dados do PDF
            try:
                df = aux.extrairtextopdf(caminhodestino, 'BOLETOS')
                if df is not None and not df.empty:
                    total_rows = df[df.columns[0]].count()
                    df.insert(loc=0, column='Codigo', value=["'" + cod + "'" for _ in range(total_rows)])
                    df.insert(loc=4, column='TpoPagto', value=["'PARCELADO'" for _ in range(total_rows)])
                    df.insert(loc=5, column='Arquivo', value=["'" + caminhodestino + "'" for _ in range(total_rows)])
                    df_total = pd.concat([df_total, df]) if df_total is not None else df
            except Exception as e:
                console.print(f"[yellow][AVISO][/] Erro ao extrair dados do PDF {ano}: {str(e)}")
        else:
            console.print(f"[red][ERRO][/] Falha ao baixar boleto {ano}")

    if baixados > 0:
        console.print(f"\n[bold green][RESUMO][/] {baixados}/{len(debitos)} boleto(s) baixado(s)")
        return [cod, nrcbm, anoextracao, f'OK ({baixados} boleto(s))'], df_total
    console.print(f"\n[bold red][RESUMO][/] Nenhum boleto baixado")
    return [cod, nrcbm, anoextracao, 'Não conseguiu baixar boleto'], None


def testar_bombeiros(inscricao: str = "3605693-5", iptu: str = "",
                     cidade: str = "", cod: str = "",
                     pasta_download: str = "Downloads/Bombeiros", headless: bool = False):
    """Testa a extração de Bombeiros chamando extrairbombeiros_hibrido."""
    from playwright.sync_api import sync_playwright
    from types import SimpleNamespace

    caminho_perfil = os.path.join(
        os.path.dirname(__file__), "Profile",
        os.getenv('NOMEPROFILECBM', 'cbm_profile')
    )

    # Monta objeto e linha no mesmo formato que o lote usa
    # linha: [Codigo, CBM, IPTU, Cidade]
    objeto = SimpleNamespace(pastadownload=pasta_download)
    linha = [cod, inscricao, iptu, cidade]

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=caminho_perfil,
            headless=headless,
            accept_downloads=True,
            args=["--disable-blink-features=AutomationControlled", "--start-maximized"]
        )
        page = context.pages[0] if context.pages else context.new_page()

        try:
            dados, df = extrairbombeiros_hibrido(page, objeto, linha)
            print(f"\n[RESULTADO] {dados}")
            if df is not None:
                print(df.to_string(index=False))
        finally:
            context.close()


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Extrator IPTU Rio - Híbrido")
    parser.add_argument("modo", nargs="?", default="iptu",
                        choices=["iptu", "nadaconsta", "certidao", "bombeiros"],
                        help="Modo de extração (padrão: iptu)")
    parser.add_argument("-i", "--inscricao", default="3.146.968-7",
                        help="Número de inscrição (ex: 3.146.968-7)")
    parser.add_argument("-t", "--tipo", default="1",
                        help="Tipo de pagamento: 1=Cota Única, 2=1ª data, 3=2ª data, 4=3ª data (só IPTU)")
    parser.add_argument("-p", "--pasta", default="Downloads/Bombeiros",
                        help="Pasta para salvar o PDF")
    parser.add_argument("--iptu", default="",
                        help="Inscrição predial/IPTU (fallback pra bombeiros)")
    parser.add_argument("--cidade", default="",
                        help="Nome da cidade para busca por inscrição predial (default: Rio de Janeiro)")
    parser.add_argument("--cod", default="",
                        help="Código do cliente (usado no nome do arquivo e cabeçalho do PDF)")
    parser.add_argument("--visible", action="store_true",
                        help="Abrir janela do navegador (padrão: headless)")
    args = parser.parse_args()

    match args.modo:
        case "iptu":
            testar_extracao(
                inscricao=args.inscricao,
                tipo_pagamento=args.tipo,
                pasta_download=args.pasta,
                headless=not args.visible,
            )
        case "nadaconsta":
            if not args.visible:
                print("[AVISO] Nada Consta requer modo visível (reCAPTCHA v3). Forçando --visible.")
            testar_nadaconsta(
                inscricao=args.inscricao,
                pasta_download=args.pasta,
                headless=False,
            )
        case "certidao":
            testar_certidao(
                inscricao=args.inscricao,
                pasta_download=args.pasta,
            )
        case "bombeiros":
            testar_bombeiros(
                inscricao=args.inscricao,
                iptu=args.iptu,
                cidade=args.cidade,
                cod=args.cod,
                pasta_download=args.pasta,
                headless=not args.visible,
            )
