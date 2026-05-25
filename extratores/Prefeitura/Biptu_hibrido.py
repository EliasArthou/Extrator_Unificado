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
from auxiliares import utils as aux
from auxiliares import aux_patches  # registra adicionarcabecalhopdf_topo_adaptativo em aux
from auxiliares.utils import console, _print_lock, _flush_log
from auxiliares.captcha import (
    resolver_captcha,
    _detectar_sitekey_v2,
    _detectar_tipo_captcha,
    _resolver_captcha_imagem,
    _resolver_recaptcha_v2,
)
from datetime import date, datetime
from bs4 import BeautifulSoup
from playwright.sync_api import Page, Response
from dotenv import load_dotenv
import logging
import urllib3

load_dotenv()

# Suprime warnings de conexão do requests/urllib3 (ex: ConnectionResetError ao reusar sessão)
logging.getLogger("urllib3").setLevel(logging.CRITICAL)
logging.getLogger("urllib3.util.retry").setLevel(logging.CRITICAL)
logging.getLogger("requests").setLevel(logging.CRITICAL)
urllib3.disable_warnings()
import warnings
warnings.filterwarnings('ignore', module='urllib3')
warnings.filterwarnings('ignore', message='.*Connection.*')


def _esta_visivel(loc, timeout_ms: int = 2000) -> bool:
    """Compat com Playwright >= 1.40: is_visible() não aceita mais timeout=.
    Substituto: wait_for(state='visible', timeout=N) dentro de try/except."""
    try:
        loc.wait_for(state='visible', timeout=timeout_ms)
        return True
    except Exception:
        return False


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

            # Sempre navega de volta se o input não está visível
            # (a URL é a mesma para tela inicial e tela de dados)
            input_insc = page.locator("#ctl00_ePortalContent_inscricao_input")
            if not _esta_visivel(input_insc, 2000):
                page.goto(os.getenv('SITEIPTU'), wait_until="networkidle", timeout=60000)

            page.wait_for_load_state("networkidle")

            input_insc = page.locator("#ctl00_ePortalContent_inscricao_input")
            input_insc.wait_for(state="visible", timeout=15000)
            input_insc.fill(nriptu)

            try:
                with page.expect_navigation(wait_until="networkidle", timeout=30000):
                    page.locator("input[value='PROSSEGUIR']").click()
            except:
                pass

            msg = page.locator("#ctl00_ePortalContent_MSG")
            if _esta_visivel(msg, 2000):
                txt = msg.inner_text().strip()
                if ERRO_SESSAO in txt:
                    page.goto(os.getenv('SITEIPTU'), wait_until="networkidle", timeout=30000)
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
                page.goto(os.getenv('SITEIPTU'), wait_until="networkidle", timeout=30000)
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
        # 1. Marca o checkbox da data via .click() nativo do DOM.
        # Por que não usar Playwright .check(): o site renderiza o checkbox
        # invisível (0x0) e o Playwright recusa clicar nele mesmo com force=True.
        # Por que .click() em vez de dispatchEvent('change'): cbCotaUnica usa
        # onchange (SelecionarDarmUnico), mas Chk_001/002/003 usam onclick
        # (SelecionaLimpa). O .click() nativo dispara ambos os handlers.
        sel_id = "[id*='cbCotaUnica']" if tipo_pag == '1' else f"[id*='Chk_00{int(tipo_pag) - 1}']"
        loc = page.locator(sel_id)
        loc.wait_for(state='attached', timeout=10000)
        loc.evaluate("el => el.click()")

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
            # wait_for_timeout: dá tempo pro JS do site processar o checkbox
            # (SelecionarDarmUnico) e habilitar o botão. force=True pula o check
            # de visibilidade — mesmo motivo do dispatchEvent no checkbox.
            page.wait_for_timeout(300)
            page.locator("#ctl00_ePortalContent_btnDarmIndiv").click(force=True)

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

def _reset_pagina_iptu(page: Page):
    """Volta para a tela inicial do IPTU (campo de inscrição)."""
    try:
        page.goto(os.getenv('SITEIPTU'), wait_until="networkidle", timeout=15000)
    except Exception:
        try:
            page.goto(os.getenv('SITEIPTU'))
        except Exception:
            pass


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
                _reset_pagina_iptu(page)
                return [cod, nriptu, '', '', '', '', '', erro], None

            # ── Extrai dados da tela (contribuinte, endereço, exercício) ──
            contribuinte = _texto_locator(page, "#ctl00_ePortalContent_TELA_CONTRIBUINTE")
            endereco = _texto_locator(page, "#ctl00_ePortalContent_TELA_ENDERECO")
            exercicio = _texto_locator(page, "#ctl00_ePortalContent_GuiaExercicio")

            # ── Clica na guia ──
            guias = page.locator('[title*="visualizar / imprimir"]')
            if guias.count() == 0:
                # Antes de chamar de "Verificar Manualmente", checa se o IPTU está pago.
                # O site preenche o input #ctl00_ePortalContent_MENSAGEM com "Quitada..."
                # quando não há boleto a emitir porque o imóvel já quitou o exercício.
                mensagem = ""
                try:
                    mensagem = page.locator(
                        "#ctl00_ePortalContent_MENSAGEM"
                    ).input_value(timeout=1000) or ""
                except Exception:
                    pass

                if "quitad" in mensagem.lower():
                    _reset_pagina_iptu(page)
                    return [cod, nriptu, exercicio, '', '', contribuinte, endereco,
                            'PAGO (Quitada)'], None

                # Caso não seja "Quitada": dump de debug + retorna "Verificar Manualmente"
                try:
                    pasta_debug = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                               "Logs", "debug_sem_guia")
                    os.makedirs(pasta_debug, exist_ok=True)
                    ts = datetime.now().strftime("%H%M%S")
                    base = os.path.join(pasta_debug, f"{cod}_{nriptu}_{ts}")
                    page.screenshot(path=f"{base}.png", full_page=True)
                    with open(f"{base}.html", "w", encoding="utf-8") as f:
                        f.write(page.content())
                except Exception:
                    pass  # debug não pode quebrar fluxo
                _reset_pagina_iptu(page)
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
                        aux.adicionarcabecalhopdf_topo_adaptativo(
                            caminhodestino, caminhodestino, cod)
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
        _reset_pagina_iptu(page)

        return dadosiptu, df

    except Exception as e:
        _reset_pagina_iptu(page)
        return [cod, nriptu, '', '', '', '', '', f"Erro: {str(e)}"], None


# ─────────────────────────────────────────────
# NADA CONSTA (híbrido — mesmo portal iportal)
# ─────────────────────────────────────────────

def _navegar_nadaconsta(page: Page, nriptu: str) -> tuple[bool, str, str]:
    """
    Playwright: preenche inscrição + PROSSEGUIR no portal Nada Consta.
    O portal usa AJAX (sem navegação de página) — aguarda #PDF ou #MSG.
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

            # Garante que #PDF e #MSG estão ocultos antes de clicar
            # (reset visual de requisição anterior na mesma sessão)
            page.evaluate("""
                document.getElementById('PDF') && (document.getElementById('PDF').innerHTML = '');
                document.getElementById('MSG') && document.getElementById('MSG').classList.add('hidden');
            """)

            # Aguarda reCAPTCHA v3 carregar antes de clicar
            try:
                page.wait_for_function(
                    "() => typeof grecaptcha !== 'undefined' && typeof grecaptcha.execute === 'function'",
                    timeout=10000
                )
            except:
                pass  # Continua mesmo que não detecte — pode funcionar sem

            # Clica PROSSEGUIR — portal usa reCAPTCHA v3 + AJAX, sem navegação
            page.locator("#Avancar").click()

            # Aguarda o AJAX terminar: #PDF visível (sucesso) ou #MSG visível (erro)
            # O JS injeta a resposta em #PDF quando tem <object>, ou mostra #MSG
            try:
                page.wait_for_function(
                    """() => {
                        const pdf = document.getElementById('PDF');
                        const msg = document.getElementById('MSG');
                        const pdfOk = pdf && !pdf.classList.contains('hidden') && pdf.innerHTML.trim() !== '';
                        const msgOk = msg && !msg.classList.contains('hidden');
                        return pdfOk || msgOk;
                    }""",
                    timeout=30000
                )
            except:
                pass

            # Verifica se é erro
            msg_el = page.locator("#MSG")
            try:
                if _esta_visivel(msg_el, 1000):
                    txt = page.locator("#MSGText").inner_text(timeout=1000).strip()
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
                # ── AJAX injetou a resposta em #PDF — agora o content() tem o HTML ──
                os.makedirs(os.path.dirname(caminhodestino) or ".", exist_ok=True)

                # 1ª tentativa: base64 no HTML do #PDF (caminho principal)
                html_pdf_div = page.locator("#PDF").inner_html(timeout=5000)
                pdf_b64 = _extrair_pdf_base64(html_pdf_div)

                if pdf_b64:
                    with open(caminhodestino, "wb") as f:
                        f.write(base64.b64decode(pdf_b64))
                else:
                    # 2ª tentativa: lê href do link "aqui" diretamente do atributo
                    links_aqui = page.locator("#PDF a")
                    qtd_links = links_aqui.count()

                    if qtd_links == 0:
                        dadosiptu = [cod, nriptu, anoextracao, 'Verificar (Extrair Manualmente)']
                    else:
                        for i in range(qtd_links):
                            nome_arquivo = (f"{cod}_{nriptu}.pdf" if i == 0
                                            else f"{cod}_{nriptu}_{i}.pdf")
                            destino_i = os.path.join(objeto.pastadownload, nome_arquivo)

                            try:
                                href = links_aqui.nth(i).get_attribute("href") or ""

                                if href.startswith("data:application/pdf;base64,"):
                                    b64 = href.split(",", 1)[1].strip()
                                    with open(destino_i, "wb") as f:
                                        f.write(base64.b64decode(b64))
                                elif href.startswith("http"):
                                    import requests as _req
                                    resp = _req.get(href, timeout=60)
                                    resp.raise_for_status()
                                    with open(destino_i, "wb") as f:
                                        f.write(resp.content)
                                else:
                                    with page.expect_download(timeout=60000) as dl_info:
                                        links_aqui.nth(i).click()
                                    dl_info.value.save_as(destino_i)

                                if i == 0:
                                    caminhodestino = destino_i

                            except Exception as e:
                                dadosiptu = [cod, nriptu, anoextracao, f'Erro download: {e}']
                                break

                if os.path.isfile(caminhodestino):
                    try:
                        aux.adicionarcabecalhopdf_topo_adaptativo(
                            caminhodestino, caminhodestino, cod)
                    except Exception:
                        pass

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
            console.print(f"[green]\\[OK] Certidão salva em: {caminho_pdf}[/green]")
        finally:
            browser.close()

    # Remoção do HTML temporário feita pelo chamador (finally block)


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
            # Retry automático para ConnectionResetError / ConnectionError
            from urllib3.util.retry import Retry
            from requests.adapters import HTTPAdapter
            retry_strategy = Retry(
                total=3, backoff_factor=2,
                status_forcelist=[429, 500, 502, 503, 504],
            )
            session.mount('http://', HTTPAdapter(max_retries=retry_strategy))
            session.mount('https://', HTTPAdapter(max_retries=retry_strategy))

            session.headers.update({
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
            })

            captcha_tentativas = 0
            captcha_erros = 0

            def _captcha_msg(nriptu, tentativas, erros, extra=''):
                erros_txt = f', {erros} erro(s)' if erros else ''
                return f"\\[CAPTCHA] {nriptu}: tentativa {tentativas}{erros_txt}{extra}"

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

                    # 2. Resolve CAPTCHA (spinner só durante a resolução)
                    captcha_tentativas += 1
                    with console.status(
                        _captcha_msg(nriptu, captcha_tentativas, captcha_erros, ' — resolvendo...'),
                        spinner='dots'
                    ):
                        texto_captcha, solver_captcha = _resolver_captcha_imagem(
                            session, url_captcha, objeto.pastadownload, url_referer=url_base)
                    if not texto_captcha:
                        captcha_erros += 1
                        console.print(f"[yellow]{_captcha_msg(nriptu, captcha_tentativas, captcha_erros, ' — falhou')}[/yellow]")
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
                        captcha_erros += 1
                        if solver_captcha:
                            solver_captcha.report_incorrect_image_captcha()
                        continue

                    # Inscrição inválida?
                    if 'INSCRIÇÃO IMOBILIÁRIA INVÁLIDA' in html_resp.upper():
                        dadosiptu = [cod, nriptu, anoextracao, 'Inscrição inválida!', 'N/A']
                        resolveu = True
                        break

                    # Sessão expirada? Precisa refazer do zero (nova sessão)
                    if 'expirada' in html_resp.lower() or 'sessão expirada' in html_resp.lower():
                        session.close()
                        session = requests.Session()
                        session.mount('http://', HTTPAdapter(max_retries=retry_strategy))
                        session.mount('https://', HTTPAdapter(max_retries=retry_strategy))
                        session.headers.update({
                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                            'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
                        })
                        console.print(f"[yellow]\\[CERTIDÃO] Sessão expirada para {nriptu}, recriando sessão...[/yellow]")
                        continue

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
                    caminho_html_temp = os.path.abspath(
                        caminho_arquivo.replace('.pdf', '.html'))
                    try:
                        with open(caminho_html_temp, 'wb') as f:
                            f.write(resp.content)

                        _html_para_pdf(caminho_html_temp, caminho_arquivo,
                                       base_url=url_base)

                    except Exception as e:
                        dadosiptu = [cod, nriptu, anoextracao, f'Erro ao salvar: {str(e)}', 'N/A']
                        break
                    finally:
                        # Garante remoção do HTML temporário
                        if os.path.isfile(caminho_html_temp):
                            try:
                                os.remove(caminho_html_temp)
                            except Exception:
                                pass

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
                            aux.adicionarcabecalhopdf_topo_adaptativo(
                                caminho_arquivo, novonomearquivo, aux.left(cod, 4))
                            # Apaga o original se foi movido pra subpasta
                            if (os.path.normpath(caminho_arquivo) != os.path.normpath(novonomearquivo)
                                    and os.path.isfile(caminho_arquivo)
                                    and os.path.isfile(novonomearquivo)):
                                os.remove(caminho_arquivo)
                            caminho_arquivo = novonomearquivo
                        except Exception:
                            pass
                    else:
                        possui_divida = dfdivida['Possui_Divida'] if dfdivida else 'N/A'
                        dadosiptu = [cod, nriptu, anoextracao, 'Verificar (Extrair Manualmente)',
                                     possui_divida]

                except (ConnectionError, requests.exceptions.ConnectionError) as e:
                    # Servidor resetou conexão — aguarda antes de tentar de novo
                    import time
                    wait = min(5 * (tentativa + 1), 30)
                    console.print(f"[yellow]\\[CERTIDÃO] Conexão resetada, aguardando {wait}s...[/yellow]")
                    time.sleep(wait)
                    # Recria sessão (cookies/socket podem estar corrompidos)
                    session.close()
                    session = requests.Session()
                    session.mount('http://', HTTPAdapter(max_retries=retry_strategy))
                    session.mount('https://', HTTPAdapter(max_retries=retry_strategy))
                    session.headers.update({
                        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                        'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
                    })
                    if tentativa == limitetentativas - 1:
                        dadosiptu = [cod, nriptu, anoextracao, f'Erro: {str(e)}', 'N/A']
                except Exception as e:
                    if tentativa == limitetentativas - 1:
                        dadosiptu = [cod, nriptu, anoextracao, f'Erro: {str(e)}', 'N/A']

            session.close()

            # Resumo do captcha pra esta inscrição
            if captcha_tentativas > 0:
                cor = 'green' if resolveu else 'red'
                status_txt = 'OK' if resolveu else 'FALHA'
                erros_txt = f', {captcha_erros} erro(s)' if captcha_erros else ''
                console.print(f"[{cor}]\\[CAPTCHA] {nriptu}: {captcha_tentativas}x tentativa(s){erros_txt} — {status_txt}[/{cor}]")

            # Data limite (impresso após o resumo de captcha)
            if dfdivida is not None and 'Data_Limite' in dfdivida:
                console.print(f"[cyan]Data limite encontrada: {dfdivida['Data_Limite']}[/cyan]")

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


# Índices das colunas no resultado da query SQLCBM
# Codigo=0, CBM=1, IPTU=2, Cidade=3


# ─────────────────────────────────────────────
# CAPTCHA — Resolvedor unificado (imagem + reCAPTCHA v2)
# ─────────────────────────────────────────────


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Extrator IPTU Rio - Híbrido")
    parser.add_argument("modo", nargs="?", default="iptu",
                        choices=["iptu", "nadaconsta", "certidao"],
                        help="Modo de extração (padrão: iptu)")
    parser.add_argument("-i", "--inscricao", default="3.146.968-7",
                        help="Número de inscrição (ex: 3.146.968-7)")
    parser.add_argument("-t", "--tipo", default="1",
                        help="Tipo de pagamento: 1=Cota Única, 2=1ª data, 3=2ª data, 4=3ª data (só IPTU)")
    parser.add_argument("-p", "--pasta", default="Downloads/IPTU",
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
