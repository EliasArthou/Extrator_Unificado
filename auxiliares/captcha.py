"""
captcha.py - Resolvedor unificado de CAPTCHA via AntiCaptcha.

Centraliza a logica de resolucao de CAPTCHA usada pelos extratores
(IPTU, Certidao Negativa, Bombeiros) e pelo wrapper Selenium antigo
(core/web.py).

Tipos suportados:
    - imagem        (CAPTCHA de imagem)
    - recaptcha v2  (checkbox "Nao sou um robo")

Configuracao: CHAVEANTICAPTCHA deve estar setada no .env.
"""

from __future__ import annotations

import os
import threading
import time as _time

import requests
from anticaptchaofficial.imagecaptcha import imagecaptcha
from anticaptchaofficial.recaptchav2proxyless import recaptchaV2Proxyless
from playwright.sync_api import Page

from auxiliares.utils import console, _print_lock


# ============================================================
# CAPTCHA de IMAGEM
# ============================================================

def _resolver_captcha_imagem(session: requests.Session, url_captcha: str,
                              pasta_download: str,
                              url_referer: str = "") -> tuple[str | None, imagecaptcha | None]:
    """
    Baixa a imagem do CAPTCHA e resolve via AntiCaptcha.

    Retorna: (texto_resolvido, solver) - o solver e retornado pra poder reportar
    erro depois via solver.report_incorrect_image_captcha().
    """
    import tempfile
    fd, caminho_captcha = tempfile.mkstemp(suffix=".png", prefix="captcha_", dir=pasta_download)
    os.close(fd)

    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "image/webp,image/apng,image/*,*/*;q=0.8",
        }
        if url_referer:
            headers["Referer"] = url_referer

        resp = session.get(url_captcha, headers=headers, timeout=30)
        content_type = resp.headers.get("Content-Type", "")
        if resp.status_code == 200 and len(resp.content) > 100:
            # Verifica se realmente e imagem (servidor pode retornar HTML de erro)
            if "text/html" in content_type or resp.content[:5] in (b"<html", b"<!DOC", b"<HTML"):
                console.print(
                    f"[yellow]\\[CAPTCHA] Servidor retornou HTML em vez de imagem "
                    f"({len(resp.content)} bytes). Aguardando 10s...[/yellow]"
                )
                _time.sleep(10)
                return None, None

            # Converte pra PNG (AntiCaptcha so aceita JPEG/PNG/GIF)
            from PIL import Image
            from io import BytesIO
            img = Image.open(BytesIO(resp.content))
            img.save(caminho_captcha, "PNG")

            solver = imagecaptcha()
            solver.set_verbose(0)
            solver.set_key(os.getenv("CHAVEANTICAPTCHA"))

            texto = solver.solve_and_return_solution(caminho_captcha)

            if texto and len(str(texto)) > 0:
                return texto, solver
        else:
            console.print(
                f"[yellow]\\[CAPTCHA] Imagem muito pequena ou erro: "
                f"status={resp.status_code}, tamanho={len(resp.content)} bytes[/yellow]"
            )
    except Exception as e:
        console.print(f"[red]\\[CAPTCHA] Erro ao resolver: {e}[/red]")
    finally:
        # Sempre apaga o captcha.png, mesmo em caso de erro
        if os.path.isfile(caminho_captcha):
            try:
                os.remove(caminho_captcha)
            except Exception:
                pass

    return None, None


# ============================================================
# DETECCAO de tipo de CAPTCHA
# ============================================================

def _detectar_tipo_captcha(page: Page | None) -> str | None:
    """
    Detecta o tipo de CAPTCHA presente na pagina Playwright.
    Retorna: 'recaptchav2', 'recaptchav3', 'imagem' ou None.
    """
    if page is None:
        return None

    try:
        info = page.evaluate("""() => {
            // reCAPTCHA v2 - checkbox visivel ou iframe anchor
            const hasV2Iframe = !!document.querySelector('iframe[src*="recaptcha"][src*="anchor"]');
            const hasV2Checkbox = !!document.querySelector('.g-recaptcha');
            const hasV2Textarea = !!document.querySelector('textarea[name="g-recaptcha-response"]');

            // reCAPTCHA v3 - invisivel, geralmente so script com render=sitekey
            const scripts = Array.from(document.querySelectorAll('script[src*="recaptcha"]'));
            const hasV3Render = scripts.some(s => s.src.includes('render=') && !s.src.includes('render=explicit'));

            // CAPTCHA de imagem - procura img com src contendo captcha
            const hasImageCaptcha = !!document.querySelector('img[src*="captcha"]');

            return { hasV2Iframe, hasV2Checkbox, hasV2Textarea, hasV3Render, hasImageCaptcha };
        }""")

        if info.get("hasV2Iframe") or info.get("hasV2Checkbox") or info.get("hasV2Textarea"):
            return "recaptchav2"
        if info.get("hasV3Render"):
            return "recaptchav3"
        if info.get("hasImageCaptcha"):
            return "imagem"
        return None
    except Exception:
        return None


def _detectar_sitekey_v2(page: Page) -> str | None:
    """Extrai o sitekey do reCAPTCHA v2 da pagina."""
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


# ============================================================
# reCAPTCHA v2 (com timer)
# ============================================================

def _resolver_recaptcha_v2(page_url: str, sitekey: str) -> str | None:
    """
    Resolve reCAPTCHA v2 usando AntiCaptcha (recaptchaV2Proxyless).
    Mostra progresso com timer atualizado a cada 5 segundos.
    Retorna o token g-recaptcha-response ou None.
    """
    chave_api = os.getenv("CHAVEANTICAPTCHA")
    if not chave_api:
        console.print("[red][ERRO][/] CHAVEANTICAPTCHA nao configurada no .env")
        return None

    if not sitekey:
        console.print("[red][ERRO][/] sitekey nao fornecida para reCAPTCHA v2")
        return None

    solver = recaptchaV2Proxyless()
    solver.set_verbose(0)
    solver.set_key(chave_api)
    solver.set_website_url(page_url)
    solver.set_website_key(sitekey)

    # Thread separada pra mostrar timer enquanto resolve
    resultado: list[str | None] = [None]
    erro: list[str | None] = [None]
    finalizado = threading.Event()

    def _resolver():
        # Supressao do ConnectionResetError feita pelo _FilteredStream no modo lote
        try:
            resultado[0] = solver.solve_and_return_solution()
            if not resultado[0]:
                erro[0] = str(solver.error_code)
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

    import sys as _sys

    # Linha inicial com timer inline (usa \r pra atualizar sem quebrar linha)
    with _print_lock:
        _sys.stdout.write(f"[CAPTCHA] Resolvendo reCAPTCHA v2... 00:00")
        _sys.stdout.flush()

    # Atualiza timer a cada 1 segundo na mesma linha
    while not finalizado.wait(timeout=1):
        elapsed_parcial = _time.time() - inicio
        with _print_lock:
            # \r volta o cursor pro inicio + padding com espacos pra apagar residuos
            _sys.stdout.write(f"\r[CAPTCHA] Resolvendo reCAPTCHA v2... {_fmt(elapsed_parcial)}")
            _sys.stdout.flush()

    elapsed = _time.time() - inicio

    # Limpa a linha do timer antes de imprimir resultado final
    with _print_lock:
        _sys.stdout.write("\r" + " " * 60 + "\r")
        _sys.stdout.flush()

    if resultado[0]:
        with _print_lock:
            console.print(f"[green]\\[CAPTCHA] Resolvido em {_fmt(elapsed)} ({len(resultado[0])} chars)[/green]")
        return resultado[0]
    else:
        with _print_lock:
            console.print(f"[red]\\[CAPTCHA] Falha em {_fmt(elapsed)}: {erro[0]}[/red]")
        return None


# ============================================================
# WRAPPER unificado
# ============================================================

def resolver_captcha(tipo: str = "auto", *,
                     # -- parametros para imagem (tipo='imagem') --
                     session: requests.Session | None = None,
                     url_captcha: str = "",
                     pasta_download: str = "",
                     url_referer: str = "",
                     # -- parametros para recaptcha v2 (tipo='recaptchav2') --
                     page_url: str = "",
                     sitekey: str = "",
                     # -- parametros para auto-deteccao (tipo='auto') --
                     page: Page | None = None,
                     ) -> tuple[str | None, object | None]:
    """
    Resolvedor unificado de CAPTCHA via AntiCaptcha.

    Tipos suportados:
        'imagem'       - CAPTCHA de imagem (baixa e resolve via imagecaptcha)
        'recaptchav2'  - reCAPTCHA v2 checkbox (resolve via recaptchaV2Proxyless)
        'auto'         - Detecta automaticamente a partir da page do Playwright

    Retorna: (solucao, solver)
        - imagem: (texto, solver_imagecaptcha) - solver pra report_incorrect se precisar
        - recaptchav2: (token, None) - token g-recaptcha-response
        - None em caso de erro
    """
    if tipo == "auto":
        tipo = _detectar_tipo_captcha(page)
        if not tipo:
            console.print("[yellow][CAPTCHA][/] Nenhum CAPTCHA detectado na pagina")
            return None, None
        console.print(f"[cyan][CAPTCHA][/] Tipo detectado: {tipo}")

    if tipo == "imagem":
        return _resolver_captcha_imagem(session, url_captcha, pasta_download, url_referer)
    elif tipo == "recaptchav2":
        token = _resolver_recaptcha_v2(page_url, sitekey)
        return token, None
    else:
        console.print(f"[red][CAPTCHA][/] Tipo nao suportado: {tipo}")
        return None, None
