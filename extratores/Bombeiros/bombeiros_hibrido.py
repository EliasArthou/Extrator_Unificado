"""
bombeiros_hibrido.py - Extrator de Taxa CBM (Bombeiros) hibrido.

Playwright navega + intercepta resposta + AntiCaptcha resolve reCAPTCHA v2.

Ordem de tentativa para localizar o imovel:
  1. Por CBMERJ (numero da taxa de incendio)
  2. Por inscricao predial/IPTU (fallback se CBMERJ falhar ou nao tiver)

Fluxo (extrairbombeiros_hibrido):
  1. _navegar_bombeiros: vai pra https://www.funesbom.rj.gov.br/sistema/imovel
     e tenta CBMERJ, depois IPTU (resolve reCAPTCHA v2)
  2. _extrair_debitos_bombeiros: parseia tabela de vencimentos
  3. _baixar_boleto_bombeiros: pra cada debito, gera o PDF via fetch
  4. Adiciona cabecalho no PDF com codigo do cliente

Modos de execucao:
  - Lote: extratores/Bombeiros/testar_bombeiros_lote.py
  - Avulso: python -m extratores.Bombeiros.bombeiros_hibrido --iptu X --cod Y --cidade Z
  - Via UI Tkinter: core/extracao.py:extrairboletosbombeiro (Pass 2)
"""

from __future__ import annotations

import os
import re
import base64
from datetime import date

import pandas as pd
from bs4 import BeautifulSoup
from playwright.sync_api import Page
from dotenv import load_dotenv

from auxiliares import utils as aux
from auxiliares import aux_patches  # noqa: F401  (registra adicionarcabecalhopdf_topo_adaptativo em aux)
from auxiliares.utils import console, _flush_log
from auxiliares.captcha import resolver_captcha, _detectar_sitekey_v2

load_dotenv()


# ============================================================
# Constantes
# ============================================================

SITEKEY_BOMBEIROS = "6LdZ1EQkAAAAAAKRMr1Dhld-8WiyW5Qt0HgCXyfa"
URL_IFRAME_BOMBEIROS = "https://www.funesbom.rj.gov.br/sistema/imovel"

# Indices das colunas no resultado da query SQLCBM
# Codigo=0, CBM=1, IPTU=2, Cidade=3
Codigo = 0
NrCBM = 1
NrIPTU_CBM = 2
NrCidade = 3


# ============================================================
# Helper: resolver codigo do municipio no select FUNESBOM
# ============================================================

def _resolver_codigo_municipio(page: Page, cidade: str) -> str:
    """Busca o codigo do municipio direto no select da pagina FUNESBOM.
    Faz match exato, depois case-insensitive, depois parcial.
    Default: '064' (Rio de Janeiro)."""
    if not cidade:
        return "064"
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
        // Parcial (contem)
        for (const o of opts) {{
            if (o.text.toUpperCase().includes(nomeUpper) || nomeUpper.includes(o.text.toUpperCase())) return o.value;
        }}
        return null;
    }}""")

    if resultado:
        return resultado
    console.print(f"[yellow][AVISO][/] Municipio '{cidade}' nao encontrado no combo FUNESBOM, usando Rio de Janeiro (064)")
    return "064"


# ============================================================
# Submissao do form de busca (via fetch)
# ============================================================

def _submeter_busca_bombeiros(page: Page, campos: dict,
                              destino: str = "imovel/busca",
                              endpoint: str = "/sistema/imovel/busca",
                              _log: list | None = None) -> tuple[bool, str]:
    """
    Submete a busca no site dos Bombeiros via fetch.
    campos: dict com os campos do form (cbmerj/cbmerj_dv ou inscricao/municipio).
    destino: valor do campo hidden 'destino' (varia por aba).
    endpoint: URL do POST (varia por aba).
    _log: se fornecido, acumula mensagens em vez de imprimir.
    Retorna: (sucesso, mensagem_erro)
    """
    _msg = lambda m: _log.append(m) if _log is not None else console.print(m)

    # Resolve reCAPTCHA v2 via AntiCaptcha
    sitekey = _detectar_sitekey_v2(page) or SITEKEY_BOMBEIROS
    token, _ = resolver_captcha("recaptchav2", page_url=URL_IFRAME_BOMBEIROS, sitekey=sitekey)
    if not token:
        return False, "Falha ao resolver reCAPTCHA v2"

    # Pega o _token CSRF do form
    csrf_token = page.evaluate("""() => {
        const input = document.querySelector('input[name="_token"]');
        return input ? input.value : '';
    }""")

    # Monta os campos do form
    campos_js = "\n".join([f"formData.append('{k}', '{str(v).replace(chr(39), chr(92)+chr(39))}');" for k, v in campos.items()])

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
        status = resultado.get("status", 0)
        html_resp = resultado.get("html", "")
        url_resp = resultado.get("url", "")
        tam_post = len(html_resp)
        tam_post_fmt = f"{tam_post / 1024:.1f} KB" if tam_post >= 1024 else f"{tam_post} B"
        if tam_post < 500:
            _msg(f"[yellow]\\[AVISO] POST resposta: status {status}, tamanho: {tam_post_fmt} - resposta muito pequena, possivel erro[/yellow]")
        else:
            _msg(f"[cyan]\\[PW] POST resposta: status {status}, tamanho: {tam_post_fmt}[/cyan]")

        if "Dados Cadastrais" in html_resp or "Vencimentos" in html_resp:
            # NAO usar page.set_content() - perde a sessao/cookies do site.
            # Guarda o HTML na page via window._htmlBombeiros para uso posterior.
            page.evaluate("(html) => { window._htmlBombeiros = html; }", html_resp)
            return True, ""
        elif "nao encontrado" in html_resp.lower() or "não encontrado" in html_resp.lower():
            return False, "Imovel nao encontrado"
        elif (
            "realize a consulta através do número da inscri" in html_resp.lower()
            or "realize a consulta atraves do numero da inscri" in html_resp.lower()
        ):
            # Cidades especiais (Macae, Sao Goncalo, Campos dos Goytacazes etc.)
            # exigem busca por inscricao predial/IPTU em vez de CBMERJ.
            return False, "CIDADE_ESPECIAL"
        else:
            trecho = html_resp[:500].replace("\n", " ")
            _msg(f"[dim]\\[DEBUG] Resposta: {trecho}[/dim]")
            return False, f"Resposta inesperada (status {status})"
    else:
        return False, "Sem resposta do servidor"


# ============================================================
# Navegacao: tenta CBMERJ, fallback pra inscricao predial
# ============================================================

def _navegar_bombeiros(page: Page, nrcbm: str, nriptu: str = "",
                       cidade: str = "", _log: list | None = None) -> tuple[bool, str]:
    """
    Playwright: navega ao iframe de busca dos Bombeiros.
    Ordem de tentativa:
      1. CBMERJ (se tiver)
      2. Inscricao predial / IPTU (se tiver)
    Se nao tiver nenhum dos dois, retorna erro sem navegar.
    (Busca por CPF/CNPJ nao implementada - requer autenticacao gov.br)
    """
    _msg = lambda m: _log.append(m) if _log is not None else console.print(m)

    nrcbm_limpo = nrcbm.replace("-", "").replace(".", "").replace(" ", "")
    nriptu_limpo = nriptu.replace(".", "").replace("-", "").replace(" ", "")

    tem_cbm = len(nrcbm_limpo) >= 2
    tem_iptu = len(nriptu_limpo) >= 4

    # Se nao tem nenhuma chave, nem tenta navegar
    if not tem_cbm and not tem_iptu:
        return False, "Sem CBMERJ e sem IPTU - nada para buscar"

    try:
        page.goto(URL_IFRAME_BOMBEIROS, wait_until="networkidle", timeout=60000)
        page.wait_for_selector("input[name='cbmerj']", state="visible", timeout=15000)

        # Tentativa 1: por CBMERJ
        if tem_cbm:
            inscricao = nrcbm_limpo[:-1]
            dv = nrcbm_limpo[-1]

            _msg(f"[cyan]\\[PW] Buscando por CBMERJ: {inscricao}-{dv}[/cyan]")
            page.fill("input[name='cbmerj']", inscricao)
            page.fill("input[name='cbmerj_dv']", dv)

            sucesso, erro = _submeter_busca_bombeiros(page, {
                "cbmerj": inscricao,
                "cbmerj_dv": dv,
            }, _log=_log)

            if sucesso:
                return True, ""

            # Cidade especial detectada: ja sabemos que CBMERJ nao funciona aqui,
            # vai direto pro fallback IPTU
            if erro == "CIDADE_ESPECIAL":
                if tem_iptu:
                    _msg(f"[yellow]\\[PW] Cidade especial (Macae/SG/Campos): CBMERJ nao funciona, indo pra IPTU...[/yellow]")
                    page.goto(URL_IFRAME_BOMBEIROS, wait_until="networkidle", timeout=60000)
                    page.wait_for_selector("input[name='cbmerj']", state="visible", timeout=15000)
                else:
                    return False, "Cidade especial exige IPTU, mas nao foi fornecido"
            # Nao encontrou - tenta inscricao predial
            elif tem_iptu:
                _msg(f"[yellow]\\[PW] CBMERJ nao encontrado, tentando por inscricao predial...[/yellow]")
                page.goto(URL_IFRAME_BOMBEIROS, wait_until="networkidle", timeout=60000)
                page.wait_for_selector("input[name='cbmerj']", state="visible", timeout=15000)
            else:
                return False, erro

        # Tentativa 2: por inscricao predial (IPTU)
        if tem_iptu:
            # Clica na aba "Insc Predial" para ativar o select de municipio
            page.evaluate("() => { const btn = document.querySelector('#btnIP'); if (btn) btn.click(); }")
            page.wait_for_timeout(500)
            cod_municipio = _resolver_codigo_municipio(page, cidade)
            _msg(f"[cyan]\\[PW] Buscando por inscricao predial: {nriptu_limpo} | Municipio: {cidade or 'Rio de Janeiro'} ({cod_municipio})[/cyan]")

            sucesso, erro = _submeter_busca_bombeiros(
                page,
                campos={
                    "inscricao": nriptu_limpo,
                    "municipio": cod_municipio,
                },
                destino="imovel/busca-inscricao-predial",
                endpoint="/sistema/imovel/busca-inscricao-predial",
                _log=_log,
            )

            if sucesso:
                return True, ""
            return False, erro

        return False, "Nenhuma busca teve sucesso"

    except Exception as e:
        return False, f"Erro na navegacao: {str(e)}"


# ============================================================
# Extracao da tabela de debitos
# ============================================================

def _extrair_debitos_bombeiros(page: Page, _log: list | None = None) -> list[dict]:
    """
    Parseia a tabela de vencimentos da pagina de resultado.
    Retorna lista de dicts com dados de cada exercicio com DEBITO.
    """
    _msg = lambda m: _log.append(m) if _log is not None else console.print(m)

    debitos = []
    # HTML pode estar em window._htmlBombeiros (via fetch) ou no page.content() (navegacao real)
    html = page.evaluate("() => window._htmlBombeiros || null") or page.content()
    soup = BeautifulSoup(html, "html.parser")

    # Procura tabela de vencimentos (busca flexivel)
    tabela = None
    for table in soup.find_all("table"):
        headers = [th.get_text(strip=True).lower() for th in table.find_all("th")]
        if "exercício" in headers or "exercicio" in headers or "vencimento" in headers:
            tabela = table
            break

    if not tabela:
        # Tenta encontrar por texto "Vencimentos" no HTML
        h_venc = soup.find(string=re.compile(r"Vencimentos", re.IGNORECASE))
        if h_venc:
            proximo = h_venc.find_next("table")
            if proximo:
                tabela = proximo

    if not tabela:
        _msg(f"[yellow]\\[AVISO] Nenhuma tabela de vencimentos encontrada[/yellow]")
        return debitos

    for linha in tabela.find_all("tr")[1:]:  # pula header
        colunas = linha.find_all("td")
        if len(colunas) < 6:
            continue

        situacao_el = colunas[5].find("span") or colunas[5]
        situacao = situacao_el.get_text(strip=True).upper()

        # Busca form de boleto na linha inteira
        form = linha.find("form")
        hidden_fields = {}
        if form:
            for inp in form.find_all("input", {"type": "hidden"}):
                nome = inp.get("name", "")
                valor = inp.get("value", "")
                if nome:
                    hidden_fields[nome] = valor

        exercicio_raw = colunas[0].get_text(strip=True)
        exercicio_limpo = re.sub(r"\D", "", exercicio_raw)  # Remove tudo que nao for digito
        exercicio_final = exercicio_limpo or exercicio_raw

        # Status temporal: Vencido se exercicio < ano_atual - 1, senao Imposto Ano Corrente
        # (regra do taxabombeiros.py legado, linha 148-151)
        try:
            exercicio_int = int(exercicio_limpo)
            ano_atual = date.today().year
            status_temporal = "Vencido" if exercicio_int < ano_atual - 1 else "Imposto Ano Corrente"
        except (ValueError, TypeError):
            status_temporal = "Indefinido"

        item = {
            "exercicio": exercicio_final,  # Fallback pro original se ficar vazio
            "vencimento": colunas[1].get_text(strip=True),
            "taxa": colunas[2].get_text(strip=True),
            "multa": colunas[3].get_text(strip=True),
            "situacao": situacao,
            "status_temporal": status_temporal,
            "form_fields": hidden_fields,
            "tem_boleto": bool(form),
        }
        # So inclui exercicios com debito (ignora QUITADO, DESCARTADO, etc.)
        if "DÉBITO" in situacao or "DEBITO" in situacao:
            debitos.append(item)

    com_debito = [d for d in debitos if "DÉBITO" in d["situacao"] or "DEBITO" in d["situacao"]]
    return com_debito


# ============================================================
# Download do boleto (via fetch interceptado)
# ============================================================

def _baixar_boleto_bombeiros(page: Page, form_fields: dict,
                             _log: list | None = None) -> bytes | None:
    """
    Gera o boleto via Playwright interceptando a resposta do POST.
    _log: se fornecido, acumula mensagens em vez de imprimir.
    Retorna bytes do PDF ou None.
    """
    _msg = lambda m: _log.append(m) if _log is not None else console.print(m)
    url_boleto = "https://www.funesbom.rj.gov.br/sistema/pagamentos/gerar/boletoSemRegistro"

    try:
        # Monta form dinamico via JS e submete interceptando a resposta
        # Ao inves de clicar no botao (que abre nova aba), fazemos fetch
        fields_safe = {k: str(v).replace('"', '\\"') for k, v in form_fields.items()}
        fields_js = ", ".join([f'"{k}": "{v}"' for k, v in fields_safe.items()])

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

        if resultado and resultado.get("data"):
            pdf_bytes = base64.b64decode(resultado["data"])
            content_type = resultado.get("contentType", "")
            tamanho = resultado.get("size", len(pdf_bytes))
            tam_fmt = f"{tamanho / 1024:.1f} KB" if tamanho >= 1024 else f"{tamanho} B"
            if tamanho < 5000:
                _msg(f"[yellow]\\[AVISO] Boleto recebido: {tam_fmt}, tipo: {content_type} - tamanho suspeito, pode nao ser PDF valido[/yellow]")
            else:
                _msg(f"[green]\\[OK] Boleto recebido: {tam_fmt}, tipo: {content_type}[/green]")

            # Verifica se e realmente PDF
            if pdf_bytes[:4] == b"%PDF":
                return pdf_bytes
            else:
                # Pode ser HTML de erro
                try:
                    texto = pdf_bytes.decode("utf-8", errors="replace")[:500]
                    _msg(f"[yellow]\\[AVISO] Resposta nao e PDF: {texto}[/yellow]")
                except Exception:
                    pass
                return None
        else:
            erro = resultado.get("error", "desconhecido") if resultado else "sem resposta"
            _msg(f"[red]\\[ERRO] Falha ao baixar boleto: {erro}[/red]")
            return None

    except Exception as e:
        _msg(f"[red]\\[ERRO] Excecao ao baixar boleto: {str(e)}[/red]")
        return None


# ============================================================
# Funcao principal: extrai boleto de bombeiros
# ============================================================

def extrairbombeiros_hibrido(page: Page, objeto, linha, dataatual="",
                              cota_unica: bool = False,
                              respeitar_horario: bool = False):
    """
    Extrai boleto de bombeiros usando Playwright + AntiCaptcha reCAPTCHA v2.
    Acumula todas as mensagens em memoria e imprime de uma vez no final.

    Parametros:
        cota_unica: Se True, filtra debitos mantendo so um por exercicio (descarta parcelas).
        respeitar_horario: Se True (default), nao tenta extrair apos as 22h
            (site CBM bloqueia geracao de boletos nesse horario).

    Retorna: (dados_list, df) - dados_list: [codigo, nrcbm, ano, status], df: DataFrame ou None
    """
    # Backlog: acumula mensagens e imprime tudo junto no final
    _log: list[str] = []

    if dataatual == "":
        dataatual = aux.hora("America/Sao_Paulo", "DATA")

    cod = str(linha[Codigo]).strip()
    nrcbm_preview = str(linha[NrCBM]).strip()
    anoextracao_preview = str(date.today().year)

    # Regra: site CBM nao gera boletos apos as 22h
    if respeitar_horario:
        from datetime import time as _time_cls
        hora_atual = aux.hora("America/Sao_Paulo", "HORA")
        if hora_atual >= _time_cls(22, 0):
            _log.append(
                f"[yellow]\\[INFO] Site CBM nao gera boletos apos as 22h "
                f"(hora atual: {hora_atual.strftime('%H:%M')}). Pulando.[/yellow]"
            )
            _flush_log(_log)
            return [cod, nrcbm_preview, anoextracao_preview, "Fora do horario (apos 22h)"], None

    cod = str(linha[Codigo]).strip()
    nrcbm = str(linha[NrCBM]).strip()
    nrcbm_limpo = nrcbm.replace("-", "").replace(".", "").replace(" ", "")
    # Coluna IPTU (indice 2) - usada como fallback na busca
    nriptu = str(linha[NrIPTU_CBM]).strip() if len(linha) > NrIPTU_CBM else ""
    # Coluna cidade (indice 3) - nome da cidade para o select de municipio
    cidade = str(linha[NrCidade]).strip() if len(linha) > NrCidade else ""
    anoextracao = str(date.today().year)

    # Se cod vazio (teste avulso sem banco), usa o CBM limpo como identificador
    tem_cod_cliente = bool(cod)
    if not cod:
        cod = nrcbm_limpo

    _log.append(f"[bold]━━━ {cod} | CBM: {nrcbm} ━━━[/bold]")

    os.makedirs(objeto.pastadownload, exist_ok=True)

    # Verifica se ja existe algum arquivo desse cliente na pasta (skip por cliente)
    prefixo = f"{cod}_{nrcbm_limpo}_" if tem_cod_cliente else f"{nrcbm_limpo}_"
    arquivos_existentes = [
        f for f in os.listdir(objeto.pastadownload)
        if f.startswith(prefixo) and f.endswith(".pdf")
    ]
    if arquivos_existentes:
        _log.append(f"[yellow]\\[SKIP] Cliente {cod} ja tem {len(arquivos_existentes)} arquivo(s) na pasta - pulando[/yellow]")
        _flush_log(_log)
        return [cod, nrcbm, anoextracao, f"SKIP ({len(arquivos_existentes)} arquivo(s) existente(s))"], None

    # Navegacao (tenta CBMERJ, fallback pra inscricao predial/IPTU)
    sucesso, erro = _navegar_bombeiros(page, nrcbm, nriptu=nriptu, cidade=cidade, _log=_log)
    if not sucesso:
        _log.append(f"[red]\\[ERRO] {erro}[/red]")
        _flush_log(_log)
        return [cod, nrcbm, anoextracao, erro], None

    # Extrai debitos
    debitos = _extrair_debitos_bombeiros(page, _log=_log)
    if not debitos:
        _log.append(f"[yellow]\\[INFO] Sem debitos[/yellow]")
        _flush_log(_log)
        return [cod, nrcbm, anoextracao, "Sem debitos"], None

    _log.append(f"[cyan]\\[PW] {len(debitos)} debito(s) encontrado(s)[/cyan]")
    for d in debitos:
        _log.append(f"[dim]  Exercicio {d['exercicio']} | Venc: {d['vencimento']} | "
                     f"Taxa: {d['taxa']} | Multa: {d['multa']} | Sit: {d['situacao']}[/dim]")

    # Detecta multi-parcela (varias linhas pro mesmo exercicio) e marca no log
    # Isso ajuda a identificar clientes pra refinar a regra de cota_unica
    from collections import Counter
    contagem_por_exercicio = Counter(d["exercicio"] for d in debitos)
    multi_parcela = {ex: n for ex, n in contagem_por_exercicio.items() if n > 1}
    if multi_parcela:
        for ex, n in multi_parcela.items():
            _log.append(
                f"[bold magenta]\\[!!! MULTI-PARCELA] Cliente {cod} tem {n} "
                f"linhas pro exercicio {ex} - analisar pra refinar cota_unica[/bold magenta]"
            )

    # Regra: cota unica — agrupa por exercicio, mantem so o primeiro de cada
    # (filtra parcelas; cada exercicio gera 1 boleto so)
    if cota_unica:
        vistos = set()
        debitos_filtrados = []
        for d in debitos:
            if d["exercicio"] not in vistos:
                vistos.add(d["exercicio"])
                debitos_filtrados.append(d)
        if len(debitos) != len(debitos_filtrados):
            _log.append(
                f"[cyan]\\[CotaUnica] {len(debitos)} debito(s) -> "
                f"{len(debitos_filtrados)} apos filtrar parcelas[/cyan]"
            )
        debitos = debitos_filtrados

    # Baixa boleto de cada exercicio com debito
    df_total = None
    baixados = 0
    for debito in debitos:
        ano = debito["exercicio"]
        # Nome: {cod}_{cbm}_{ano} (lote) ou {cbm}_{ano} (teste avulso)
        if tem_cod_cliente:
            nome_arquivo = f"{cod}_{nrcbm_limpo}_{ano}.pdf"
        else:
            nome_arquivo = f"{nrcbm_limpo}_{ano}.pdf"
        caminhodestino = os.path.join(objeto.pastadownload, nome_arquivo)

        if os.path.isfile(caminhodestino):
            _log.append(f"[yellow]\\[SKIP] Ja existe: {os.path.basename(caminhodestino)}[/yellow]")
            baixados += 1
            continue

        if not debito["form_fields"]:
            _log.append(f"[yellow]\\[AVISO] Exercicio {ano} sem formulario de boleto[/yellow]")
            continue

        pdf_bytes = _baixar_boleto_bombeiros(page, debito["form_fields"], _log=_log)
        if pdf_bytes:
            with open(caminhodestino, "wb") as f:
                f.write(pdf_bytes)
            tamanho_kb = len(pdf_bytes) / 1024

            # Adiciona cabecalho (mesmo padrao do fluxo_pw_new: origem = destino)
            try:
                aux.adicionarcabecalhopdf_topo_adaptativo(caminhodestino, caminhodestino, cod)
                _log.append(f"[green]\\[OK] Boleto {ano}: {os.path.basename(caminhodestino)} "
                            f"({tamanho_kb:.1f} KB) - cabecalho: {cod}[/green]")
            except Exception as e:
                _log.append(f"[green]\\[OK] Boleto {ano}: {os.path.basename(caminhodestino)} "
                            f"({tamanho_kb:.1f} KB)[/green] [yellow](sem cabecalho: {e})[/yellow]")

            baixados += 1

            # Extrai dados do PDF
            try:
                df = aux.extrairtextopdf(caminhodestino, "BOLETOS")
                if df is not None and not df.empty:
                    total_rows = df[df.columns[0]].count()
                    df.insert(loc=0, column="Codigo", value=["'" + cod + "'" for _ in range(total_rows)])
                    df.insert(loc=4, column="TpoPagto", value=["'PARCELADO'" for _ in range(total_rows)])
                    df.insert(loc=5, column="Arquivo", value=["'" + caminhodestino + "'" for _ in range(total_rows)])
                    df_total = pd.concat([df_total, df]) if df_total is not None else df
            except Exception as e:
                _log.append(f"[yellow]\\[AVISO] Erro ao extrair dados do PDF {ano}: {str(e)}[/yellow]")
        else:
            _log.append(f"[red]\\[ERRO] Falha ao baixar boleto {ano}[/red]")

    if baixados > 0:
        _log.append(f"[bold green]\\[RESUMO] {baixados}/{len(debitos)} boleto(s) baixado(s)[/bold green]")
        _flush_log(_log)
        return [cod, nrcbm, anoextracao, f"OK ({baixados} boleto(s))"], df_total

    _log.append(f"[bold red]\\[RESUMO] Nenhum boleto baixado[/bold red]")
    _flush_log(_log)
    return [cod, nrcbm, anoextracao, "Nao conseguiu baixar boleto"], None


# ============================================================
# Teste avulso
# ============================================================

def testar_bombeiros(inscricao: str = "3605693-5", iptu: str = "",
                     cidade: str = "", cod: str = "",
                     pasta_download: str = "Downloads/Bombeiros", headless: bool = False):
    """Testa a extracao de Bombeiros chamando extrairbombeiros_hibrido."""
    from playwright.sync_api import sync_playwright
    from types import SimpleNamespace

    # Perfil Chrome dedicado pra CBM (mesma var de ambiente que web.py usava)
    caminho_perfil = os.path.join(
        aux.caminhoprojeto(), "Profile",
        os.getenv("NOMEPROFILECBM", "cbm_profile")
    )
    os.makedirs(caminho_perfil, exist_ok=True)

    # Garante que a pasta de download existe
    if not os.path.isabs(pasta_download):
        pasta_download = os.path.join(aux.caminhoprojeto(), pasta_download)
    os.makedirs(pasta_download, exist_ok=True)

    # Monta objeto e linha no mesmo formato que o lote usa
    # linha: [Codigo, CBM, IPTU, Cidade]
    objeto = SimpleNamespace(pastadownload=pasta_download)
    linha = [cod, inscricao, iptu, cidade]

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=caminho_perfil,
            headless=headless,
            accept_downloads=True,
            args=["--disable-blink-features=AutomationControlled", "--start-maximized"],
        )
        page = context.pages[0] if context.pages else context.new_page()

        try:
            dados, df = extrairbombeiros_hibrido(page, objeto, linha)
            print(f"\n[RESULTADO] {dados}")
            if df is not None:
                print(df.to_string(index=False))
        finally:
            context.close()


# ============================================================
# CLI standalone
# ============================================================

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Extrator Bombeiros - Hibrido (Playwright + AntiCaptcha)")
    parser.add_argument("-i", "--inscricao", default="",
                        help="Numero CBMERJ (formato: 1234567-8). Vazio = pula direto pra IPTU.")
    parser.add_argument("--iptu", default="",
                        help="Inscricao predial/IPTU (usado se CBMERJ falhar ou estiver vazio)")
    parser.add_argument("--cidade", default="",
                        help="Nome da cidade para busca por inscricao predial (default: Rio de Janeiro)")
    parser.add_argument("--cod", default="",
                        help="Codigo do cliente (usado no nome do arquivo e cabecalho do PDF)")
    parser.add_argument("-p", "--pasta", default="Downloads/Bombeiros",
                        help="Pasta para salvar o PDF")
    parser.add_argument("--visible", action="store_true",
                        help="Abrir janela do navegador (padrao: headless)")
    args = parser.parse_args()

    testar_bombeiros(
        inscricao=args.inscricao,
        iptu=args.iptu,
        cidade=args.cidade,
        cod=args.cod,
        pasta_download=args.pasta,
        headless=not args.visible,
    )
