"""
Extrator de boletos IPTU Rio de Janeiro via HTTP (requests + BeautifulSoup).
Substitui a navegação via Playwright por requisições HTTP diretas.
Elimina problemas de sincronismo do browser.

Fluxo:
1. GET TelaSelecao.aspx → pega ViewState + cookies de sessão
2. POST TelaSelecao.aspx com inscrição → página com dados do imóvel + guias
3. POST TelaSelecao.aspx com __EVENTTARGET da guia → DARM.aspx com cotas
4. POST DARM.aspx com checkboxes + botão GERAR → SegundaTela.aspx com PDF em base64
5. Extrai base64 do HTML e salva como PDF
"""

import os
import re
import base64
import requests
from bs4 import BeautifulSoup
from typing import Optional

# URLs
URL_BASE = "https://iportal.rio.rj.gov.br/PF331IPTUATUAL/pages/ParcelamentoIptuDs"
URL_TELA_SELECAO = f"{URL_BASE}/TelaSelecao.aspx"
URL_DARM = f"{URL_BASE}/DARM.aspx"
URL_SEGUNDA_TELA = f"{URL_BASE}/SegundaTela.aspx"

# Mapeamento tipo_pagamento → checkbox
# 1 = Cota Única, 2 = 1ª data, 3 = 2ª data, 4 = 3ª data
CHECKBOX_MAP = {
    1: "ctl00$ePortalContent$cbCotaUnica",
    2: "ctl00$ePortalContent$Chk_001",
    3: "ctl00$ePortalContent$Chk_002",
    4: "ctl00$ePortalContent$Chk_003",
}


def _extrair_campos_aspnet(soup: BeautifulSoup) -> dict:
    """Extrai __VIEWSTATE, __EVENTVALIDATION e outros campos hidden do ASP.NET."""
    campos = {}
    for nome in ["__VIEWSTATE", "__VIEWSTATEGENERATOR", "__VIEWSTATEENCRYPTED",
                  "__EVENTVALIDATION", "__EVENTTARGET", "__EVENTARGUMENT"]:
        tag = soup.find("input", {"name": nome})
        if tag:
            campos[nome] = tag.get("value", "")
    return campos


def _extrair_campos_hidden(soup: BeautifulSoup) -> dict:
    """Extrai todos os campos hidden do formulário."""
    campos = {}
    for inp in soup.find_all("input", {"type": "hidden"}):
        name = inp.get("name", "")
        if name:
            campos[name] = inp.get("value", "")
    return campos


def _extrair_pdf_base64(html: str) -> Optional[str]:
    """Extrai o PDF em base64 do HTML da SegundaTela.aspx."""
    # O PDF vem dentro de <object data="data:application/pdf;base64,XXXXX">
    match = re.search(r'data:application/pdf;base64,([A-Za-z0-9+/=\s]+)', html)
    if match:
        return match.group(1).replace('\n', '').replace('\r', '').replace(' ', '')
    return None


class ExtratorIPTU:
    """Extrai boletos de IPTU do portal da Prefeitura do Rio via HTTP."""

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        })

    def _get_inicial(self) -> BeautifulSoup:
        """Passo 1: GET na TelaSelecao para obter ViewState e cookies."""
        resp = self.session.get(URL_TELA_SELECAO)
        resp.raise_for_status()
        return BeautifulSoup(resp.text, "html.parser")

    def _post_inscricao(self, campos_aspnet: dict, inscricao: str) -> BeautifulSoup:
        """Passo 2: POST com a inscrição → retorna página com dados do imóvel."""
        data = {
            **campos_aspnet,
            "ctl00$ePortalContent$inscricao_input": inscricao,
            "ctl00$ePortalContent$DefiniGuia": "PROSSEGUIR",
            "ctl00$ePortalContent$EXERCICIO": "2026",
        }
        resp = self.session.post(URL_TELA_SELECAO, data=data)
        resp.raise_for_status()
        return BeautifulSoup(resp.text, "html.parser")

    def _selecionar_guia(self, soup_dados: BeautifulSoup, indice_guia: int = 0) -> BeautifulSoup:
        """Passo 3: Clica na guia (link) para ir para DARM.aspx.

        Args:
            soup_dados: HTML da página de dados do imóvel
            indice_guia: Índice da guia (0 = primeira guia com link)
        """
        campos = _extrair_campos_aspnet(soup_dados)

        # Encontra os links das guias (que têm __doPostBack)
        guia_links = soup_dados.find_all("a", href=re.compile(r"__doPostBack"))
        if not guia_links:
            raise ValueError("Nenhuma guia encontrada na página de dados do imóvel")

        # Filtra links que parecem ser de guias (contém título 'visualizar')
        guias_validas = []
        for link in guia_links:
            title = link.get("title", "")
            text = link.get_text(strip=True)
            # Guias geralmente têm formato XX/YYYY
            if re.match(r'\d{2}/\d{4}', text) or "visualizar" in title.lower():
                guias_validas.append(link)

        if not guias_validas:
            # Fallback: pega qualquer link com __doPostBack
            guias_validas = guia_links

        if indice_guia >= len(guias_validas):
            indice_guia = len(guias_validas) - 1

        guia = guias_validas[indice_guia]
        href = guia.get("href", "")

        # Extrai o __EVENTTARGET do javascript:__doPostBack('ALVO','')
        match = re.search(r"__doPostBack\('([^']+)'", href)
        if not match:
            raise ValueError(f"Não foi possível extrair __EVENTTARGET da guia: {href}")

        event_target = match.group(1)

        # Inclui todos os campos hidden do formulário
        all_hidden = _extrair_campos_hidden(soup_dados)
        data = {
            **all_hidden,
            "__EVENTTARGET": event_target,
            "__EVENTARGUMENT": "",
        }

        resp = self.session.post(URL_TELA_SELECAO, data=data)
        resp.raise_for_status()
        return BeautifulSoup(resp.text, "html.parser")

    def _gerar_boleto(self, soup_darm: BeautifulSoup, tipo_pagamento: int = 2) -> str:
        """Passo 4: Marca checkboxes e gera o boleto → retorna HTML com PDF.

        Args:
            soup_darm: HTML da página DARM com as cotas
            tipo_pagamento: 1=Cota Única, 2=1ª data, 3=2ª data, 4=3ª data
        """
        campos = _extrair_campos_aspnet(soup_darm)
        all_hidden = _extrair_campos_hidden(soup_darm)

        checkbox_name = CHECKBOX_MAP.get(tipo_pagamento)
        if not checkbox_name:
            raise ValueError(f"tipo_pagamento inválido: {tipo_pagamento}. Use 1-4.")

        # Monta o POST data
        data = {
            **all_hidden,
            "__EVENTTARGET": "",
            "__EVENTARGUMENT": "",
            checkbox_name: "on",
            "ctl00$ePortalContent$btnDarmIndiv": "GERAR BOLETO",
        }

        # Se for cota única, marca o checkbox correspondente
        # Se for parcelado (Chk_001/002/003), marca todos os checkboxes individuais da coluna
        if tipo_pagamento >= 2:
            col = tipo_pagamento - 1  # coluna 1, 2 ou 3
            for cota in range(1, 11):  # cotas 1 a 10
                cb_name = f"ctl00$ePortalContent${cota}-{col}"
                cb_tag = soup_darm.find("input", {"name": cb_name})
                if cb_tag:
                    data[cb_name] = "on"

        # O form action da DARM aponta para SegundaTela.aspx
        resp = self.session.post(URL_SEGUNDA_TELA, data=data)
        resp.raise_for_status()
        return resp.text

    def extrair_dados_imovel(self, soup: BeautifulSoup) -> dict:
        """Extrai informações do imóvel da página de dados."""
        dados = {}
        # Tenta pegar inscrição
        tag = soup.find(id="ctl00_ePortalContent_TELA_INSCRICAO")
        if tag:
            dados["inscricao"] = tag.get_text(strip=True)

        # Contribuinte
        tag = soup.find(id="ctl00_ePortalContent_TELA_CONTRIBUINTE")
        if tag:
            dados["contribuinte"] = tag.get_text(strip=True)

        # Guia/Exercício
        tag = soup.find(id="ctl00_ePortalContent_GuiaExercicio")
        if tag:
            dados["exercicio"] = tag.get_text(strip=True)

        # Mensagem de erro
        tag = soup.find(id="ctl00_ePortalContent_MSG")
        if tag:
            texto = tag.get_text(strip=True)
            if texto:
                dados["erro"] = texto

        return dados

    def extrair_datas_vencimento(self, soup_darm: BeautifulSoup) -> list:
        """Extrai as datas de vencimento disponíveis na tela DARM."""
        datas = []
        for th in soup_darm.find_all("th"):
            texto = th.get_text(strip=True)
            match = re.match(r'(\d{2}/\d{2}/\d{4})', texto)
            if match:
                datas.append(match.group(1))
        return datas

    def extrair_boleto(self, inscricao: str, tipo_pagamento: int = 2,
                       indice_guia: int = -1, pasta_download: str = "Downloads") -> dict:
        """Extrai um boleto de IPTU completo.

        Args:
            inscricao: Número de inscrição do IPTU (ex: "31469687" ou "3146968-7")
            tipo_pagamento: 1=Cota Única, 2=1ª data, 3=2ª data, 4=3ª data
            indice_guia: Índice da guia (-1 = última guia, geralmente a ativa)
            pasta_download: Pasta para salvar o PDF

        Returns:
            dict com resultado da extração
        """
        # Limpa inscrição (remove pontos e hífens)
        inscricao_limpa = inscricao.replace(".", "").replace("-", "")

        resultado = {
            "inscricao": inscricao,
            "status": "erro",
            "mensagem": "",
            "arquivo": "",
            "contribuinte": "",
            "exercicio": "",
        }

        try:
            # Passo 1: GET inicial
            soup_inicio = self._get_inicial()
            campos = _extrair_campos_aspnet(soup_inicio)

            # Passo 2: POST com inscrição
            soup_dados = self._post_inscricao(campos, inscricao_limpa)

            # Verifica erro
            dados_imovel = self.extrair_dados_imovel(soup_dados)
            if "erro" in dados_imovel:
                resultado["mensagem"] = dados_imovel["erro"]
                return resultado

            resultado["contribuinte"] = dados_imovel.get("contribuinte", "")

            # Verifica se tem guias disponíveis
            guia_links = soup_dados.find_all("a", href=re.compile(r"__doPostBack"))
            guias_validas = [l for l in guia_links if re.match(r'\d{2}/\d{4}', l.get_text(strip=True))]

            if not guias_validas:
                resultado["mensagem"] = "Nenhuma guia disponível"
                return resultado

            # Passo 3: Seleciona a guia
            # indice_guia=-1 pega a última (geralmente a ativa)
            soup_darm = self._selecionar_guia(soup_dados, indice_guia)

            # Extrai datas de vencimento disponíveis
            datas = self.extrair_datas_vencimento(soup_darm)
            resultado["datas_disponiveis"] = datas

            # Passo 4: Gera o boleto
            html_boleto = self._gerar_boleto(soup_darm, tipo_pagamento)

            # Passo 5: Extrai o PDF
            pdf_b64 = _extrair_pdf_base64(html_boleto)
            if not pdf_b64:
                resultado["mensagem"] = "PDF não encontrado na resposta"
                return resultado

            # Salva o PDF
            os.makedirs(pasta_download, exist_ok=True)
            nome_arquivo = f"{inscricao_limpa}.pdf"
            caminho = os.path.join(pasta_download, nome_arquivo)
            with open(caminho, "wb") as f:
                f.write(base64.b64decode(pdf_b64))

            resultado["status"] = "ok"
            resultado["mensagem"] = "Boleto extraído com sucesso"
            resultado["arquivo"] = caminho

        except Exception as e:
            resultado["mensagem"] = f"Erro: {str(e)}"

        return resultado


# ---- USO ----
if __name__ == "__main__":
    extrator = ExtratorIPTU()

    # Exemplo: extrair boleto para inscrição 3.146.968-7
    # tipo_pagamento: 1=Cota Única, 2=1ª data, 3=2ª data, 4=3ª data
    resultado = extrator.extrair_boleto(
        inscricao="3.146.968-7",
        tipo_pagamento=2,       # 1ª data de vencimento
        indice_guia=-1,         # última guia disponível
        pasta_download="Downloads"
    )

    print(f"Status: {resultado['status']}")
    print(f"Mensagem: {resultado['mensagem']}")
    if resultado.get("arquivo"):
        print(f"Arquivo: {resultado['arquivo']}")
    if resultado.get("contribuinte"):
        print(f"Contribuinte: {resultado['contribuinte']}")
    if resultado.get("datas_disponiveis"):
        print(f"Datas disponíveis: {resultado['datas_disponiveis']}")
