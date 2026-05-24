import os
import auxiliares as aux
from datetime import date, datetime
from playwright.sync_api import sync_playwright, Page
from dotenv import load_dotenv

load_dotenv()

# --- Configurações de Índices ---
Codigo, NrIPTU, NrCBM = 0, 1, 1


# --- 1. FUNÇÕES TÉCNICAS (Interação com o Site) ---

def extrairIPTU_pw(page: Page, objeto, linha, nrguia=0):
    """Extração técnica com trava de sincronia e tratamento de falha de requisição."""
    cod, nriptu = str(linha[Codigo]), str(linha[NrIPTU]).strip()
    gerar = not bool(objeto.visual.somentevalores.get())
    salvar_pdf = objeto.visual.codigosdebarra.get()
    tipo_pag = str(objeto.visual.tipopagamento.get())

    caminho = os.path.join(objeto.pastadownload, f"{cod}_{nriptu}.pdf")
    ERRO_SESSAO = 'Exception: Message TelaSelecao was received'

    try:
        # 1. LOOP DE RESILIÊNCIA
        for _ in range(3):
            # Trata a "FALHA NA REQUISIÇÃO" (Imagem 3)
            if page.get_by_text("FALHA NA REQUISIÇÃO").is_visible():
                page.get_by_role("link", name="clique aqui").click()
                page.wait_for_load_state("networkidle")

            if os.getenv('SITEIPTU') not in page.url:
                page.goto(os.getenv('SITEIPTU'), wait_until="networkidle", timeout=60000)

            page.wait_for_load_state("networkidle")

            input_insc = page.locator("#ctl00_ePortalContent_inscricao_input")
            input_insc.wait_for(state="visible")
            input_insc.fill(nriptu)

            # 🚀 TRAVA: Aguarda a navegação real após o clique em PROSSEGUIR (Imagem 1)
            try:
                with page.expect_navigation(wait_until="networkidle", timeout=30000):
                    page.locator("input[value='PROSSEGUIR']").click()
            except:
                pass

                # Verifica mensagens de erro na tela
            msg = page.locator("#ctl00_ePortalContent_MSG")
            if msg.is_visible(timeout=2000):
                txt = msg.inner_text().strip()
                if ERRO_SESSAO in txt:
                    page.goto(os.getenv('SITEIPTU'))
                    continue
                return [cod, nriptu, '', '', '', '', '', txt], None

            # Se a URL não mudou para a tela de dados (Imagem 2), tenta de novo
            if "TelaSelecao.aspx" in page.url and not page.get_by_text("DADOS DO IMÓVEL").is_visible():
                continue
            break

        # 2. TELA DE DADOS E EXTRAÇÃO (Imagem 2)
        page.wait_for_selector("text='DADOS DO IMÓVEL'", state="visible", timeout=15000)

        contribuinte = page.locator("#ctl00_ePortalContent_TELA_CONTRIBUINTE").inner_text().strip()
        exercicio = page.locator("#ctl00_ePortalContent_GuiaExercicio").inner_text().strip()
        dados_ret = [cod, nriptu, exercicio, 1, "Ver PDF", contribuinte, "", "Ok"]

        # Lógica de Download
        if gerar:
            guias = page.locator('[title*="visualizar / imprimir"]')
            if guias.count() > 0:
                guias.nth(nrguia).click()
                sel_id = "[id*='cbCotaUnica']" if tipo_pag == '1' else f"[id*='Chk_00{int(tipo_pag) - 1}']"
                page.locator(sel_id).check()

                with page.expect_download(timeout=60000) as dl_info:
                    page.locator("#ctl00_ePortalContent_btnDarmIndiv").click()
                    if page.locator("#popup_ok").is_visible(timeout=2000):
                        page.locator("#popup_ok").click()

                dl_info.value.save_as(caminho)
                aux.adicionarcabecalhopdf_topo_adaptativo(caminho, caminho, cod)

        # 🚀 3. RESET DE SINCRONIA: Usa o botão VOLTAR (ID ctl00_ePortalContent_Retornar)
        btn_voltar = page.locator("#ctl00_ePortalContent_Retornar")
        if btn_voltar.is_visible():
            with page.expect_navigation(wait_until="networkidle", timeout=30000):
                btn_voltar.click()
        else:
            page.goto(os.getenv('SITEIPTU'), wait_until="networkidle")

        return dados_ret, None

    except Exception as e:
        # Se falhar, tenta resetar a página para o próximo item
        try:
            page.goto(os.getenv('SITEIPTU'))
        except:
            pass
        return [cod, nriptu, '', '', '', '', '', f"Erro: {str(e)}"], None


def extrairnadaconsta_pw(page: Page, objeto, linha):
    """Versão Playwright para Nada Consta."""
    cod, nriptu = str(linha[Codigo]), str(linha[NrIPTU]).strip()
    caminho = os.path.join(objeto.pastadownload, f"{cod}_{nriptu}_nada_consta.pdf")

    try:
        page.goto(os.getenv('SITENADACONSTA'))
        page.locator("#Inscricao").fill(nriptu)
        page.locator("#Avancar").click()

        # Se houver link 'aqui', existe débito/guia
        link_download = page.locator("text='aqui'")
        if link_download.is_visible(timeout=3000):
            with page.expect_download() as dl:
                link_download.first.click()
            dl.value.save_as(caminho)
            aux.adicionarcabecalhopdf_topo_adaptativo(caminho, caminho, cod)
            return [cod, nriptu, "Débitos Encontrados", "Ok"], None

        return [cod, nriptu, "Nada Consta", "Ok"], None
    except Exception as e:
        return [cod, nriptu, "Erro", str(e)], None


def extraircertidaonegativa_pw(page: Page, objeto, linha):
    """Certidão Negativa com PDF nativo Playwright e Solver de Captcha."""
    cod, nriptu = str(linha[Codigo]), str(linha[NrIPTU]).strip()
    caminho = os.path.join(objeto.pastadownload, f"certidao_{cod}_{nriptu}.pdf")

    try:
        page.goto(os.getenv('SITECERTIDAOENFITEUTICA'))
        page.locator("[name='inscricao']").fill(nriptu)

        # Captura e resolve Captcha (usando sua lógica de solver)
        captcha_path = "temp_captcha.png"
        page.locator("#img").screenshot(path=captcha_path)
        # Aqui você chamaria o solver: texto = solver.solve(captcha_path)

        # Simulação de envio (ajuste para sua chave de solver)
        # page.locator("#texto_imagem").fill(texto)
        # page.locator("[name='btConsultar']").click()

        # Salva PDF da tela
        page.pdf(path=caminho, format="A4", print_background=True)
        return [cod, nriptu, "Certidão Gerada", "Ok"], None
    except Exception as e:
        return [cod, nriptu, "Erro", str(e)], None


# --- 2. CONTROLADOR PRINCIPAL (O Motor do Loop) ---

def controlaextracao_unificada(self):
    """Gerencia o navegador e o loop sincronizado para os 900+ itens."""
    self.listadados = []

    # Caminho do Perfil (Configurado no seu .env)
    caminho_perfil = os.path.join(aux.caminhoprojeto('Profile'), os.getenv('NOMEPROFILEIPTU'))

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=caminho_perfil,
            headless=False,
            accept_downloads=True,
            args=["--disable-blink-features=AutomationControlled", "--start-maximized"]
        )
        page = context.pages[0] if context.pages else context.new_page()

        for indice, linha in enumerate(self.resultado):
            cod = str(linha[Codigo])

            # --- ATUALIZAÇÃO DA GUI ---
            self.visual.mudartexto('labelcodigocliente', f'Código Cliente: {cod}')
            self.visual.configurarbarra('barraextracao', len(self.resultado), indice + 1)

            # 🚀 SINCRONIZAÇÃO: Força a janela a redesenhar ANTES de processar
            if hasattr(self.visual, 'update'):
                self.visual.update()

            try:
                servico = self.visual.tiposervico.get()
                if servico == 'Prefeitura' or servico == 'IPTU':
                    # A função agora 'trava' aqui até o site estar pronto
                    dados, df = extrairIPTU_pw(page, self, linha)

                    if df is not None:
                        # Padronização e Banco (Sua lógica original)
                        df = df.applymap(lambda x: x.replace("'", "") if isinstance(x, str) else x)
                        df['Valor Total'] = df['Valor Total'].str.replace('.', '').str.replace(',', '.')
                        df['Valor Total'] = df['Valor Total'].astype('float')
                        df['Parcela'] = df['Parcela'].str.replace("'", "").astype(int)
                        self.bd.adicionardf('[Codigos IPTUs]', df, 7)

                # ... outros serviços (Nada Consta, etc) ...

                if dados: self.listadados.append(dados)

            except Exception as e:
                print(f"Erro no item {cod}: {e}")
                # Pausa curta para o site respirar em caso de erro grave
                page.wait_for_timeout(2000)

        context.close()
    print("🏁 Processamento finalizado!")
