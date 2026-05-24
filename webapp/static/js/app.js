/**
 * app.js — Lógica do frontend do ExtratorUnificado Web.
 * Gerencia formulário de extração, opções dinâmicas e progresso via SSE.
 * Fontes de dados e templates são dinâmicos conforme empresa/fonte configurada.
 */

document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('formExtracao');
    const selectTipo = document.getElementById('tipo');
    const selectSubtipo = document.getElementById('subtipo');
    const empresaSelect = document.getElementById('empresaSelect');

    // Config vindo do template
    const config = window.EDX_CONFIG || { empresaId: 0, temSeletorEmpresa: false };

    // Divs condicionais
    const divSubtipo = document.getElementById('divSubtipo');
    const divPagamento = document.getElementById('divPagamento');
    const divOpcoesIPTU = document.getElementById('divOpcoesIPTU');
    const divCondominios = document.getElementById('divCondominios');
    const divDelay = document.getElementById('divDelay');
    const divProgresso = document.getElementById('divProgresso');

    const groupPagamento = document.getElementById('groupPagamento');
    const numWorkers = document.getElementById('numWorkers');
    const btnExecutar = document.getElementById('btnExecutar');
    const btnTemplate = document.getElementById('btnTemplate');
    const templateHint = document.getElementById('templateHint');
    const dataMes = document.getElementById('dataMes');

    // Estado das fontes carregadas
    let fontesDisponiveis = [];
    let empresaAtual = config.empresaId || 0;

    // ── Preenche data padrão (mm/aaaa) ───────────────────────────────────────
    const now = new Date();
    const mesAtual = String(now.getMonth() + 1).padStart(2, '0') + '/' + now.getFullYear();
    if (dataMes) dataMes.value = mesAtual;

    // ═══════════════════════════════════════════════════════════════════════════
    // CARREGAMENTO DINÂMICO DE FONTES
    // ═══════════════════════════════════════════════════════════════════════════

    // Se admin global: carregar empresas no dropdown
    console.log('[EDX] config:', config, 'empresaSelect:', !!empresaSelect);
    if (config.temSeletorEmpresa && empresaSelect) {
        carregarEmpresas();

        empresaSelect.addEventListener('change', () => {
            empresaAtual = parseInt(empresaSelect.value) || 0;
            if (empresaAtual) {
                selectTipo.disabled = false;
                carregarFontes(empresaAtual);
            } else {
                selectTipo.disabled = true;
                selectTipo.innerHTML = '<option value="">Selecione a empresa primeiro...</option>';
                fontesDisponiveis = [];
                atualizarTemplate();
            }
        });
    } else {
        // Usuário com empresa: carregar fontes direto
        if (config.empresaId) {
            carregarFontes(config.empresaId);
        } else {
            selectTipo.innerHTML = '<option value="">Nenhuma empresa vinculada</option>';
            selectTipo.disabled = true;
        }
    }

    async function carregarEmpresas() {
        try {
            const resp = await fetch('/api/empresas-disponiveis');
            const empresas = await resp.json();
            empresaSelect.innerHTML = '<option value="">Selecione a empresa...</option>';
            empresas.forEach(emp => {
                const opt = document.createElement('option');
                opt.value = emp.id;
                opt.textContent = emp.nome;
                empresaSelect.appendChild(opt);
            });
        } catch (err) {
            console.error('Erro ao carregar empresas:', err);
        }
    }

    async function carregarFontes(eid) {
        selectTipo.innerHTML = '<option value="">Carregando...</option>';
        selectTipo.disabled = true;

        try {
            const resp = await fetch(`/api/fontes-disponiveis?empresa_id=${eid}`);
            fontesDisponiveis = await resp.json();

            selectTipo.innerHTML = '<option value="">Selecione...</option>';

            // Filtrar _Faltante (variante interna do IPTU)
            const fontesVisiveis = fontesDisponiveis.filter(f => !f.tipo_extracao.endsWith('_Faltante'));

            if (fontesVisiveis.length === 0) {
                selectTipo.innerHTML = '<option value="">Nenhuma fonte configurada</option>';
                selectTipo.disabled = true;
            } else {
                // Agrupar: Prefeitura__* vira um único "Prefeitura", o resto fica individual
                const temPrefeitura = fontesVisiveis.some(f => f.tipo_extracao.startsWith('Prefeitura__'));
                const outrosTipos = fontesVisiveis.filter(f => !f.tipo_extracao.startsWith('Prefeitura__'));

                if (temPrefeitura) {
                    const opt = document.createElement('option');
                    opt.value = 'Prefeitura';
                    opt.textContent = 'Prefeitura';
                    selectTipo.appendChild(opt);
                }

                outrosTipos.forEach(f => {
                    const opt = document.createElement('option');
                    opt.value = f.tipo_extracao;
                    opt.textContent = f.nome_fonte;
                    opt.dataset.fonteId = f.id;
                    opt.dataset.colunas = JSON.stringify(f.colunas);
                    selectTipo.appendChild(opt);
                });
                selectTipo.disabled = false;
            }
        } catch (err) {
            console.error('Erro ao carregar fontes:', err);
            selectTipo.innerHTML = '<option value="">Erro ao carregar fontes</option>';
        }

        atualizarTemplate();
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // MUDANÇA DE TIPO DE EXTRAÇÃO
    // ═══════════════════════════════════════════════════════════════════════════

    selectTipo.addEventListener('change', () => {
        const tipo = selectTipo.value;

        // Reset opções condicionais
        divSubtipo.style.display = 'none';
        divPagamento.style.display = 'none';
        divOpcoesIPTU.style.display = 'none';
        divCondominios.style.display = 'none';
        if (divDelay) divDelay.style.display = 'none';
        if (numWorkers) { numWorkers.value = 1; numWorkers.disabled = false; }
        groupPagamento.innerHTML = '';

        if (tipo === 'Prefeitura') {
            // Mostrar combo de subtipos (IPTU, Nada Consta, Certidão Negativa)
            if (numWorkers) { numWorkers.value = 1; numWorkers.disabled = true; }
            divSubtipo.style.display = '';

            // Popular subtipo com fontes Prefeitura__ da empresa (exceto _Faltante)
            selectSubtipo.innerHTML = '<option value="">Selecione o serviço...</option>';
            const fontesPref = fontesDisponiveis.filter(f =>
                f.tipo_extracao.startsWith('Prefeitura__') && !f.tipo_extracao.endsWith('_Faltante')
            );
            fontesPref.forEach(f => {
                const opt = document.createElement('option');
                opt.value = f.tipo_extracao;
                opt.textContent = f.nome_fonte;
                opt.dataset.fonteId = f.id;
                opt.dataset.colunas = JSON.stringify(f.colunas);
                selectSubtipo.appendChild(opt);
            });
        } else if (tipo.startsWith('Condomin') || tipo === 'Condominios') {
            divCondominios.style.display = '';
        }
        // Bombeiros e outros: sem opções extras

        atualizarTemplate();
    });

    // ═══════════════════════════════════════════════════════════════════════════
    // MUDANÇA DE SUBTIPO (Prefeitura → IPTU / Nada Consta / Certidão)
    // ═══════════════════════════════════════════════════════════════════════════

    if (selectSubtipo) {
        selectSubtipo.addEventListener('change', () => {
            const sub = selectSubtipo.value;

            // Reset opções específicas
            divPagamento.style.display = 'none';
            divOpcoesIPTU.style.display = 'none';
            if (divDelay) divDelay.style.display = 'none';
            groupPagamento.innerHTML = '';

            if (sub.includes('IPTU')) {
                divPagamento.style.display = '';
                divOpcoesIPTU.style.display = '';
                if (divDelay) divDelay.style.display = '';
                criarRadiosPagamento([
                    { value: 1, label: 'Cota Única' },
                    { value: 2, label: 'Data 1', checked: true },
                    { value: 3, label: 'Data 2' },
                    { value: 4, label: 'Data 3' },
                ]);
            }

            atualizarTemplate();
        });
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // TEMPLATE DOWNLOAD
    // ═══════════════════════════════════════════════════════════════════════════

    // Retorna o tipo_extracao efetivo (subtipo quando Prefeitura)
    function getTipoEfetivo() {
        const tipo = selectTipo.value;
        if (tipo === 'Prefeitura' && selectSubtipo && selectSubtipo.value) {
            return selectSubtipo.value;  // ex: "Prefeitura__IPTU"
        }
        return tipo;
    }

    // Retorna o option da fonte selecionada (do subtipo ou do tipo)
    function getOptFonteSelecionada() {
        const tipo = selectTipo.value;
        if (tipo === 'Prefeitura' && selectSubtipo && selectSubtipo.selectedOptions[0]) {
            return selectSubtipo.selectedOptions[0];
        }
        return selectTipo.selectedOptions[0];
    }

    function atualizarTemplate() {
        const tipoEfetivo = getTipoEfetivo();
        if (!tipoEfetivo) {
            btnTemplate.classList.add('disabled');
            btnTemplate.removeAttribute('href');
            if (templateHint) {
                templateHint.innerHTML = '<i class="bi bi-info-circle"></i> Selecione o tipo de extração para habilitar o download do template.';
            }
            return;
        }

        // Buscar info da fonte selecionada
        const opt = getOptFonteSelecionada();
        const colunas = opt ? JSON.parse(opt.dataset.colunas || '[]') : [];

        if (colunas.length > 0) {
            btnTemplate.classList.remove('disabled');
            let url = `/api/template/${encodeURIComponent(tipoEfetivo)}`;
            if (empresaAtual) url += `?empresa_id=${empresaAtual}`;
            btnTemplate.setAttribute('href', url);

            if (templateHint) {
                templateHint.innerHTML = '<i class="bi bi-info-circle"></i> ' +
                    'Template com ' + colunas.length + ' coluna(s): <strong>' + colunas.join(', ') + '</strong>';
            }
        } else {
            btnTemplate.classList.add('disabled');
            btnTemplate.removeAttribute('href');
            if (templateHint) {
                templateHint.innerHTML = '<i class="bi bi-exclamation-triangle text-warning"></i> ' +
                    'Nenhuma coluna configurada para esta fonte. Configure no painel de Empresas.';
            }
        }
    }

    if (btnTemplate) {
        btnTemplate.addEventListener('click', (e) => {
            if (btnTemplate.classList.contains('disabled')) {
                e.preventDefault();
                return;
            }
            // href já está definido pelo atualizarTemplate()
        });
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // HELPERS DE UI
    // ═══════════════════════════════════════════════════════════════════════════

    function criarRadiosPagamento(opcoes) {
        groupPagamento.innerHTML = '';
        opcoes.forEach(op => {
            const input = document.createElement('input');
            input.type = 'radio';
            input.className = 'btn-check';
            input.name = 'tipopagamento';
            input.id = 'pag_' + op.value;
            input.value = op.value;
            input.autocomplete = 'off';
            if (op.checked) input.checked = true;

            const label = document.createElement('label');
            label.className = 'btn btn-outline-primary';
            label.htmlFor = 'pag_' + op.value;
            label.textContent = op.label;

            groupPagamento.appendChild(input);
            groupPagamento.appendChild(label);
        });
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // ABA ATIVA (upload vs banco)
    // ═══════════════════════════════════════════════════════════════════════════

    function getActiveTab() {
        const paneUpload = document.getElementById('pane-upload');
        if (paneUpload && paneUpload.classList.contains('show')) return 'upload';
        return 'banco';
    }

    function atualizarOpcoesBanco() {
        const isBanco = getActiveTab() === 'banco';
        // opcao-banco: esconde na aba upload (mas mantém se não tem abas)
        document.querySelectorAll('.opcao-banco').forEach(el => {
            el.style.display = isBanco ? '' : 'none';
            if (!isBanco) {
                const cb = el.querySelector('input[type="checkbox"]');
                if (cb) cb.checked = false;
            }
        });
        // opcao-somente-banco: só aparece quando na aba Banco
        document.querySelectorAll('.opcao-somente-banco').forEach(el => {
            el.style.display = isBanco ? '' : 'none';
            if (!isBanco) {
                const cb = el.querySelector('input[type="checkbox"]');
                if (cb) cb.checked = false;
            }
        });
    }

    document.querySelectorAll('#tabFonteDados button[data-bs-toggle="tab"]').forEach(btn => {
        btn.addEventListener('shown.bs.tab', atualizarOpcoesBanco);
    });

    atualizarOpcoesBanco();

    // ═══════════════════════════════════════════════════════════════════════════
    // SUBMIT DO FORMULÁRIO
    // ═══════════════════════════════════════════════════════════════════════════

    form.addEventListener('submit', async (e) => {
        e.preventDefault();

        const tipoEfetivo = getTipoEfetivo();
        if (!tipoEfetivo) {
            alert('Selecione o tipo de extração.');
            return;
        }

        const activeTab = getActiveTab();
        const arquivoDados = document.getElementById('arquivoDados');
        const arquivoBanco = document.getElementById('arquivoBanco');

        if (activeTab === 'upload') {
            if (!arquivoDados || !arquivoDados.files.length) {
                alert('Selecione um arquivo Excel ou CSV para upload.');
                return;
            }
        }

        btnExecutar.disabled = true;
        btnExecutar.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Iniciando...';

        let url, formData;

        if (activeTab === 'upload' && arquivoDados && arquivoDados.files.length) {
            url = '/api/upload-dados';
            formData = new FormData(form);
            formData.delete('arquivo_dados');
            formData.delete('arquivo_banco');
            formData.append('arquivo_dados', arquivoDados.files[0]);
            if (empresaAtual) formData.set('empresa_id', empresaAtual);
        } else {
            url = '/api/extrair';
            formData = new FormData(form);
            formData.delete('arquivo_dados');
            if (empresaAtual) formData.set('empresa_id', empresaAtual);
        }

        // Sobrescrever "tipo" com o valor efetivo (subtipo quando Prefeitura)
        formData.set('tipo', tipoEfetivo);

        try {
            const resp = await fetch(url, {
                method: 'POST',
                body: formData,
            });

            if (resp.status === 429) {
                const data = await resp.json();
                alert(data.erro || 'Limite diário atingido.');
                resetBotao();
                return;
            }

            if (resp.status === 400) {
                const data = await resp.json();
                alert(data.erro || 'Erro nos dados enviados.');
                resetBotao();
                return;
            }

            if (!resp.ok) {
                const text = await resp.text();
                alert('Erro ao iniciar extração: ' + (text || resp.status));
                resetBotao();
                return;
            }

            const data = await resp.json();

            if (data.registros) {
                const lblProgresso = document.getElementById('lblProgresso');
                if (lblProgresso) lblProgresso.textContent = data.mensagem || '';
            }

            iniciarProgresso(data.job_id);

        } catch (err) {
            alert('Erro de comunicação: ' + err.message);
            resetBotao();
        }
    });

    function resetBotao() {
        btnExecutar.disabled = false;
        btnExecutar.innerHTML = '<i class="bi bi-play-fill"></i> Executar Extração';
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // PROGRESSO VIA SSE
    // ═══════════════════════════════════════════════════════════════════════════

    function iniciarProgresso(jobId) {
        divProgresso.style.display = '';
        const barra = document.getElementById('barraProgresso');
        const lblProgresso = document.getElementById('lblProgresso');
        const lblTempo = document.getElementById('lblTempo');
        const lblETA = document.getElementById('lblETA');

        const evtSource = new EventSource(`/api/progresso/${jobId}`);

        evtSource.onmessage = (event) => {
            const d = JSON.parse(event.data);

            if (d.erro) {
                lblProgresso.textContent = d.erro;
                evtSource.close();
                resetBotao();
                return;
            }

            const pct = d.total > 0 ? Math.round((d.current / d.total) * 100) : 0;
            barra.style.width = pct + '%';
            barra.textContent = pct + '%';
            lblProgresso.textContent = d.mensagem || '';
            lblTempo.textContent = formatarTempo(d.elapsed || 0);
            lblETA.textContent = d.eta > 0 ? formatarTempo(d.eta) : '--:--';

            barra.classList.remove('bg-success', 'bg-danger');
            if (d.status === 'concluido') {
                barra.classList.remove('progress-bar-animated');
                barra.classList.add('bg-success');
                evtSource.close();
                resetBotao();
                setTimeout(() => location.reload(), 2000);
            } else if (d.status === 'erro') {
                barra.classList.remove('progress-bar-animated');
                barra.classList.add('bg-danger');
                evtSource.close();
                resetBotao();
            }
        };

        evtSource.onerror = () => {
            evtSource.close();
            resetBotao();
        };
    }

    function formatarTempo(totalSeg) {
        const h = Math.floor(totalSeg / 3600);
        const m = Math.floor((totalSeg % 3600) / 60);
        const s = Math.floor(totalSeg % 60);
        if (h > 0) return `${h}:${pad(m)}:${pad(s)}`;
        return `${pad(m)}:${pad(s)}`;
    }

    function pad(n) {
        return String(n).padStart(2, '0');
    }
});
