/**
 * itbl — Interactive Table Component
 *
 * Uso:
 *   <table class="itbl" id="meuId"
 *          data-itbl-search="#inputBusca"
 *          data-itbl-count="#spanCount"
 *          data-itbl-clear="#btnLimpar"
 *          data-itbl-label="registro|registros">
 *
 *   Cada <th> pode ter:
 *     data-col="0"                         → índice da coluna (obrigatório para sort/filter)
 *     data-filter-type="text|check|none"   → tipo de filtro
 *     data-filter-options='["A","B"]'      → opções para filtro check
 *     data-sort-type="text|number|date|datetime|none" → tipo de ordenação (auto-detecta se omitido)
 *
 *   Cada <td> pode ter:
 *     data-sort-value="..."     → valor para ordenação (prioridade sobre textContent)
 *     data-filter-value="..."   → valor para filtro     (prioridade sobre textContent)
 *
 *   Dentro do <th>:
 *     <span class="th-label">Nome <i class="bi bi-funnel col-icon"></i></span>
 *     <span class="resize-handle"></span>   (opcional)
 */
(function () {
    'use strict';

    document.querySelectorAll('table.itbl').forEach(initTable);

    function initTable(table) {
        var thead = table.querySelector('thead');
        var tbody = table.querySelector('tbody');
        if (!thead || !tbody) return;

        // Elementos externos (opcionais)
        var searchSel = table.getAttribute('data-itbl-search');
        var countSel  = table.getAttribute('data-itbl-count');
        var clearSel  = table.getAttribute('data-itbl-clear');
        var labelAttr = table.getAttribute('data-itbl-label') || 'registro|registros';
        var labels    = labelAttr.split('|');
        var labelSing = labels[0];
        var labelPlur = labels[1] || labels[0] + 's';

        var searchInput = searchSel ? document.querySelector(searchSel) : null;
        var countEl     = countSel  ? document.querySelector(countSel)  : null;
        var btnLimpar   = clearSel  ? document.querySelector(clearSel)  : null;

        // ═══ Estado ═══
        var colState = {};
        var openPopup = null;

        // ═══ Init colunas ═══
        thead.querySelectorAll('th[data-col]').forEach(function (th) {
            var colIdx     = parseInt(th.getAttribute('data-col'));
            var filterType = th.getAttribute('data-filter-type') || 'none';
            var sortType   = th.getAttribute('data-sort-type') || 'auto';

            colState[colIdx] = {
                sort: null,
                sortType: sortType,
                filterType: filterType,
                filterVal: null
            };

            // Gerar popup
            var label = th.querySelector('.th-label');
            if (!label) return;  // coluna sem interação

            var popup = document.createElement('div');
            popup.className = 'col-popup';
            popup.setAttribute('data-popup-col', colIdx);

            var html = '';

            // Ordenação (exceto se sort-type="none")
            if (sortType !== 'none') {
                html += '<div class="popup-item" data-action="sort-asc"><i class="bi bi-sort-down"></i> Ordenar crescente</div>';
                html += '<div class="popup-item" data-action="sort-desc"><i class="bi bi-sort-up"></i> Ordenar decrescente</div>';
                html += '<div class="popup-divider"></div>';
            }

            // Filtro
            if (filterType === 'text') {
                html += '<input type="text" class="popup-filter-input" placeholder="Filtrar..." data-action="filter-text">';
            } else if (filterType === 'check') {
                var opts = [];
                try { opts = JSON.parse(th.getAttribute('data-filter-options') || '[]'); } catch (e) {}
                if (opts.length === 0) {
                    // Auto-detectar opções a partir dos dados
                    opts = autoDetectOptions(tbody, colIdx);
                    // Guardar para referência
                    th.setAttribute('data-filter-options', JSON.stringify(opts));
                }
                html += '<div class="popup-check-list">';
                opts.forEach(function (opt) {
                    html += '<label class="popup-check-item"><input type="checkbox" value="' +
                        escHtml(opt) + '" checked data-action="filter-check"> ' + escHtml(opt) + '</label>';
                });
                html += '</div>';
            }

            html += '<div class="popup-divider"></div>';
            html += '<div class="popup-clear" data-action="clear"><i class="bi bi-x-circle"></i> Limpar</div>';

            popup.innerHTML = html;
            th.appendChild(popup);

            // Toggle popup
            label.addEventListener('click', function (e) {
                e.stopPropagation();
                if (openPopup === popup) {
                    closeAllPopups();
                } else {
                    closeAllPopups();
                    popup.classList.add('show');
                    openPopup = popup;
                    var txtIn = popup.querySelector('.popup-filter-input');
                    if (txtIn) setTimeout(function () { txtIn.focus(); }, 50);
                }
            });

            // Ações dentro do popup
            popup.addEventListener('click', function (e) {
                e.stopPropagation();
                var target = e.target.closest('[data-action]');
                if (!target) return;
                var action = target.getAttribute('data-action');

                if (action === 'sort-asc') { doSort(colIdx, true); closeAllPopups(); }
                else if (action === 'sort-desc') { doSort(colIdx, false); closeAllPopups(); }
                else if (action === 'clear') { clearColFilter(colIdx); closeAllPopups(); }
                else if (action === 'filter-check') { applyCheckFilter(colIdx, popup); }
            });

            // Input de texto
            var txtInput = popup.querySelector('.popup-filter-input');
            if (txtInput) {
                txtInput.addEventListener('input', function () {
                    colState[colIdx].filterVal = this.value.trim().toLowerCase() || null;
                    updateLabelState(colIdx, th);
                    applyAllFilters();
                });
                txtInput.addEventListener('click', function (e) { e.stopPropagation(); });
                txtInput.addEventListener('keydown', function (e) {
                    if (e.key === 'Escape') closeAllPopups();
                });
            }
        });

        // Fechar popups ao clicar fora
        document.addEventListener('click', function () { closeAllPopups(); });

        function closeAllPopups() {
            table.querySelectorAll('.col-popup.show').forEach(function (p) { p.classList.remove('show'); });
            openPopup = null;
        }

        // ═══ SORT ═══
        function doSort(colIdx, asc) {
            // Reset sort de todas as colunas
            Object.keys(colState).forEach(function (k) { colState[k].sort = null; });
            colState[colIdx].sort = asc ? 'asc' : 'desc';

            var rows = Array.from(tbody.querySelectorAll('tr'));
            var st = colState[colIdx];

            rows.sort(function (a, b) {
                var aVal = getSortValue(a, colIdx);
                var bVal = getSortValue(b, colIdx);

                var cmp = compareValues(aVal, bVal, st.sortType);
                return asc ? cmp : -cmp;
            });

            rows.forEach(function (row) { tbody.appendChild(row); });

            // Atualizar labels e highlights
            thead.querySelectorAll('th[data-col]').forEach(function (th) {
                var ci = parseInt(th.getAttribute('data-col'));
                updateLabelState(ci, th);
                var pop = th.querySelector('.col-popup');
                if (pop) {
                    pop.querySelectorAll('.popup-item[data-action^="sort"]').forEach(function (item) {
                        item.classList.remove('active');
                    });
                    if (colState[ci].sort) {
                        var sel = pop.querySelector('.popup-item[data-action="sort-' + colState[ci].sort + '"]');
                        if (sel) sel.classList.add('active');
                    }
                }
            });
        }

        function getSortValue(row, colIdx) {
            var cell = row.children[colIdx];
            if (!cell) return '';
            return (cell.getAttribute('data-sort-value') ||
                    cell.getAttribute('data-filter-value') ||
                    cell.textContent).trim();
        }

        function compareValues(a, b, sortType) {
            // Auto-detect
            if (sortType === 'auto' || !sortType) {
                // Tentar número
                var aN = parseNum(a);
                var bN = parseNum(b);
                if (aN !== null && bN !== null) return aN - bN;

                // Tentar data/datetime
                var aD = parseDateTime(a);
                var bD = parseDateTime(b);
                if (aD && bD) return aD - bD;

                // Fallback: texto
                return a.toLowerCase().localeCompare(b.toLowerCase(), 'pt-BR');
            }

            if (sortType === 'number') {
                return (parseNum(a) || 0) - (parseNum(b) || 0);
            }
            if (sortType === 'date' || sortType === 'datetime') {
                var da = parseDateTime(a);
                var db = parseDateTime(b);
                return (da || 0) - (db || 0);
            }
            // text
            return a.toLowerCase().localeCompare(b.toLowerCase(), 'pt-BR');
        }

        function parseNum(s) {
            if (!s) return null;
            // Remover caracteres não numéricos exceto . , - e espaços
            var clean = s.replace(/[^\d.,-]/g, '').replace(',', '.');
            var n = parseFloat(clean);
            return isNaN(n) ? null : n;
        }

        function parseDateTime(s) {
            if (!s) return null;
            // dd/mm/yyyy HH:mm
            var m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})$/);
            if (m) return new Date(m[3], m[2] - 1, m[1], m[4], m[5]).getTime();

            // dd/mm/yyyy
            m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
            if (m) return new Date(m[3], m[2] - 1, m[1]).getTime();

            // dd/mm HH:mm (assume ano corrente)
            m = s.match(/^(\d{2})\/(\d{2})\s+(\d{2}):(\d{2})$/);
            if (m) return new Date(new Date().getFullYear(), m[2] - 1, m[1], m[3], m[4]).getTime();

            // yyyy-mm-dd (ISO sort value)
            m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
            if (m) return new Date(m[1], m[2] - 1, m[3]).getTime();

            return null;
        }

        // ═══ FILTROS ═══
        function getFilterValue(row, colIdx) {
            var cell = row.children[colIdx];
            if (!cell) return '';
            return (cell.getAttribute('data-filter-value') || cell.textContent).trim().toLowerCase();
        }

        function applyCheckFilter(colIdx, popup) {
            var checked = new Set();
            popup.querySelectorAll('.popup-check-item input:checked').forEach(function (cb) {
                checked.add(cb.value.toLowerCase());
            });
            var allChecks = popup.querySelectorAll('.popup-check-item input');
            colState[colIdx].filterVal = (checked.size === allChecks.length) ? null : checked;

            var th = popup.closest('th');
            updateLabelState(colIdx, th);
            applyAllFilters();
        }

        function clearColFilter(colIdx) {
            var state = colState[colIdx];
            state.filterVal = null;
            state.sort = null;

            var th = thead.querySelector('th[data-col="' + colIdx + '"]');
            if (!th) return;

            var txtInput = th.querySelector('.popup-filter-input');
            if (txtInput) txtInput.value = '';

            th.querySelectorAll('.popup-check-item input').forEach(function (cb) { cb.checked = true; });

            var popup = th.querySelector('.col-popup');
            if (popup) {
                popup.querySelectorAll('.popup-item[data-action^="sort"]').forEach(function (item) {
                    item.classList.remove('active');
                });
            }

            updateLabelState(colIdx, th);
            applyAllFilters();
        }

        function applyAllFilters() {
            var globalTerm = searchInput ? searchInput.value.toLowerCase().trim() : '';
            var hasAnyFilter = !!globalTerm;

            Object.keys(colState).forEach(function (k) {
                if (colState[k].filterVal !== null) hasAnyFilter = true;
            });

            if (btnLimpar) btnLimpar.style.display = hasAnyFilter ? '' : 'none';

            var visibleCount = 0;
            var totalCount = 0;

            tbody.querySelectorAll('tr').forEach(function (row) {
                totalCount++;
                var show = true;

                // Busca global
                if (globalTerm) {
                    var data = (row.getAttribute('data-search') || row.textContent).toLowerCase();
                    if (data.indexOf(globalTerm) === -1) show = false;
                }

                // Filtros de coluna
                if (show) {
                    for (var ci in colState) {
                        var st = colState[ci];
                        if (st.filterVal === null) continue;

                        var cellText = getFilterValue(row, parseInt(ci));

                        if (st.filterType === 'text') {
                            if (cellText.indexOf(st.filterVal) === -1) { show = false; break; }
                        } else if (st.filterType === 'check') {
                            if (!st.filterVal.has(cellText)) { show = false; break; }
                        }
                    }
                }

                row.style.display = show ? '' : 'none';
                if (show) visibleCount++;
            });

            updateCount(visibleCount, totalCount);
        }

        // ═══ UI helpers ═══
        function updateLabelState(colIdx, th) {
            var label = th.querySelector('.th-label');
            if (!label) return;
            var icon = label.querySelector('.col-icon');
            if (!icon) return;
            var st = colState[colIdx];

            var isActive = (st.sort !== null) || (st.filterVal !== null);
            label.classList.toggle('has-filter', isActive);

            if (st.sort === 'asc') {
                icon.className = 'bi bi-sort-down col-icon';
            } else if (st.sort === 'desc') {
                icon.className = 'bi bi-sort-up col-icon';
            } else if (st.filterVal !== null) {
                icon.className = 'bi bi-funnel-fill col-icon';
            } else {
                icon.className = 'bi bi-funnel col-icon';
            }
        }

        function updateCount(visible, total) {
            if (!countEl) return;
            if (visible === total) {
                countEl.textContent = total + ' ' + (total !== 1 ? labelPlur : labelSing);
            } else {
                countEl.textContent = visible + ' de ' + total + ' ' + (total !== 1 ? labelPlur : labelSing);
            }
        }

        // Auto-detectar opções únicas de uma coluna
        function autoDetectOptions(tbody, colIdx) {
            var seen = new Set();
            var ordered = [];
            tbody.querySelectorAll('tr').forEach(function (row) {
                var cell = row.children[colIdx];
                if (!cell) return;
                var val = (cell.getAttribute('data-filter-value') || cell.textContent).trim();
                if (val && !seen.has(val)) {
                    seen.add(val);
                    ordered.push(val);
                }
            });
            return ordered;
        }

        function escHtml(s) {
            var d = document.createElement('div');
            d.textContent = s;
            return d.innerHTML;
        }

        // ═══ Busca global ═══
        if (searchInput) {
            searchInput.addEventListener('input', applyAllFilters);
        }

        // ═══ Botão limpar ═══
        if (btnLimpar) {
            btnLimpar.addEventListener('click', function () {
                if (searchInput) searchInput.value = '';
                Object.keys(colState).forEach(function (ci) { clearColFilter(parseInt(ci)); });
                applyAllFilters();
            });
        }

        // ═══ Resize de colunas ═══
        var resizing = null;

        thead.querySelectorAll('.resize-handle').forEach(function (handle) {
            handle.addEventListener('mousedown', function (e) {
                e.preventDefault();
                e.stopPropagation();
                var th = this.parentElement;
                resizing = { th: th, startX: e.pageX, startWidth: th.offsetWidth };
                handle.classList.add('active');
                document.body.style.cursor = 'col-resize';
                document.body.style.userSelect = 'none';
            });
        });

        document.addEventListener('mousemove', function (e) {
            if (!resizing) return;
            var newWidth = resizing.startWidth + (e.pageX - resizing.startX);
            if (newWidth >= 50) resizing.th.style.width = newWidth + 'px';
        });

        document.addEventListener('mouseup', function () {
            if (!resizing) return;
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
            var handle = resizing.th.querySelector('.resize-handle');
            if (handle) handle.classList.remove('active');
            resizing = null;
        });

        // ═══ Init count ═══
        var totalRows = tbody.querySelectorAll('tr').length;
        updateCount(totalRows, totalRows);
    }
})();
