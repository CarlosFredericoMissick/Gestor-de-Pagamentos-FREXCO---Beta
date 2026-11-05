// ========================================
// TABELA DE NOTAS — ordenação, renderização das linhas e filtro por time
// ========================================

let currentSort = {
    field: null,
    direction: null // 'asc' ou 'desc'
};

function sortTable(field) {
    // Alternar direção ou iniciar nova ordenação
    if (currentSort.field === field) {
        // Já está ordenando por este campo, alternar direção
        if (currentSort.direction === 'asc') {
            currentSort.direction = 'desc';
        } else if (currentSort.direction === 'desc') {
            // Terceiro clique: remover ordenação
            currentSort.field = null;
            currentSort.direction = null;
        } else {
            currentSort.direction = 'asc';
        }
    } else {
        // Novo campo, começar com ascendente
        currentSort.field = field;
        currentSort.direction = 'asc';
    }

    // Ordenar array
    if (currentSort.field && currentSort.direction) {
        extractedResults.sort((a, b) => {
            let valueA, valueB;

            switch (field) {
                case 'responsavel':
                    valueA = (a.responsavel?.nome || a.prestador || '').toLowerCase();
                    valueB = (b.responsavel?.nome || b.prestador || '').toLowerCase();
                    break;
                case 'motoristas':
                    valueA = ((a.motoristasNaNota?.[0]?.nome) || (a.responsavel?.nome) || '').toLowerCase();
                    valueB = ((b.motoristasNaNota?.[0]?.nome) || (b.responsavel?.nome) || '').toLowerCase();
                    break;
                case 'arquivo':
                    valueA = (a.arquivo || '').toLowerCase();
                    valueB = (b.arquivo || '').toLowerCase();
                    break;
                case 'vliquido':
                    valueA = a.vLiquido || 0;
                    valueB = b.vLiquido || 0;
                    break;
                case 'vservico':
                    valueA = a.vServico || 0;
                    valueB = b.vServico || 0;
                    break;
                default:
                    return 0;
            }

            if (currentSort.direction === 'asc') {
                return valueA > valueB ? 1 : valueA < valueB ? -1 : 0;
            } else {
                return valueA < valueB ? 1 : valueA > valueB ? -1 : 0;
            }
        });
    }

    // Re-renderizar tabela
    const tbody = document.getElementById('resultsBody');
    tbody.innerHTML = '';
    extractedResults.forEach(result => renderRow(result));

    // Atualizar ícones
    updateSortIcons();
    lucide.createIcons();
}

function updateSortIcons() {
    // Resetar todos os ícones
    ['arquivo', 'responsavel', 'motoristas', 'vliquido', 'vservico'].forEach(field => {
        const icon = document.getElementById(`sort-icon-${field}`);
        if (icon) {
            icon.setAttribute('data-lucide', 'arrow-up-down');
        }
    });

    // Atualizar ícone do campo atual
    if (currentSort.field && currentSort.direction) {
        const icon = document.getElementById(`sort-icon-${currentSort.field}`);
        if (icon) {
            if (currentSort.direction === 'asc') {
                icon.setAttribute('data-lucide', 'arrow-up');
            } else {
                icon.setAttribute('data-lucide', 'arrow-down');
            }
        }
    }
}

function renderRow(row) {
    const tr = document.createElement('tr');
    tr.className = "hover:bg-primary/5 transition-colors border-b border-gray-100";
    const notaIdx = extractedResults.indexOf(row);
    tr.dataset.notaIdx = notaIdx;

    // Time da nota: derivado dos motoristas associados
    const times = new Set(
        (row.motoristasNaNota || (row.responsavel ? [row.responsavel] : []))
            .map(m => { const p = findPrestadorByVehicleId(m.vehicleId); return p?.time || ''; })
            .filter(Boolean)
    );
    row._times = [...times];

    const fmtMoney = (v) => v ? v.toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'}) : 'R$ 0,00';
    const nfColor = row.numeroNF === "?" ? "text-red-600 font-bold" : "text-gray-600";
    const modeloDisplay = row.method === "OCR" ? `OCR - ${row.modelo}` : row.modelo;
    const modeloColor = row.method === "OCR" ? "text-orange-600 font-semibold" : "text-primary font-medium";
    const ocrEditBtn = row.method === "OCR"
        ? `<button onclick="abrirOcrEdit(${notaIdx})" class="ml-1 p-0.5 text-orange-500 hover:text-orange-700 rounded transition-colors flex-shrink-0 inline-flex" title="Editar campos OCR"><svg xmlns="http://www.w3.org/2000/svg" class="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20h9"/><path d="M16.5 3.5a2.121 2.121 0 013 3L7 19l-4 1 1-4L16.5 3.5z"/></svg></button>`
        : '';

    const contratoStatus = row.temContratoStatus || (row.temContrato ? 'sim' : 'nao');
    const contratoIcon = contratoStatus === 'sim'
        ? `<span class="inline-flex items-center gap-1 px-2 py-1 bg-green-100 text-green-700 rounded-md text-xs font-semibold" title="${row.nomeContrato || 'Contrato encontrado'}"><i data-lucide="check-circle" class="w-3 h-3"></i> Sim</span>`
        : contratoStatus === 'parcial'
            ? `<span class="inline-flex items-center gap-1 px-2 py-1 bg-yellow-100 text-yellow-700 rounded-md text-xs font-semibold" title="Alguns motoristas sem contrato"><i data-lucide="alert-circle" class="w-3 h-3"></i> Parcial</span>`
            : `<span class="inline-flex items-center gap-1 px-2 py-1 bg-red-100 text-red-700 rounded-md text-xs font-semibold"><i data-lucide="x-circle" class="w-3 h-3"></i> Não</span>`;

    const responsavelNome = row.responsavel ? row.responsavel.nome : (row.prestador || '-');
    const responsavelCell = `<span class="font-medium text-gray-800 truncate block max-w-[130px]" title="${responsavelNome}">${responsavelNome}</span>`;

    const motoristas = row.motoristasNaNota || (row.responsavel ? [row.responsavel] : []);
    const rotulos = rotulosMotoristas(motoristas, prestadoresList); // "Jorge Comissão" / "Jorge Rotas"
    const motoristasChips = motoristas.length > 0
        ? rotulos.map(r => `<span class="inline-block bg-blue-50 text-blue-700 border border-blue-200 rounded px-1.5 py-0.5 text-[10px] font-medium truncate max-w-[90px]" title="${r}">${r}</span>`).join('')
        : `<span class="text-gray-400 text-xs italic">—</span>`;
    const motoristasCell = `
        <div class="flex items-center gap-1 flex-wrap">
            ${motoristasChips}
            <button onclick="abrirMotoristaModal(${notaIdx})" class="ml-1 p-0.5 text-gray-400 hover:text-frexco rounded transition-colors flex-shrink-0" title="Editar motoristas">
                <svg xmlns="http://www.w3.org/2000/svg" class="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
            </button>
        </div>`;

    tr.innerHTML = `
        <td class="px-3 py-2">
            <input type="checkbox" class="nota-checkbox w-4 h-4 accent-frexco cursor-pointer" data-idx="${notaIdx}" onchange="onNotaCheckChange()">
        </td>
        <td class="px-3 py-2 max-w-[130px]">
            <div class="flex items-center gap-1">
                <span class="truncate text-xs" title="${row.arquivo}">${row.arquivo}</span>${ocrEditBtn}
            </div>
        </td>
        <td class="px-3 py-2 max-w-[140px]">${responsavelCell}</td>
        <td class="px-3 py-2 max-w-[160px]">${motoristasCell}</td>
        <td class="px-3 py-2 truncate max-w-[160px]" title="${row.prestador}">${row.prestador}</td>
        <td class="px-3 py-2 text-[10px] font-mono text-gray-500">${row.cnpj || '-'}</td>
        <td class="px-3 py-2 text-gray-600">${fmtMoney(row.vLiquido)}</td>
        <td class="px-3 py-2 font-bold text-primary">${fmtMoney(row.vServico)}</td>
        <td class="px-3 py-2 text-gray-500">${row.dataEmissao || '-'}</td>
        <td class="px-3 py-2 ${nfColor}">${row.numeroNF}</td>
        <td class="px-3 py-2">${contratoIcon}</td>
        <td class="px-3 py-2 text-[10px] ${modeloColor}">${modeloDisplay}</td>
        <td class="px-3 py-2 text-xs text-gray-600 truncate max-w-[150px]" title="${row.chavePix || ''}">${row.chavePix || '-'}</td>
    `;
    document.getElementById('resultsBody').appendChild(tr);
    lucide.createIcons();
}

function toggleFiltroTimeDropdown() {
    const dd = document.getElementById('filtroTimeDropdown');
    dd.classList.toggle('hidden');
    if (!dd.classList.contains('hidden')) {
        document.addEventListener('click', closeFiltroTimeOnOutside, { once: true });
    }
}
function closeFiltroTimeOnOutside(e) {
    const wrapper = document.getElementById('filtroTimeWrapper');
    if (!wrapper || !wrapper.contains(e.target)) {
        document.getElementById('filtroTimeDropdown')?.classList.add('hidden');
    }
}
function onFiltroTimeTodosChange(checked) {
    document.querySelectorAll('#filtroTimeCheckboxes input[type=checkbox]').forEach(cb => cb.checked = checked);
    _filtroTimesAtivos = checked ? new Set() : new Set();
    if (!checked) {
        // nenhum selecionado = todos (padrão)
    }
    atualizarFiltroTimeBtnLabel();
    aplicarFiltroTimeExtrator();
}
function onFiltroTimeCbChange(time, checked) {
    if (checked) _filtroTimesAtivos.add(time);
    else _filtroTimesAtivos.delete(time);
    // se todos marcados → equivale a "Todos"
    const allCbs = document.querySelectorAll('#filtroTimeCheckboxes input[type=checkbox]');
    const allChecked = [...allCbs].every(cb => cb.checked);
    document.getElementById('filtroTimeTodos').checked = allChecked;
    if (allChecked) _filtroTimesAtivos.clear();
    atualizarFiltroTimeBtnLabel();
    aplicarFiltroTimeExtrator();
}
function atualizarFiltroTimeBtnLabel() {
    const lbl = document.getElementById('filtroTimeBtnLabel');
    if (!lbl) return;
    if (_filtroTimesAtivos.size === 0) lbl.textContent = 'Todos os times';
    else if (_filtroTimesAtivos.size === 1) lbl.textContent = [..._filtroTimesAtivos][0];
    else lbl.textContent = `${_filtroTimesAtivos.size} times`;
}
function aplicarFiltroTimeExtrator() {
    const tbody = document.getElementById('resultsBody');
    tbody.innerHTML = '';
    getNotasFiltradas().forEach(r => renderRow(r));
    updateCoverageStats();
    resetNotasSelecao();
}

function getNotasFiltradas() {
    if (_filtroTimesAtivos.size === 0) return extractedResults;
    return extractedResults.filter(r => getTimesNota(r).some(t => _filtroTimesAtivos.has(t)));
}

function atualizarOpcoesTime() {
    const container = document.getElementById('filtroTimeCheckboxes');
    if (!container) return;
    const times = new Set();
    extractedResults.forEach(r => getTimesNota(r).forEach(t => { if(t) times.add(t); }));
    container.innerHTML = '';
    [...times].sort().forEach(t => {
        const ativo = _filtroTimesAtivos.size === 0 || _filtroTimesAtivos.has(t);
        container.innerHTML += `<label class="flex items-center gap-2 px-3 py-1.5 hover:bg-gray-50 cursor-pointer text-xs text-gray-700">
            <input type="checkbox" ${ativo ? 'checked' : ''} class="w-3.5 h-3.5 accent-frexco" onchange="onFiltroTimeCbChange('${t}', this.checked)">
            ${t}
        </label>`;
    });
    const allCbs = [...container.querySelectorAll('input')];
    const allChecked = allCbs.every(cb => cb.checked);
    const todosEl = document.getElementById('filtroTimeTodos');
    if (todosEl) todosEl.checked = allChecked || allCbs.length === 0;
    atualizarFiltroTimeBtnLabel();
}

// Helper: obtém times de uma nota dinamicamente (sem depender de _times)
function getTimesNota(nota) {
    const motoristas = (nota.motoristasNaNota && nota.motoristasNaNota.length > 0)
        ? nota.motoristasNaNota
        : (nota.responsavel ? [nota.responsavel] : []);

    const times = new Set();
    motoristas.forEach(m => {
        const p = findPrestadorByVehicleId(m.vehicleId);
        if (!p) return;
        const t = p.tipo === 'lider' ? p.timeVinculado : p.time;
        if (t) times.add(t);
    });

    return [...times];
}

