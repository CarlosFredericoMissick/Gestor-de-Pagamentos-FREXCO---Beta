// ========================================
// CONCILIAÇÃO — recebível, tabela de conciliação, filtro de times e detalhe
// ========================================

async function excluirRecebivel() {
    if (recebivelData.length === 0) { alert('A base de recebível já está vazia.'); return; }
    if (!confirm(`Excluir a base de recebível (${recebivelData.length} registros)? Esta ação não pode ser desfeita.`)) return;
    recebivelData = [];
    document.getElementById('recebivelStatus').textContent = '';
    if (useSharePoint) {
        try {
            await deleteRecebivelSP();
        } catch(e) { console.error('Erro ao excluir recebível SP:', e); }
    }
    alert('✅ Base de recebível excluída.');
}

// ========================================
// IMPORTAR PLANILHA DE RECEBÍVEL
// Nova estrutura: Data(0)|ID vehicle(1)|Time(2)|Veículo(3)|Rota(4)|ID Motorista driver_id(5)|
//                 Valor Rotas(6)|Ajuste(7)|Adicional(8)|Descarga(9)|Suporte(10)|Total Motorista(11)|Comissão(12)|Total(13)
// ========================================

async function importarRecebivelExcel(file) {
    try {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array', cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }).slice(1).filter(r => r[0]);

        recebivelData = rows.map(r => {
            let data;
            if (r[0] instanceof Date) {
                data = r[0];
            } else if (typeof r[0] === 'number') {
                const epoch = new Date(Date.UTC(1899, 11, 30));
                data = new Date(epoch.getTime() + r[0] * 86400000);
            } else {
                data = new Date(r[0]);
            }
            data = new Date(data.getFullYear(), data.getMonth(), data.getDate());
            return {
                data,
                vehicleId:      parseInt(r[1]),
                time:           (r[2] || '').toString().trim(),
                nome:           (r[3] || '').toString().trim(),
                rota:           r[4] ? String(r[4]).trim() : '',
                driverId:       parseInt(r[5]),
                valorRotas:     parseFloat(r[6]) || 0,
                ajuste:         parseFloat(r[7]) || 0,
                adicional:      parseFloat(r[8])  || 0,
                descarga:       parseFloat(r[9])  || 0,
                suporte:        parseFloat(r[10]) || 0,
                totalMotorista: parseFloat(r[11]) || 0,
                comissao:       parseFloat(r[12]) || 0,
            };
        }).filter(r => !isNaN(r.driverId));

        // Salvar no SharePoint
        if (useSharePoint) {
            saveRecebivelSP().catch(e => console.error('Erro ao salvar recebível SP:', e));
        }

        alert(`✅ Recebível importado: ${recebivelData.length} registros`);
    } catch(e) {
        console.error(e);
        alert('Erro ao processar planilha de recebível.');
    }
}

let _concilSort = { field: null, dir: 1 }; // dir: 1=asc, -1=desc
let _concilRows = []; // cache das linhas calculadas para re-sort sem recalcular

function sortConcil(field) {
    if (_concilRows.length === 0) return;
    if (_concilSort.field === field) {
        _concilSort.dir *= -1;
    } else {
        _concilSort.field = field;
        _concilSort.dir = 1;
    }
    // Atualizar ícones
    ['responsavel','valorNota','diferenca'].forEach(f => {
        const el = document.getElementById(`concil-sort-${f}`);
        if (!el) return;
        if (f === field) {
            el.textContent = _concilSort.dir === 1 ? '↑' : '↓';
            el.className = 'text-gray-800';
        } else {
            el.textContent = '↕';
            el.className = 'text-gray-400';
        }
    });
    renderConcilRows();
}

function renderConcilRows() {
    const rows = [..._concilRows];
    const { field, dir } = _concilSort;
    if (field) {
        rows.sort((a, b) => {
            let va, vb;
            if (field === 'responsavel') {
                va = (a.respNome || '').toLowerCase();
                vb = (b.respNome || '').toLowerCase();
                return dir * va.localeCompare(vb, 'pt-BR');
            } else if (field === 'valorNota') {
                va = a.valorNota || 0;
                vb = b.valorNota || 0;
            } else if (field === 'diferenca') {
                va = a.diff !== null ? a.diff : -Infinity;
                vb = b.diff !== null ? b.diff : -Infinity;
            }
            return dir * (va - vb);
        });
    }
    const tbody = document.getElementById('concilBody');
    tbody.innerHTML = '';
    rows.forEach(r => tbody.appendChild(r.tr));
}

function toggleConciliacaoModal() {
    const modal = document.getElementById('conciliacaoModal');
    const isHidden = modal.classList.contains('hidden');
    modal.classList.toggle('hidden');
    if (isHidden) {
        popularConcilTimes();
        // Datas padrão: mês atual
        const hoje = new Date();
        const y = hoje.getFullYear(), m = String(hoje.getMonth()+1).padStart(2,'0');
        if (!document.getElementById('concilDataInicio').value)
            document.getElementById('concilDataInicio').value = `${y}-${m}-01`;
        if (!document.getElementById('concilDataFim').value)
            document.getElementById('concilDataFim').value = `${y}-${m}-${new Date(y, hoje.getMonth()+1, 0).getDate()}`;
        lucide.createIcons();
    }
}


// ========================================
// CONCILIAÇÃO — EXECUÇÃO
// Valor devido calculado por calcularConciliacaoNotas (regras.js); tabela, detalhe e
// exportações usam o mesmo resultado guardado em _concilRows[i].resultado.
// ========================================

function executarConciliacao() {
    const inicio = document.getElementById('concilDataInicio').value;
    const fim    = document.getElementById('concilDataFim').value;
    const tolerancia = parseFloat(document.getElementById('concilTolerancia').value) || 0;

    if (!inicio || !fim) { alert('Defina o período de conciliação.'); return; }
    if (recebivelData.length === 0) { alert('Importe a planilha de recebível primeiro.'); return; }
    if (extractedResults.length === 0) { alert('Nenhuma nota carregada. Conecte o SharePoint ou extraia notas primeiro.'); return; }

    const dataInicio = new Date(inicio + 'T00:00:00');
    const dataFim    = new Date(fim    + 'T23:59:59');
    const fmtMoney   = formatarMoeda;
    const fmtPct     = (v) => v === null ? '-' : `${v > 0 ? '+' : ''}${v.toFixed(1)}%`;

    // Calculado sobre todas as notas (independe do filtro de time) para a comissão de
    // cada time ir para a nota mais recente do líder, mesmo fora do filtro.
    const resultados = calcularConciliacaoNotas(extractedResults, {
        prestadores: prestadoresList, recebivel: recebivelData, dataInicio, dataFim
    });

    // Filtro de times via checkbox
    const notasFiltradas = _concilTimesAtivos.size > 0
        ? extractedResults.filter(r => getTimesNota(r).some(t => _concilTimesAtivos.has(t)))
        : [...extractedResults];

    let ok = 0, divergente = 0, semDados = 0;
    _concilRows = [];
    _concilSort = { field: null, dir: 1 };
    ['responsavel','valorNota','diferenca'].forEach(f => {
        const el = document.getElementById(`concil-sort-${f}`);
        if (el) { el.textContent = '↕'; el.className = 'text-gray-400'; }
    });

    notasFiltradas.forEach(nota => {
        const resultado = resultados.get(nota);
        const { motoristas, isLider, valorDevido, temDados } = resultado;

        const valorNota = nota.vLiquido || 0;
        const respNome  = nota.responsavel?.nome || nota.prestador || '-';

        let status, statusClass, diff = null, diffPct = null;
        if (motoristas.length === 0) {
            status = '❌ Sem motorista'; statusClass = 'bg-red-50 text-red-700'; semDados++;
        } else if (!temDados) {
            status = '❌ Sem dados'; statusClass = 'bg-red-50 text-red-700'; semDados++;
        } else {
            diff = valorNota - valorDevido;
            diffPct = valorDevido !== 0 ? (diff / valorDevido) * 100 : null;
            if (Math.abs(diff) <= tolerancia) {
                status = '✅ Ok'; statusClass = 'bg-green-50 text-green-700'; ok++;
            } else {
                status = '⚠️ Divergente'; statusClass = 'bg-yellow-50 text-yellow-700'; divergente++;
            }
        }

        const nomesMotoristas = motoristas.map(m => m.nome).join(', ') || '-';
        const diffColor = diff === null ? ''
            : Math.abs(diff) > tolerancia
                ? (diff > 0 ? 'text-red-600 font-semibold' : 'text-orange-600 font-semibold')
                : 'text-green-600';
        const liderBadge = isLider
            ? `<span class="inline-block w-2 h-2 rounded-full bg-purple-500 mr-1"></span>`
            : '';
        const dataEmissaoDisplay = nota.dataEmissao || '-';

        // Célula Valor Devido: clicável para o detalhe da mesma linha
        const rowIdx = _concilRows.length;
        const corLink = isLider
            ? 'text-purple-700 underline cursor-pointer hover:text-purple-900'
            : 'text-blue-700 underline cursor-pointer hover:text-blue-900';
        const valorDevidoCell = temDados
            ? `<td class="px-3 py-2 text-right font-mono ${corLink}" onclick="abrirConcilDetalheNota(${rowIdx})">
                  ${fmtMoney(valorDevido)}</td>`
            : `<td class="px-3 py-2 text-right text-gray-400">-</td>`;

        const tr = document.createElement('tr');
        tr.className = `hover:bg-gray-50 transition-colors${isLider ? ' bg-purple-50/30' : ''}`;
        tr.innerHTML = `
            <td class="px-3 py-2 truncate max-w-[130px]" title="${nota.arquivo}">${nota.arquivo}</td>
            <td class="px-3 py-2 text-gray-700 truncate max-w-[120px]">${liderBadge}${respNome}</td>
            <td class="px-3 py-2 text-gray-600 truncate max-w-[160px]" title="${nomesMotoristas}">${nomesMotoristas}</td>
            <td class="px-3 py-2 text-center text-gray-500 text-[11px]">${dataEmissaoDisplay}</td>
            <td class="px-3 py-2 text-right font-mono">${fmtMoney(valorNota)}</td>
            ${valorDevidoCell}
            <td class="px-3 py-2 text-right font-mono ${diffColor}">${diff !== null ? fmtMoney(diff) : '-'}</td>
            <td class="px-3 py-2 text-right ${diffColor}">${fmtPct(diffPct)}</td>
            <td class="px-3 py-2 text-center"><span class="px-2 py-0.5 rounded-full text-[10px] font-semibold ${statusClass}">${status}</span></td>
        `;
        _concilRows.push({ tr, respNome, valorNota, diff, resultado, periodo: { dataInicio, dataFim } });
    });

    // Renderizar tabela (com sort padrão = nenhum)
    renderConcilRows();

    // Atualizar cards
    document.getElementById('concilCardsArea').classList.remove('hidden');
    document.getElementById('concilTotalNotas').textContent = notasFiltradas.length;
    document.getElementById('concilOk').textContent = ok;
    document.getElementById('concilDivergente').textContent = divergente;
    document.getElementById('concilSemDados').textContent = semDados;
    document.getElementById('concilPlaceholder').classList.add('hidden');
    document.getElementById('concilTabela').classList.remove('hidden');
}

// Conciliação time filter
function toggleConcilTimeDropdown() {
    const dd = document.getElementById('concilTimeDropdown');
    dd.classList.toggle('hidden');
    if (!dd.classList.contains('hidden')) {
        document.addEventListener('click', closeConcilTimeOnOutside, { once: true });
    }
}
function closeConcilTimeOnOutside(e) {
    const w = document.getElementById('concilFiltroTimeWrapper');
    if (!w || !w.contains(e.target)) document.getElementById('concilTimeDropdown')?.classList.add('hidden');
}
function onConcilTimeTodosChange(checked) {
    document.querySelectorAll('#concilTimeCheckboxes input').forEach(cb => cb.checked = checked);
    _concilTimesAtivos = new Set();
    atualizarConcilTimeBtnLabel();
}
function onConcilTimeCbChange(time, checked) {
    if (checked) _concilTimesAtivos.add(time);
    else _concilTimesAtivos.delete(time);
    const allCbs = document.querySelectorAll('#concilTimeCheckboxes input');
    const allChecked = [...allCbs].every(cb => cb.checked);
    document.getElementById('concilTimeTodos').checked = allChecked;
    if (allChecked) _concilTimesAtivos.clear();
    atualizarConcilTimeBtnLabel();
}
function atualizarConcilTimeBtnLabel() {
    const lbl = document.getElementById('concilTimeBtnLabel');
    if (!lbl) return;
    if (_concilTimesAtivos.size === 0) lbl.textContent = 'Todos';
    else if (_concilTimesAtivos.size === 1) lbl.textContent = [..._concilTimesAtivos][0];
    else lbl.textContent = `${_concilTimesAtivos.size} times`;
}
function popularConcilTimes() {
    const container = document.getElementById('concilTimeCheckboxes');
    if (!container) return;
    const times = new Set();
    extractedResults.forEach(r => getTimesNota(r).forEach(t => { if(t) times.add(t); }));
    container.innerHTML = '';
    [...times].sort().forEach(t => {
        container.innerHTML += `<label class="flex items-center gap-2 px-3 py-1.5 hover:bg-gray-50 cursor-pointer text-xs text-gray-700">
            <input type="checkbox" checked class="w-3.5 h-3.5 accent-gray-600" onchange="onConcilTimeCbChange('${t}', this.checked)">
            ${t}
        </label>`;
    });
    document.getElementById('concilTimeTodos').checked = true;
    _concilTimesAtivos.clear();
    atualizarConcilTimeBtnLabel();
}

// ========================================
// DETALHE DO VALOR DEVIDO (motorista ou líder)
// Mostra exatamente as linhas do recebível somadas em calcularValorDevidoNota.
// ========================================

// { nome, isLider, lider, resultado, dataInicio, dataFim } da linha aberta (usado pelo export)
let _concilDetalheAtual = null;

function switchDetalheTab(tab) {
    const tabM  = document.getElementById('concilDetalheTabMotoristas');
    const tabC  = document.getElementById('concilDetalheTabComissao');
    const painM = document.getElementById('concilDetalhePainelMotoristas');
    const painC = document.getElementById('concilDetalhePainelComissao');
    if (tab === 'motoristas') {
        tabM.classList.add('border-gray-700','text-gray-800');
        tabM.classList.remove('border-transparent','text-gray-500');
        tabC.classList.add('border-transparent','text-gray-500');
        tabC.classList.remove('border-purple-500','text-purple-700');
        painM.classList.remove('hidden');
        painC.classList.add('hidden');
    } else {
        tabC.classList.add('border-purple-500','text-purple-700');
        tabC.classList.remove('border-transparent','text-gray-500');
        tabM.classList.add('border-transparent','text-gray-500');
        tabM.classList.remove('border-gray-700','text-gray-800');
        painC.classList.remove('hidden');
        painM.classList.add('hidden');
    }
}

function cardResumo(cor, rotulo, valor) {
    const estilos = {
        gray:   ['bg-gray-50',   'border-gray-200',   'text-gray-500',   'text-gray-800'],
        purple: ['bg-purple-50', 'border-purple-200', 'text-purple-600', 'text-purple-700'],
        blue:   ['bg-blue-50',   'border-blue-200',   'text-blue-600',   'text-blue-700'],
        orange: ['bg-orange-50', 'border-orange-200', 'text-orange-600', 'text-orange-700'],
        teal:   ['bg-teal-50',   'border-teal-200',   'text-teal-600',   'text-teal-700'],
        yellow: ['bg-yellow-50', 'border-yellow-200', 'text-yellow-600', 'text-yellow-700'],
        indigo: ['bg-indigo-50', 'border-indigo-200', 'text-indigo-600', 'text-indigo-700'],
    }[cor];
    return `
        <div class="${estilos[0]} rounded-xl p-2.5 border ${estilos[1]} text-center">
            <div class="text-[10px] ${estilos[2]} mb-1">${rotulo}</div>
            <div class="text-base font-bold ${estilos[3]}">${formatarMoeda(valor)}</div>
        </div>`;
}

function abrirConcilDetalheNota(rowIdx) {
    const linha = _concilRows[rowIdx];
    if (!linha) return;
    const { resultado, respNome } = linha;
    const { dataInicio, dataFim } = linha.periodo;
    const lider = resultado.isLider ? resultado.lideres[0] : null;

    _concilDetalheAtual = {
        nome: lider ? lider.nome : respNome,
        isLider: resultado.isLider,
        lider,
        resultado,
        dataInicio,
        dataFim
    };

    const ini = new Date(dataInicio.getFullYear(), dataInicio.getMonth(), dataInicio.getDate());
    const fim = new Date(dataFim.getFullYear(),    dataFim.getMonth(),    dataFim.getDate());
    const fmtMoney = formatarMoeda;
    const fmtDate  = d => (d instanceof Date ? d : new Date(d)).toLocaleDateString('pt-BR');
    const linhas = resultado.linhasMotoristas;

    document.getElementById('concilDetalheNome').textContent = _concilDetalheAtual.nome;
    if (resultado.isLider) {
        const times = resultado.lideres.map(l => l.timeVinculado).join(', ');
        document.getElementById('concilDetalheLiderBadge').classList.remove('hidden');
        document.getElementById('concilDetalhePeriodo').textContent =
            `${ini.toLocaleDateString('pt-BR')} — ${fim.toLocaleDateString('pt-BR')} · Time: ${times}`;
        document.getElementById('concilDetalheTabs').classList.remove('hidden');
        document.getElementById('concilDetalheResumo').className = 'grid grid-cols-7 gap-2 mb-4';
        switchDetalheTab('motoristas');
    } else {
        document.getElementById('concilDetalheLiderBadge').classList.add('hidden');
        document.getElementById('concilDetalhePeriodo').textContent =
            `${ini.toLocaleDateString('pt-BR')} — ${fim.toLocaleDateString('pt-BR')} · ${linhas.length} registro(s)`;
        // Tabs — ocultas para motorista normal
        document.getElementById('concilDetalheTabs').classList.add('hidden');
        document.getElementById('concilDetalheResumo').className = 'grid grid-cols-6 gap-2 mb-4';
        document.getElementById('concilDetalhePainelMotoristas').classList.remove('hidden');
        document.getElementById('concilDetalhePainelComissao').classList.add('hidden');
    }

    const soma = (campo) => linhas.reduce((s, r) => s + (r[campo] || 0), 0);
    document.getElementById('concilDetalheResumo').innerHTML =
        cardResumo('gray', 'Total', resultado.valorDevido) +
        (resultado.isLider ? cardResumo('purple', 'Comissão Time', resultado.comissao) : '') +
        cardResumo('blue',   'Val. Rotas', soma('valorRotas')) +
        cardResumo('orange', 'Ajuste',     soma('ajuste')) +
        cardResumo('teal',   'Adicional',  soma('adicional')) +
        cardResumo('yellow', 'Descarga',   soma('descarga')) +
        cardResumo('indigo', 'Suporte',    soma('suporte'));

    // Aba 1: Motoristas na nota
    const tbody = document.getElementById('concilDetalheBody');
    tbody.innerHTML = '';
    linhas.forEach(r => {
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-gray-50';
        tr.innerHTML = `
            <td class="px-3 py-2">${fmtDate(r.data)}</td>
            <td class="px-3 py-2 text-gray-600">${r.nome || '-'}</td>
            <td class="px-3 py-2 text-gray-500">${r.rota || '-'}</td>
            <td class="px-3 py-2 text-right font-mono">${fmtMoney(r.valorRotas)}</td>
            <td class="px-3 py-2 text-right font-mono ${r.ajuste < 0 ? 'text-red-600' : r.ajuste > 0 ? 'text-green-600' : 'text-gray-400'}">${fmtMoney(r.ajuste)}</td>
            <td class="px-3 py-2 text-right font-mono text-gray-500">${fmtMoney(r.adicional)}</td>
            <td class="px-3 py-2 text-right font-mono text-yellow-700">${fmtMoney(r.descarga||0)}</td>
            <td class="px-3 py-2 text-right font-mono text-gray-500">${fmtMoney(r.suporte)}</td>
            <td class="px-3 py-2 text-right font-bold font-mono text-gray-800">${fmtMoney(r.totalMotorista)}</td>`;
        tbody.appendChild(tr);
    });

    // Aba 2: Comissão do time
    const tbodyC = document.getElementById('concilComissaoBody');
    tbodyC.innerHTML = '';
    [...resultado.linhasComissao].sort((a,b) => a.data - b.data).forEach(r => {
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-purple-50/50';
        tr.innerHTML = `
            <td class="px-3 py-2 text-gray-700">${fmtDate(r.data)}</td>
            <td class="px-3 py-2 text-gray-700">${r.nome || '-'}</td>
            <td class="px-3 py-2 text-right font-bold font-mono text-purple-700">${fmtMoney(r.comissao || 0)}</td>`;
        tbodyC.appendChild(tr);
    });
    resultado.comissaoEmOutraNota.forEach(c => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td colspan="3" class="px-3 py-2 text-gray-500 italic">Comissão do time ${c.time} no período entra só na nota mais recente do líder: ${c.arquivo}.</td>`;
        tbodyC.appendChild(tr);
    });

    document.getElementById('concilDetalheModal').classList.remove('hidden');
    lucide.createIcons();
}

function fecharConcilDetalhe() {
    document.getElementById('concilDetalheModal').classList.add('hidden');
}
