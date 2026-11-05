// ========================================
// LÍDERES — CRUD, sincronização com os times da base e contrato
// ========================================

function nextLiderId() {
    const existing = prestadoresList
        .map(p => String(p.vehicleId || ''))
        .filter(id => /^L\d+$/.test(id))
        .map(id => parseInt(id.slice(1)))
        .filter(n => !isNaN(n));
    const max = existing.length > 0 ? Math.max(...existing) : 0;
    return 'L' + (max + 1);
}

function getLideres() {
    return prestadoresList.filter(p => p.tipo === 'lider');
}

// Sincroniza líderes com os times presentes na base de motoristas
// Chamado após importarBaseMotoristasExcel
async function sincronizarLideresManual() {
    if (prestadoresList.filter(p => p.tipo !== 'lider').length === 0) {
        alert('Nenhuma base de motoristas carregada. Importe a base primeiro.');
        return;
    }
    const { criados, excluidos } = await sincronizarLideres();
    if (useSharePoint) {
        await savePrestadoresSP_internal();
        await loadPrestadoresFromSharePoint();
    }
    updatePrestadoresUI();
    renderLideresList();
    alert(`✅ Líderes sincronizados!\n\n${criados} criado(s)\n${excluidos} excluído(s)\n${getLideres().length} líderes ativos`);
    switchPrestadoresTab('lideres');
}

async function sincronizarLideres() {
    const times = new Set(
        prestadoresList
            .filter(p => p.tipo !== 'lider' && p.inUse !== false && p.time)
            .map(p => p.time.trim())
            .filter(Boolean)
    );

    let criados = 0, excluidos = 0;

    // Excluir líderes cujo time não existe mais na base
    const lideresAtuais = getLideres();
    for (const lider of lideresAtuais) {
        if (!times.has(lider.timeVinculado)) {
            const idx = prestadoresList.indexOf(lider);
            if (idx !== -1) prestadoresList.splice(idx, 1);
            excluidos++;
        }
    }

    // Criar líderes para times novos
    for (const time of times) {
        const jaExiste = prestadoresList.some(p => p.tipo === 'lider' && p.timeVinculado === time);
        if (!jaExiste) {
            const novoLider = {
                vehicleId: nextLiderId(),
                nome: time,
                tipo: 'lider',
                timeVinculado: time,
                motoristasVinculados: [],
                cd: '', time: '', inUse: true,
                cnpjs: [], contrato: null,
                fromSharePoint: useSharePoint,
                dataAdicao: new Date().toISOString()
            };
            prestadoresList.push(novoLider);
            criados++;
        }
    }

    if (criados > 0 || excluidos > 0) {
        console.log(`✅ Líderes: ${criados} criados, ${excluidos} excluídos`);
    }
    document.getElementById('lideresCount').textContent = getLideres().length;
    return { criados, excluidos };
}

// Render lista de líderes
function renderLideresList() {
    const lideres = getLideres();
    const query = (document.getElementById('lideresSearch')?.value || '').toLowerCase().trim();
    const container = document.getElementById('lideresListContainer');
    if (!container) return;
    document.getElementById('lideresCount').textContent = lideres.length;
    document.getElementById('lideresListCount').textContent = lideres.length;

    const visiveis = lideres.filter(l => {
        if (!query) return true;
        if ((l.nome || '').toLowerCase().includes(query)) return true;
        const qd = query.replace(/\D/g, '');
        if (qd && Array.isArray(l.cnpjs) && l.cnpjs.some(c => (c.cnpj||'').replace(/\D/g,'').includes(qd))) return true;
        return false;
    });

    if (visiveis.length === 0) {
        container.innerHTML = '<p class="text-sm text-gray-400 text-center py-10">Nenhum líder encontrado.</p>';
        return;
    }

    const fmtMoney = v => (v||0).toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
    container.innerHTML = visiveis.map(l => {
        const index = prestadoresList.indexOf(l);
        const cnpjHtml = (Array.isArray(l.cnpjs) && l.cnpjs.length > 0)
            ? l.cnpjs.map(c => `<div class="flex items-center gap-1"><span class="text-xs font-mono text-gray-600 bg-gray-100 px-1 py-0.5 rounded">${c.cnpj||'—'}</span>${c.chavePix ? `<span class="text-xs text-frexco" title="${c.chavePix}">→ ${c.chavePix}</span>` : ''}</div>`).join('')
            : `<span class="text-xs text-orange-500 bg-orange-50 px-1.5 py-0.5 rounded border border-orange-200">⚠ sem CNPJ</span>`;
        const contratoHtml = l.contrato
            ? `<div class="flex items-center gap-1">
                <span class="text-xs text-green-700 bg-green-50 px-2 py-0.5 rounded-full border border-green-200 truncate max-w-[100px]" title="${l.contrato.nomeArquivo}">✓ ${l.contrato.nomeArquivo}</span>
                <button onclick="viewContratoLider(${index})" class="p-1 text-blue-600 hover:bg-blue-50 rounded" title="Ver"><svg xmlns="http://www.w3.org/2000/svg" class="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg></button>
                <button onclick="removeContratoLider(${index})" class="p-1 text-red-400 hover:bg-red-50 rounded" title="Remover"><svg xmlns="http://www.w3.org/2000/svg" class="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>
               </div>`
            : `<button onclick="uploadContratoLider(${index})" class="flex items-center gap-1 text-xs text-blue-600 hover:bg-blue-50 px-2 py-1 rounded-lg border border-blue-200">↑ Anexar PDF</button>`;
        const motoristasHtml = (l.motoristasVinculados || []).length > 0
            ? l.motoristasVinculados.map(m => `<span class="inline-block bg-purple-50 text-purple-700 border border-purple-200 rounded px-1.5 py-0.5 text-[10px]">${m.nome}</span>`).join('')
            : `<span class="text-xs text-gray-400 italic">Nenhum vinculado</span>`;
        return `<div class="grid grid-cols-12 gap-2 items-start p-3 bg-purple-50 hover:bg-purple-100/60 rounded-xl border border-purple-200 transition-colors text-sm">
            <div class="col-span-3 min-w-0">
                <div class="flex items-center gap-1.5">
                    <span class="inline-block w-2 h-2 rounded-full bg-purple-500 flex-shrink-0"></span>
                    <p class="font-semibold text-gray-900 truncate">${l.nome}</p>
                </div>
                <span class="text-xs text-purple-600 bg-purple-100 px-1.5 py-0.5 rounded mt-0.5 inline-block">Líder · ${l.timeVinculado}</span>
                <span class="text-xs text-gray-400 block">#${l.vehicleId}</span>
            </div>
            <div class="col-span-3 min-w-0 space-y-0.5">${cnpjHtml}</div>
            <div class="col-span-2 min-w-0">${contratoHtml}</div>
            <div class="col-span-2 min-w-0 text-[10px] flex flex-wrap gap-1">${motoristasHtml}</div>
            <div class="col-span-2 flex items-start justify-end gap-1 pt-0.5">
                <button onclick="editLiderEntry(${index})" class="p-1.5 text-gray-500 hover:text-purple-600 hover:bg-purple-50 rounded-lg" title="Editar"><svg xmlns="http://www.w3.org/2000/svg" class="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg></button>
            </div>
        </div>`;
    }).join('');
}

// Edit líder form
let _liderEditIdx = null;
let _liderMotoristasTemp = [];

function editLiderEntry(index) {
    _liderEditIdx = index;
    const l = prestadoresList[index];
    _liderMotoristasTemp = [...(l.motoristasVinculados || [])];
    document.getElementById('liderNome').value = l.nome;
    document.getElementById('liderTime').value = l.timeVinculado;
    document.getElementById('liderFormTitle').textContent = `Editar Líder — ${l.nome}`;
    // CNPJs
    const container = document.getElementById('cnpjRowsLiderContainer');
    container.innerHTML = '';
    const cnpjs = (l.cnpjs || []).length > 0 ? l.cnpjs : [{ cnpj: '', chavePix: '' }];
    cnpjs.forEach(c => adicionarCnpjRowLiderData(c.cnpj, c.chavePix));
    renderLiderMotoristasVinculados();
    document.getElementById('liderForm').classList.remove('hidden');
    document.getElementById('liderForm').scrollIntoView({ behavior: 'smooth' });
}

function cancelLiderForm() {
    _liderEditIdx = null;
    _liderMotoristasTemp = [];
    document.getElementById('liderForm').classList.add('hidden');
}

function adicionarCnpjRowLiderData(cnpj = '', pix = '') {
    const container = document.getElementById('cnpjRowsLiderContainer');
    const idx = container.querySelectorAll('.cnpj-lider-row').length;
    const div = document.createElement('div');
    div.className = 'cnpj-lider-row flex gap-2 items-center';
    div.innerHTML = `<input type="text" placeholder="CNPJ" value="${cnpj}" class="cnpj-lider-input flex-1 px-2 py-1.5 border border-gray-300 rounded-lg text-xs focus:ring-1 focus:ring-purple-400 outline-none font-mono">
        <input type="text" placeholder="Chave PIX" value="${pix}" class="pix-lider-input flex-1 px-2 py-1.5 border border-gray-300 rounded-lg text-xs focus:ring-1 focus:ring-purple-400 outline-none">
        <button onclick="this.closest('.cnpj-lider-row').remove()" class="text-red-400 hover:text-red-600 text-xs px-2 py-1.5 rounded flex-shrink-0">✕</button>`;
    container.appendChild(div);
}

function addCnpjRowLider() { adicionarCnpjRowLiderData('', ''); }

function renderLiderMotoristasVinculados() {
    const container = document.getElementById('liderMotoristasVinculados');
    if (!container) return;
    container.innerHTML = _liderMotoristasTemp.map((m, i) =>
        `<span class="inline-flex items-center gap-1 bg-purple-100 text-purple-800 border border-purple-200 rounded-full px-2 py-0.5 text-xs font-medium">
            ${m.nome}
            <button onclick="removerMotoristaLider(${i})" class="text-purple-400 hover:text-purple-700 leading-none ml-0.5">✕</button>
        </span>`
    ).join('');
}

function removerMotoristaLider(idx) {
    _liderMotoristasTemp.splice(idx, 1);
    renderLiderMotoristasVinculados();
}

function filtrarMotoristasLider(query) {
    const sugestoes = document.getElementById('liderMotoristaSugestoes');
    if (!query.trim()) { sugestoes.classList.add('hidden'); return; }
    const motoristas = prestadoresList.filter(p =>
        p.tipo !== 'lider' && p.inUse !== false &&
        (p.nome || '').toLowerCase().includes(query.toLowerCase()) &&
        !_liderMotoristasTemp.some(m => m.vehicleId === p.vehicleId)
    ).slice(0, 8);
    if (motoristas.length === 0) { sugestoes.classList.add('hidden'); return; }
    sugestoes.innerHTML = motoristas.map(m =>
        `<div onclick="adicionarMotoristaLider('${m.vehicleId}', '${(m.nome||'').replace(/'/g,"\\'")}', this)"
            class="px-3 py-2 hover:bg-purple-50 cursor-pointer text-sm text-gray-700 flex items-center gap-2">
            <span class="font-medium">${m.nome}</span>
            ${m.time ? `<span class="text-xs text-blue-600 bg-blue-50 px-1.5 rounded">${m.time}</span>` : ''}
        </div>`
    ).join('');
    sugestoes.classList.remove('hidden');
}

function adicionarMotoristaLider(vehicleId, nome) {
    _liderMotoristasTemp.push({ vehicleId, nome });
    document.getElementById('liderMotoristaBusca').value = '';
    document.getElementById('liderMotoristaSugestoes').classList.add('hidden');
    renderLiderMotoristasVinculados();
}

async function saveLiderEntry() {
    if (_liderEditIdx === null) return;
    const l = prestadoresList[_liderEditIdx];

    // Coletar CNPJs
    const cnpjRows = document.querySelectorAll('#cnpjRowsLiderContainer .cnpj-lider-row');
    const novosCnpjs = Array.from(cnpjRows).map(r => ({
        cnpj: r.querySelector('.cnpj-lider-input').value.trim(),
        chavePix: r.querySelector('.pix-lider-input').value.trim()
    })).filter(c => c.cnpj);

    // Validar unicidade de CNPJ
    const trocas = [];
    for (const nc of novosCnpjs) {
        const outro = prestadoresList.find(p => p !== l && Array.isArray(p.cnpjs) && p.cnpjs.some(c => c.cnpj === nc.cnpj));
        if (outro) {
            outro.cnpjs = outro.cnpjs.filter(c => c.cnpj !== nc.cnpj);
            trocas.push(outro.nome);
        }
    }

    l.cnpjs = novosCnpjs;
    l.motoristasVinculados = [..._liderMotoristasTemp];

    if (useSharePoint) {
        await savePrestadoresSP_internal();
        await loadPrestadoresFromSharePoint();
    } else {
        await savePrestadorLocal(l);
    }
    reprocessExtractedResults();
    cancelLiderForm();
    renderLideresList();

    if (trocas.length > 0) {
        alert(`⚠️ CNPJ movido de: ${trocas.join(', ')}\n\nO(s) registro(s) acima perdeu(ram) o CNPJ que foi alocado para ${l.nome}.`);
    }
}

// Contrato do líder
function uploadContratoLider(index) {
    const input = document.getElementById('contratoLiderUploadInput');
    input.onchange = async (e) => {
        const file = e.target.files[0]; if (!file) return;
        const l = prestadoresList[index];
        const dataUrl = await new Promise(res => { const r = new FileReader(); r.onload = () => res(r.result); r.readAsDataURL(file); });
        l.contrato = { nomeArquivo: file.name, fileData: dataUrl, dataUpload: new Date().toISOString() };
        if (useSharePoint) { await savePrestadoresSP_internal(); await loadPrestadoresFromSharePoint(); }
        else await savePrestadorLocal(l);
        renderLideresList();
        input.value = '';
    };
    input.click();
}
function viewContratoLider(index) {
    const l = prestadoresList[index];
    if (!l?.contrato?.fileData) return;
    const w = window.open();
    w.document.write(`<iframe src="${l.contrato.fileData}" style="width:100%;height:100vh;border:none;"></iframe>`);
}
async function removeContratoLider(index) {
    const l = prestadoresList[index];
    if (!l?.contrato) return;
    if (!confirm('Remover contrato?')) return;
    l.contrato = null;
    if (useSharePoint) { await savePrestadoresSP_internal(); await loadPrestadoresFromSharePoint(); }
    else await savePrestadorLocal(l);
    renderLideresList();
}
