// ========================================
// MODAL — MOTORISTAS NA NOTA
// Salvar grava a nota no SharePoint com motoristasManual = true; a partir daí o
// reprocessamento não sobrescreve mais a lista (ver reconciliarMotoristasDaNota em regras.js).
// ========================================

let _motoristasModalIdx = null;
let _motoristasModalSelecionados = []; // [{vehicleId, nome}]

function abrirMotoristaModal(notaIdx) {
    _motoristasModalIdx = notaIdx;
    const nota = extractedResults[notaIdx];
    document.getElementById('motoristasNotaArquivo').textContent = nota.arquivo || '';

    // Inicializa com os motoristas já associados (ou responsável como default)
    _motoristasModalSelecionados = nota.motoristasNaNota
        ? nota.motoristasNaNota.map(m => ({...m}))
        : (nota.responsavel ? [{ ...nota.responsavel }] : []);

    // Nota automática emitida por líder: o valor devido soma comissão do time + vinculados
    // do líder + lista. Tudo isso aparece selecionado para que a lista salva (manual, fonte
    // de verdade) mantenha o mesmo valor e o usuário veja o que remove.
    const liderCnpj = !nota.motoristasManual ? findPrestadorByCnpj(nota.cnpj) : null;
    if (liderCnpj?.tipo === 'lider') {
        _motoristasModalSelecionados = listaExplicitaDoLider(liderCnpj, _motoristasModalSelecionados);
    }

    document.getElementById('motoristasNotaBusca').value = '';
    renderMotoristasSelecionados();
    renderMotoristasLista('');
    document.getElementById('motoristasNotaModal').classList.remove('hidden');
    document.getElementById('motoristasNotaBusca').focus();
    lucide.createIcons();
}

function fecharMotoristaModal() {
    _motoristasModalIdx = null;
    _motoristasModalSelecionados = [];
    document.getElementById('motoristasNotaModal').classList.add('hidden');
}

function renderMotoristasSelecionados() {
    const container = document.getElementById('motoristasNotaSelecionados');
    if (_motoristasModalSelecionados.length === 0) {
        container.innerHTML = '<span class="text-xs text-gray-400 italic">Nenhum selecionado</span>';
        return;
    }
    container.innerHTML = _motoristasModalSelecionados.map((m, i) => `
        <span class="inline-flex items-center gap-1 bg-frexco/10 text-frexco-dark border border-frexco/30 rounded-full px-2 py-0.5 text-xs font-medium">
            ${m.nome}
            <button onclick="removerMotoristaSelecionado(${i})" class="text-frexco/60 hover:text-red-500 ml-0.5 leading-none">✕</button>
        </span>
    `).join('');
}

function removerMotoristaSelecionado(idx) {
    _motoristasModalSelecionados.splice(idx, 1);
    renderMotoristasSelecionados();
    filtrarMotoristasBusca();
}

function filtrarMotoristasBusca() {
    const q = document.getElementById('motoristasNotaBusca').value;
    renderMotoristasLista(q);
}

function renderMotoristasLista(query) {
    const container = document.getElementById('motoristasNotaLista');
    const q = (query || '').toLowerCase().trim();
    const selecionadosIds = new Set(_motoristasModalSelecionados.map(m => idMotorista(m.vehicleId)));

    const ativos = prestadoresList.filter(p => p.inUse !== false);
    const filtrados = ativos.filter(p =>
        (!q || (p.nome || '').toLowerCase().includes(q)) &&
        !selecionadosIds.has(idMotorista(p.vehicleId))
    ).slice(0, 50);

    if (filtrados.length === 0) {
        container.innerHTML = '<p class="text-xs text-gray-400 text-center py-4">Nenhum motorista encontrado</p>';
        return;
    }
    // argJs: IDs de líder/manual ("L3", "M5") precisam ir entre aspas no onclick
    container.innerHTML = filtrados.map(p => `
        <div onclick="adicionarMotoristaSelecionado(${argJs(p.vehicleId)}, ${argJs(p.nome || '')})"
            class="flex items-center gap-2 px-3 py-2 hover:bg-frexco/5 cursor-pointer transition-colors">
            <span class="font-medium text-gray-800 text-sm flex-1">${p.nome}</span>
            ${p.time ? `<span class="text-xs text-blue-600 bg-blue-50 px-1.5 py-0.5 rounded">${p.time}</span>` : ''}
        </div>
    `).join('');
}

function adicionarMotoristaSelecionado(vehicleId, nome) {
    if (_motoristasModalSelecionados.some(m => idMotorista(m.vehicleId) === idMotorista(vehicleId))) return;
    _motoristasModalSelecionados.push({ vehicleId, nome });
    document.getElementById('motoristasNotaBusca').value = '';
    renderMotoristasSelecionados();
    renderMotoristasLista('');
}

async function salvarMotoristasDaNota() {
    if (_motoristasModalIdx === null) return;
    const nota = extractedResults[_motoristasModalIdx];

    const CAMPOS = ['motoristasNaNota', 'motoristasManual', 'motoristasEditadoEm', 'motoristasEditadoPor',
                    'temContrato', 'temContratoStatus', 'nomeContrato', '_spFileName'];
    const anterior = Object.fromEntries(CAMPOS.map(c => [c, nota[c]]));

    nota.motoristasNaNota    = _motoristasModalSelecionados.map(m => ({...m}));
    nota.motoristasManual    = true;
    nota.motoristasEditadoEm = new Date().toISOString();
    nota.motoristasEditadoPor = spAccount?.name || spAccount?.username || null;

    // Verificar contrato para todos os motoristas (mesma regra da coluna Contrato)
    const contratoInfo = verificarContratoNota(nota);
    nota.temContrato       = contratoInfo.status !== 'nao';
    nota.temContratoStatus = contratoInfo.status;
    nota.nomeContrato      = contratoInfo.contrato;

    if (useSharePoint) {
        const btn = document.getElementById('btnSalvarMotoristasNota');
        if (btn) { btn.disabled = true; btn.textContent = 'Salvando...'; }
        const salvou = await persistirNotaEditada(nota, CAMPOS.filter(c => c !== '_spFileName'));
        if (btn) { btn.disabled = false; btn.textContent = 'Salvar'; }
        if (!salvou) {
            // Desfaz em memória para a tela não mostrar algo que os outros usuários não verão;
            // o modal continua aberto com a seleção para nova tentativa.
            CAMPOS.forEach(c => { if (anterior[c] === undefined) delete nota[c]; else nota[c] = anterior[c]; });
            return;
        }
    } else {
        await persistirNotaEditada(nota); // só avisa que a alteração é local
    }

    // Re-renderizar tabela
    const tbody = document.getElementById('resultsBody');
    tbody.innerHTML = '';
    extractedResults.forEach(r => renderRow(r));

    updateCoverageStats();
    fecharMotoristaModal();
}
