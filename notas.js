// ========================================
// NOTAS — limpeza, seleção/exclusão, cards de cobertura, reprocessamento e edição OCR
// ========================================

// Grava a nota editada no SharePoint e avisa o usuário se a alteração não foi compartilhada.
// campos: os campos que esta edição alterou (em conflito, só eles são reaplicados).
async function persistirNotaEditada(nota, campos) {
    if (!useSharePoint) {
        alert('⚠️ SharePoint não conectado: a alteração ficou só neste navegador e será perdida ao recarregar.\n\nConecte o SharePoint e salve novamente para compartilhar com os outros usuários.');
        return false;
    }
    const res = await saveNotaSP(nota, { campos });
    if (!res.ok) {
        alert(`❌ Não foi possível salvar a nota "${nota.arquivo}" no SharePoint (${res.mensagem}).\n\nA alteração NÃO foi compartilhada. Tente novamente; se persistir, reconecte o SharePoint.`);
        return false;
    }
    if (res.conflitos.length > 0) {
        alert(`⚠️ Outro usuário alterou a nota "${nota.arquivo}" enquanto você editava.\n\nCampos alterados pelos dois: ${res.conflitos.join(', ')} — ficou a sua versão.\nOs demais campos foram atualizados com a versão do SharePoint.`);
    }
    return true;
}

const CAMPOS_EDICAO_OCR = ['prestador', 'cnpj', 'numeroNF', 'vServico', 'vLiquido', 'dataEmissao', 'modelo',
                           'responsavel', 'chavePix', 'temContrato', 'temContratoStatus', 'nomeContrato'];


async function clearResults() {
    if (extractedResults.length === 0) return;
    if (!confirm(`Apagar ${extractedResults.length} nota(s)? Esta ação não pode ser desfeita.`)) return;

    // Deletar arquivos individuais no SharePoint
    if (useSharePoint) {
        for (const nota of extractedResults) await deleteNotaSP(nota);
    }

    extractedResults = [];
    document.getElementById('resultsBody').innerHTML = '';
    document.getElementById('btnExport').classList.add('hidden');
    document.getElementById('btnExport').disabled = true;
    updateCoverageStats();
}

function onNotaCheckChange() {
    const checked = document.querySelectorAll('.nota-checkbox:checked');
    const all = document.querySelectorAll('.nota-checkbox');
    const selAll = document.getElementById('selectAllNotas');
    selAll.checked = all.length > 0 && checked.length === all.length;
    selAll.indeterminate = checked.length > 0 && checked.length < all.length;
    const bar = document.getElementById('notasSelecaoBar');
    if (checked.length > 0) {
        bar.classList.remove('hidden'); bar.classList.add('flex');
        document.getElementById('notasSelCount').textContent = checked.length;
    } else {
        bar.classList.add('hidden'); bar.classList.remove('flex');
    }
}

function toggleSelectAllNotas(checked) {
    document.querySelectorAll('.nota-checkbox').forEach(cb => cb.checked = checked);
    onNotaCheckChange();
}

function resetNotasSelecao() {
    document.querySelectorAll('.nota-checkbox').forEach(cb => cb.checked = false);
    const selAll = document.getElementById('selectAllNotas');
    if (selAll) { selAll.checked = false; selAll.indeterminate = false; }
    const bar = document.getElementById('notasSelecaoBar');
    bar.classList.add('hidden'); bar.classList.remove('flex');
}

async function excluirNotasSelecionadas() {
    const checked = document.querySelectorAll('.nota-checkbox:checked');
    if (!checked.length) return;
    if (!confirm(`Excluir ${checked.length} nota(s)? Esta ação não pode ser desfeita.`)) return;

    setButtonLoading('btnExcluirNotas', 'btnExcluirNotasIcon', 'btnExcluirNotasText', 'Excluindo...');

    const indices = Array.from(checked)
        .map(cb => parseInt(cb.dataset.idx))
        .sort((a, b) => b - a);

    for (const idx of indices) {
        const nota = extractedResults[idx];
        if (useSharePoint) await deleteNotaSP(nota);
        extractedResults.splice(idx, 1);
    }

    const tbody = document.getElementById('resultsBody');
    tbody.innerHTML = '';
    getNotasFiltradas().forEach(r => renderRow(r));
    updateCoverageStats();
    atualizarOpcoesTime();
    resetButton('btnExcluirNotas', 'btnExcluirNotasIcon', 'btnExcluirNotasText', 'Excluir selecionadas');
    resetNotasSelecao();

    if (extractedResults.length === 0) {
        document.getElementById('btnExport').classList.add('hidden');
        document.getElementById('btnExport').disabled = true;
    }
}

function updateCoverageStats() {
    const fmtMoney = (v) => (v || 0).toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'});
    const filtradas = getNotasFiltradas();

    // Mostrar/ocultar botão Excel sempre que os dados mudarem
    const btnExport = document.getElementById('btnExport');
    if (btnExport) {
        if (extractedResults.length > 0) {
            btnExport.classList.remove('hidden');
            btnExport.disabled = false;
        } else {
            btnExport.classList.add('hidden');
            btnExport.disabled = true;
        }
    }

    if (extractedResults.length === 0) {
        document.getElementById('statsPanel').classList.add('hidden');
        document.getElementById('totalLiquidoPanel').classList.add('hidden');
        document.getElementById('totalServicoPanel').classList.add('hidden');
        document.getElementById('pixCoveragePanel').classList.add('hidden');
        return;
    }

    const totalNotas = filtradas.length;
    if (totalNotas === 0) {
        // Filtro ativo sem resultados — esconde cards
        ['statsPanel','totalLiquidoPanel','totalServicoPanel','pixCoveragePanel']
            .forEach(id => document.getElementById(id).classList.add('hidden'));
        return;
    }

    // Cobertura de contratos: por motorista único nas notas filtradas
    const motoristasSet = new Map(); // vehicleId → tem contrato?
    for (const r of filtradas) {
        const motoristas = (r.motoristasNaNota && r.motoristasNaNota.length > 0)
            ? r.motoristasNaNota
            : (r.responsavel ? [r.responsavel] : []);
        for (const m of motoristas) {
            if (!m.vehicleId) continue;
            const p = findPrestadorByVehicleId(m.vehicleId);
            const hasContrato = !!(p && p.contrato);
            // Se o motorista aparece em múltiplas notas, considera que tem contrato se tiver em qualquer uma
            if (!motoristasSet.has(m.vehicleId) || hasContrato) {
                motoristasSet.set(m.vehicleId, hasContrato);
            }
        }
    }
    const totalMotoristas = motoristasSet.size;
    const motoristasComContrato = [...motoristasSet.values()].filter(Boolean).length;
    const percentage = totalMotoristas > 0 ? Math.round((motoristasComContrato / totalMotoristas) * 100) : 0;
    document.getElementById('statsPanel').classList.remove('hidden');
    document.getElementById('coveragePercent').textContent = `${percentage}%`;
    document.getElementById('coverageText').textContent = `${motoristasComContrato} de ${totalMotoristas} com contrato`;

    const totalLiquido = filtradas.reduce((sum, r) => sum + (r.vLiquido || 0), 0);
    const totalServico = filtradas.reduce((sum, r) => sum + (r.vServico || 0), 0);
    const label = `${totalNotas} nota${totalNotas !== 1 ? 's' : ''}${_filtroTimesAtivos.size > 0 ? ` · ${[..._filtroTimesAtivos].join(', ')}` : ''}`;

    document.getElementById('totalLiquidoPanel').classList.remove('hidden');
    document.getElementById('totalLiquidoValue').textContent = fmtMoney(totalLiquido);
    document.getElementById('totalLiquidoCount').textContent = label;

    document.getElementById('totalServicoPanel').classList.remove('hidden');
    document.getElementById('totalServicoValue').textContent = fmtMoney(totalServico);
    document.getElementById('totalServicoCount').textContent = label;

    const notasComPix = filtradas.filter(r => r.chavePix && r.chavePix.trim() !== '').length;
    const pixPct = Math.round((notasComPix / totalNotas) * 100);
    document.getElementById('pixCoveragePanel').classList.remove('hidden');
    document.getElementById('pixCoveragePercent').textContent = `${pixPct}%`;
    document.getElementById('pixCoverageText').textContent = `${notasComPix} de ${totalNotas} com PIX`;
}

function reprocessExtractedResults() {
    if (extractedResults.length === 0) {
        console.log('Nenhum resultado para reprocessar');
        // Ainda atualiza a UI mesmo sem notas
        updateCoverageStats();
        return;
    }

    console.log(`🔄 Reprocessando ${extractedResults.length} notas...`);
    
    let contratosAtualizados = 0;
    let pixAtualizados = 0;

    const notasLegadas = []; // notas antigas classificadas agora (motoristasManual inferido)

    // Atualizar cada resultado
    extractedResults.forEach(result => {
        // Responsável pelo CNPJ + motoristas automáticos; lista manual é preservada (regras.js)
        const migracao = reconciliarMotoristasDaNota(result, prestadoresList);
        if (migracao.migrada) notasLegadas.push({ nota: result, listaAlterada: migracao.listaAlterada });

        // Re-verificar contrato para todos os motoristas da nota
        const contratoAntes = result.temContrato;
        const contratoInfo = verificarContratoNota(result);
        result.temContrato       = contratoInfo.status !== 'nao';
        result.temContratoStatus = contratoInfo.status; // 'sim' | 'parcial' | 'nao'
        result.nomeContrato      = contratoInfo.contrato;
        if (!contratoAntes && result.temContrato) contratosAtualizados++;

        // Re-verificar PIX (CNPJ-first)
        const pixAntes = result.chavePix || '';
        result.chavePix = findPixByFileName(result.arquivo, result.cnpj);
        if (!pixAntes && result.chavePix) pixAtualizados++;
    });

    // Re-renderizar tabela
    const tbody = document.getElementById('resultsBody');
    tbody.innerHTML = '';
    extractedResults.forEach(result => renderRow(result));

    atualizarOpcoesTime();
    updateCoverageStats();

    // Grava a classificação inferida para que ela não mude quando o cadastro mudar. Só com
    // prestadores do SharePoint carregados (com a lista local a inferência seria outra).
    if (notasLegadas.length > 0 && useSharePoint && prestadoresList.length > 0) {
        persistirClassificacaoNotasLegadas(notasLegadas);
    }

    console.log(`✅ Reprocessamento concluído! Contratos: +${contratosAtualizados} | PIX: +${pixAtualizados}`);
}

let _ocrEditIdx = null;

function abrirOcrEdit(notaIdx) {
    _ocrEditIdx = notaIdx;
    const r = extractedResults[notaIdx];
    document.getElementById('ocrEditArquivo').textContent = r.arquivo;
    document.getElementById('ocrPrestador').value    = r.prestador || '';
    document.getElementById('ocrCnpj').value         = r.cnpj || '';
    document.getElementById('ocrNumeroNF').value     = r.numeroNF || '';
    document.getElementById('ocrVServico').value     = r.vServico || '';
    document.getElementById('ocrVLiquido').value     = r.vLiquido || '';
    document.getElementById('ocrDataEmissao').value  = r.dataEmissao || '';
    document.getElementById('ocrModelo').value       = r.modelo || '';
    document.getElementById('ocrEditModal').classList.remove('hidden');
    lucide.createIcons();
}

function fecharOcrEdit() {
    _ocrEditIdx = null;
    document.getElementById('ocrEditModal').classList.add('hidden');
}

async function salvarOcrEdit() {
    if (_ocrEditIdx === null) return;
    const r = extractedResults[_ocrEditIdx];

    r.prestador   = document.getElementById('ocrPrestador').value.trim();
    r.cnpj        = document.getElementById('ocrCnpj').value.trim();
    r.numeroNF    = document.getElementById('ocrNumeroNF').value.trim();
    r.vServico    = parseFloat(document.getElementById('ocrVServico').value) || 0;
    r.vLiquido    = parseFloat(document.getElementById('ocrVLiquido').value) || 0;
    r.dataEmissao = document.getElementById('ocrDataEmissao').value.trim();
    r.modelo      = document.getElementById('ocrModelo').value.trim();

    // Reatualizar responsável e PIX com base no novo CNPJ
    const prestadorResp = findPrestadorByCnpj(r.cnpj);
    r.responsavel = prestadorResp ? { vehicleId: prestadorResp.vehicleId, nome: prestadorResp.nome } : null;
    const campos = [...CAMPOS_EDICAO_OCR];
    if (!r.motoristasManual && r.motoristasNaNota && r.motoristasNaNota.length === 0 && r.responsavel) {
        r.motoristasNaNota = [{ ...r.responsavel }];
        campos.push('motoristasNaNota');
    }
    r.chavePix = findPixByFileName(r.arquivo, r.cnpj);
    const contratoInfoM = verificarContratoNota(r);
    r.temContrato       = contratoInfoM.status !== 'nao';
    r.temContratoStatus = contratoInfoM.status;
    r.nomeContrato      = contratoInfoM.contrato;

    await persistirNotaEditada(r, campos);

    // Re-renderizar tabela
    const tbody = document.getElementById('resultsBody');
    tbody.innerHTML = '';
    getNotasFiltradas().forEach(row => renderRow(row));
    updateCoverageStats();
    fecharOcrEdit();
}
