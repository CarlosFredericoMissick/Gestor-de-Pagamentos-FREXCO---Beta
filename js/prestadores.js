// ========================================
// PRESTADORES — cruzamento, importações, cadastro, lista e abas do modal
// ========================================

function sortPrestadores() {
    prestadoresList.sort((a, b) => (a.nome || '').localeCompare(b.nome || '', 'pt-BR', { sensitivity: 'base' }));
}

// Gerador de ID sequencial para cadastros manuais
function nextManualId() {
    const manuals = prestadoresList
        .map(p => String(p.vehicleId || ''))
        .filter(id => /^M\d+$/.test(id))
        .map(id => parseInt(id.slice(1)))
        .filter(n => !isNaN(n));
    const max = manuals.length > 0 ? Math.max(...manuals) : 0;
    return 'M' + (max + 1);
}

// ========================================
// CRUZAMENTO (regras em regras.js, aplicadas à lista carregada)
// ========================================

function findPrestadorByCnpj(cnpj) {
    return buscarPrestadorPorCnpj(prestadoresList, cnpj);
}

function findPrestadorByVehicleId(vehicleId) {
    return buscarPrestadorPorVehicleId(prestadoresList, vehicleId);
}

function findPixByCnpj(cnpj) {
    if (!cnpj) return '';
    const norm = normalizeCnpj(cnpj);
    for (const p of prestadoresList) {
        if (p.inUse === false) continue; // apenas motoristas em uso
        if (!Array.isArray(p.cnpjs)) continue;
        const entry = p.cnpjs.find(c => normalizeCnpj(c.cnpj) === norm);
        if (entry) return entry.chavePix || '';
    }
    return '';
}

// Verifica contrato para todos os motoristas na nota
// Retorna: { status: 'sim'|'nao'|'parcial', contrato: nomeContrato|null }
function verificarContratoNota(nota) {
    const motoristas = (nota.motoristasNaNota && nota.motoristasNaNota.length > 0)
        ? nota.motoristasNaNota
        : (nota.responsavel ? [nota.responsavel] : []);

    if (motoristas.length === 0) return { status: 'nao', contrato: null };

    let comContrato = 0;
    let primeiroContrato = null;
    for (const m of motoristas) {
        const p = findPrestadorByVehicleId(m.vehicleId);
        if (p && p.contrato) {
            comContrato++;
            if (!primeiroContrato) primeiroContrato = p.contrato.nomeArquivo;
        }
    }
    if (comContrato === 0)              return { status: 'nao',     contrato: null };
    if (comContrato === motoristas.length) return { status: 'sim',  contrato: primeiroContrato };
    return { status: 'parcial', contrato: primeiroContrato };
}

function findPixByFileName(fileName, cnpj) {
    // PIX resolvido exclusivamente pelo CNPJ — sem fallback por nome de arquivo
    if (cnpj) return findPixByCnpj(cnpj);
    return '';
}


async function importarCnpjPixExcel(file) {
    try {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }).slice(1)
            .filter(r => r[0] || r[1]);

        const conflitos = [];
        let atualizados = 0, ignorados = 0;

        // ── Fase 1: Simular todas as mudanças em cópias virtuais ──
        // Cada entrada do mapa: sid (String vehicleId) → { prestador, cnpjs simulados }
        const simulacao = new Map();

        for (const row of rows) {
            const rawId    = row[0] != null ? String(row[0]).trim() : '';
            const nomeRow  = (row[1] || '').toString().trim();
            const cnpj     = (row[2] || '').toString().trim();
            const chavePix = (row[3] || '').toString().trim();

            if (!cnpj && !chavePix) { ignorados++; continue; }
            if (!rawId) { conflitos.push(`ID ausente — linha "${nomeRow}" ignorada`); continue; }

            // Líderes (IDs "L…") também estão em prestadoresList
            const prestador = prestadoresList.find(p => String(p.vehicleId) === rawId) || null;
            if (!prestador) { conflitos.push(`Não encontrado: ID "${rawId}" / "${nomeRow}"`); continue; }

            // Inicializa simulação para este prestador (cópia profunda dos CNPJs atuais)
            const sid = String(prestador.vehicleId);
            if (!simulacao.has(sid)) {
                simulacao.set(sid, {
                    prestador,
                    cnpjs: (prestador.cnpjs || []).map(c => ({ ...c }))
                });
            }
            const sim = simulacao.get(sid);

            // Aplica a mudança na cópia virtual
            if (cnpj) {
                const norm = normalizeCnpj(cnpj);
                const existing = sim.cnpjs.find(c => normalizeCnpj(c.cnpj) === norm);
                if (existing) {
                    if (chavePix) existing.chavePix = chavePix;
                } else {
                    sim.cnpjs.push({ cnpj, chavePix });
                }
            } else if (chavePix && sim.cnpjs.length > 0) {
                sim.cnpjs[0].chavePix = chavePix;
            }
        }

        // ── Fase 2: Validar duplicatas no estado final simulado ──
        // Mapa de CNPJ normalizado → sid do prestador dono no resultado final
        const cnpjFinalMap = new Map();
        const sidsComConflito = new Set();

        // 2a) Popula com CNPJs de prestadores NÃO modificados por esta importação
        for (const p of prestadoresList) {
            const sid = String(p.vehicleId);
            if (!simulacao.has(sid) && p.inUse !== false) {
                for (const c of (p.cnpjs || [])) {
                    const norm = normalizeCnpj(c.cnpj);
                    if (norm) cnpjFinalMap.set(norm, sid);
                }
            }
        }

        // 2b) Adiciona CNPJs do estado simulado e detecta conflitos
        for (const [sid, sim] of simulacao) {
            if (sim.prestador.inUse === false) continue;
            for (const c of sim.cnpjs) {
                const norm = normalizeCnpj(c.cnpj);
                if (!norm) continue;
                if (cnpjFinalMap.has(norm) && cnpjFinalMap.get(norm) !== sid) {
                    // CNPJ duplicado no estado final — marca ambos os prestadores
                    const outroSid = cnpjFinalMap.get(norm);
                    sidsComConflito.add(sid);
                    sidsComConflito.add(outroSid);
                    conflitos.push(`CNPJ ${c.cnpj} duplicado no resultado final — alterações para "${sim.prestador.nome}" ignoradas`);
                } else {
                    cnpjFinalMap.set(norm, sid);
                }
            }
        }

        // ── Fase 3: Efetivar apenas prestadores sem conflito ──
        for (const [sid, sim] of simulacao) {
            if (sidsComConflito.has(sid)) continue;
            sim.prestador.cnpjs = sim.cnpjs;
            sim.prestador.dataAtualizacao = new Date().toISOString();
            if (!useSharePoint) await savePrestadorLocal(sim.prestador);
            atualizados++;
        }

        if (useSharePoint) {
            await savePrestadoresSP_internal();
            await loadPrestadoresFromSharePoint();
        } else {
            await loadPrestadoresLocal();
        }
        reprocessExtractedResults();

        let msg = `✅ CNPJs/PIX importados!\n\n${atualizados} prestador(es) atualizados`;
        if (ignorados) msg += `\n${ignorados} linha(s) sem dados ignoradas`;
        if (conflitos.length) msg += `\n\n⚠️ ${conflitos.length} conflito(s):\n${conflitos.slice(0, 10).join('\n')}`;
        alert(msg);
    } catch(e) {
        console.error(e);
        alert('Erro ao processar planilha de CNPJs/PIX.');
    }
}

// ========================================
// IMPORTAR CONTRATOS EM MASSA (ZIP)
// Formato dos arquivos: contrato_{vehicleId}_{nome}.pdf
// ========================================

async function importarContratosZip(file) {
    try {
        const buf = await file.arrayBuffer();
        const zip = await JSZip.loadAsync(buf);
        const pdfFiles = Object.keys(zip.files).filter(n => n.toLowerCase().endsWith('.pdf') && !zip.files[n].dir);

        let importados = 0, substituidos = 0;
        const naoEncontrados = [];

        for (const fileName of pdfFiles) {
            const baseName = fileName.replace(/^.*[\\/]/, '');
            // Parse: contrato_{driverId}_{nome}.pdf
            const match = baseName.match(/^contrato_(\d+)_(.+)\.pdf$/i);

            let prestador = null;
            if (match) {
                const driverId = parseInt(match[1]);
                // Match exclusivo por driver_id
                prestador = findPrestadorByVehicleId(driverId);
            }
            // Sem match por nome — arquivo precisa seguir o padrão contrato_{driverId}_{nome}.pdf

            if (!prestador) {
                naoEncontrados.push(baseName);
                continue;
            }

            const jaTemContrato = prestador.contrato != null;

            // Ler o PDF como base64
            const pdfData = await zip.files[fileName].async('base64');
            const dataUrl = `data:application/pdf;base64,${pdfData}`;

            if (useSharePoint) {
                // Converter base64 para Blob e fazer upload
                const binaryStr = atob(pdfData);
                const bytes = new Uint8Array(binaryStr.length);
                for (let i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i);
                const blob = new Blob([bytes], { type: 'application/pdf' });
                const spFile = new File([blob], baseName, { type: 'application/pdf' });
                const result = await uploadContratoSP(prestador.vehicleId || prestador.id, spFile);
                if (result) {
                    if (useSharePoint && prestador.contrato && prestador.contrato.id) {
                        await deleteContratoSP(prestador.contrato.id);
                    }
                    prestador.contrato = result;
                }
            } else {
                prestador.contrato = { nomeArquivo: baseName, fileData: dataUrl, dataUpload: new Date().toISOString() };
                await savePrestadorLocal(prestador);
            }

            jaTemContrato ? substituidos++ : importados++;
        }

        if (useSharePoint) {
            await savePrestadoresSP_internal();
            await loadPrestadoresFromSharePoint();
        } else {
            await loadPrestadoresLocal();
        }
        reprocessExtractedResults();

        let msg = `✅ Contratos importados!\n\n${importados} novo(s)\n${substituidos} substituído(s)`;
        if (naoEncontrados.length) msg += `\n\n⚠️ ${naoEncontrados.length} não encontrado(s):\n${naoEncontrados.slice(0, 10).join('\n')}`;
        alert(msg);
    } catch(e) {
        console.error(e);
        alert('Erro ao processar ZIP de contratos. Verifique se o arquivo é válido.');
    }
}

async function importarBaseMotoristasExcel(file) {
    try {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const allRows = XLSX.utils.sheet_to_json(ws, { header: 1 });
        const header = allRows[0] || [];
        const rows = allRows.slice(1).filter(r => r[0]);

        // Formato único: ID Motorista(0)|Motorista(1)|ID Veículo(2)|Veículo(3)|CD(4)|In Use(5)|Time(6)
        // Match exclusivo por driver_id — vehicle_id armazenado em vehicles[] apenas
        const CDS_TESTE = ['cd teste - frexco', 'cd teste - piedade', 'cd teste'];
        const idsNaBase = new Set();
        const rowsValidos = [];

        // Agrupar por driver_id — um único prestador por motorista
        const byDriver = {};
        for (const row of rows) {
            const driverId  = parseInt(row[0]);
            const nome      = (row[1] || '').toString().trim();
            const vehicleId = row[2] ? parseInt(row[2]) : null;
            const veiculo   = (row[3] || '').toString().trim();
            const cdRaw     = (row[4] || '').toString().trim();
            const inUse     = row[5] === true || row[5] === 'true' || row[5] === 1;
            const time      = (row[6] || '').toString().trim();
            if (!nome || isNaN(driverId)) continue;
            if (cdRaw && CDS_TESTE.includes(cdRaw.toLowerCase())) continue;

            if (!byDriver[driverId]) {
                byDriver[driverId] = {
                    driverId, nome, cd: cdRaw || 'CD Desconhecido', inUse, time,
                    vehicles: []
                };
            } else {
                // Mesmo motorista em múltiplas linhas (1 por veículo):
                // inUse usa lógica OR — se qualquer linha for ativa, o motorista é ativo
                if (inUse) byDriver[driverId].inUse = true;
                // Atualiza time e cd se vieram vazios na primeira linha
                if (time && !byDriver[driverId].time) byDriver[driverId].time = time;
                if (cdRaw && byDriver[driverId].cd === 'CD Desconhecido') byDriver[driverId].cd = cdRaw;
            }
            if (vehicleId) byDriver[driverId].vehicles.push({ vehicleId, nome: veiculo || nome });
        }
        for (const entry of Object.values(byDriver)) {
            idsNaBase.add(entry.driverId);
            rowsValidos.push(entry);
        }

        let novos = 0, atualizados = 0, excluidos = 0;

        // Match EXCLUSIVO por driverId — sem fallback por vehicleId
        const driverIdMap = new Map();
        for (const entry of rowsValidos) {
            driverIdMap.set(entry.driverId, entry);
        }

        const matchedDriverIds = new Set();

        for (const p of [...prestadoresList]) {
            const vid = parseInt(p.vehicleId);
            if (isNaN(vid) || String(p.vehicleId).startsWith('M') || String(p.vehicleId).startsWith('L')) continue;

            const entry = driverIdMap.get(vid);
            if (entry) {
                if (matchedDriverIds.has(entry.driverId)) {
                    // Duplicata: mescla dados no primário e descarta
                    const primario = prestadoresList.find(x => parseInt(x.vehicleId) === entry.driverId);
                    if (primario) {
                        if ((!primario.cnpjs || !primario.cnpjs.length) && p.cnpjs?.length)
                            primario.cnpjs = p.cnpjs;
                        if (!primario.contrato && p.contrato)
                            primario.contrato = p.contrato;
                    }
                    if (!useSharePoint && p.id) await deletePrestadorLocal(p.id);
                    excluidos++;
                } else {
                    // Match: atualiza cadastro, preserva cnpjs e contrato
                    p.vehicleId = entry.driverId;
                    p.nome      = entry.nome;
                    p.cd        = entry.cd;
                    p.inUse     = entry.inUse;
                    p.time      = entry.time || p.time || '';
                    p.vehicles  = entry.vehicles || [];
                    if (!useSharePoint) await savePrestadorLocal(p);
                    matchedDriverIds.add(entry.driverId);
                    atualizados++;
                }
            } else {
                // Não está na nova base → excluir
                if (!useSharePoint && p.id) await deletePrestadorLocal(p.id);
                excluidos++;
            }
        }

        // Filtra lista: mantém migrados + manuais (M-) + líderes (L-)
        prestadoresList = prestadoresList.filter(p => {
            const id = String(p.vehicleId || '');
            if (id.startsWith('M') || id.startsWith('L')) return true;
            return idsNaBase.has(parseInt(id));
        });

        // Cria registros novos para driverIds sem match
        for (const entry of rowsValidos) {
            if (!matchedDriverIds.has(entry.driverId)) {
                const novo = {
                    vehicleId: entry.driverId, nome: entry.nome,
                    cd: entry.cd, time: entry.time || '', inUse: entry.inUse,
                    vehicles: entry.vehicles || [], cnpjs: [], contrato: null,
                    fromSharePoint: useSharePoint,
                    dataAdicao: new Date().toISOString()
                };
                if (!useSharePoint) { const id = await savePrestadorLocal(novo); novo.id = id; }
                prestadoresList.push(novo);
                novos++;
            }
        }

        if (useSharePoint) {
            sortPrestadores();
            await sincronizarLideres();
            await savePrestadoresSP_internal();
            await loadPrestadoresFromSharePoint();
        } else {
            await sincronizarLideres();
            await loadPrestadoresLocal();
        }
        reprocessExtractedResults();

        let msg = `✅ Base sincronizada!\n\n${novos} novos\n${atualizados} atualizados\n${excluidos} excluídos`;
        alert(msg);
    } catch(e) {
        console.error(e);
        alert('Erro ao processar planilha de motoristas: ' + e.message);
    }
}

async function limparBasePrestadores() {
    const total = prestadoresList.length;
    if (!confirm(`⚠️ ATENÇÃO: Isso vai excluir TODOS os ${total} motoristas cadastrados (ativos e inativos), incluindo CNPJs, PIX e contratos.\n\nUse esta opção antes de reimportar uma base nova do zero.\n\nEsta ação não pode ser desfeita. Continuar?`)) return;

    // Apagar todos os registros de motoristas (ativos e inativos)
    if (useSharePoint) {
        prestadoresList = [];
        // skipMerge=true: impede que o merge reinsira os registros excluídos
        await savePrestadoresSP_internal(true);
        updatePrestadoresUI();
    } else {
        for (const p of [...prestadoresList]) {
            if (p.id) await deletePrestadorLocal(p.id);
        }
        prestadoresList = [];
        await loadPrestadoresLocal();
    }

    reprocessExtractedResults();
    updatePrestadoresUI();
    alert(`✅ Base limpa: ${total} motoristas removidos.\n\nAgora importe a base de motoristas, CNPJs/PIX e contratos do zero.`);
}

function togglePrestadoresModal() {
    document.getElementById('prestadoresModal').classList.toggle('hidden');
}

let editingPrestadorIndex = null;

function renderCnpjRows(cnpjs) {
    const container = document.getElementById('cnpjRowsContainer');
    if (!cnpjs || cnpjs.length === 0) {
        container.innerHTML = buildCnpjRow('', '', 0);
    } else {
        container.innerHTML = cnpjs.map((c, i) => buildCnpjRow(c.cnpj, c.chavePix, i)).join('');
    }
}

function buildCnpjRow(cnpj, pix, idx) {
    return `<div class="flex gap-2 items-center cnpj-row" data-idx="${idx}">
        <input type="text" placeholder="00.000.000/0001-00" value="${cnpj}"
            class="cnpj-input flex-1 px-3 py-1.5 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-frexco/40 outline-none">
        <input type="text" placeholder="Chave PIX" value="${pix}"
            class="pix-input flex-1 px-3 py-1.5 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-frexco/40 outline-none">
        <button onclick="removeCnpjRow(${idx})" class="p-1.5 text-red-400 hover:text-red-600 rounded-lg">✕</button>
    </div>`;
}

function addCnpjRow() {
    const container = document.getElementById('cnpjRowsContainer');
    const idx = container.querySelectorAll('.cnpj-row').length;
    container.insertAdjacentHTML('beforeend', buildCnpjRow('', '', idx));
}

function removeCnpjRow(idx) {
    const rows = document.querySelectorAll('#cnpjRowsContainer .cnpj-row');
    if (rows.length <= 1) { rows[0].querySelector('.cnpj-input').value = ''; rows[0].querySelector('.pix-input').value = ''; return; }
    rows[idx]?.remove();
}

function getCnpjRowsData() {
    return Array.from(document.querySelectorAll('#cnpjRowsContainer .cnpj-row'))
        .map(row => ({ cnpj: row.querySelector('.cnpj-input').value.trim(), chavePix: row.querySelector('.pix-input').value.trim() }))
        .filter(c => c.cnpj || c.chavePix);
}

function showAddPrestadorForm() {
    editingPrestadorIndex = null;
    document.getElementById('prestadorFormTitle').textContent = 'Adicionar Prestador';
    document.getElementById('prestadorFormSaveText').textContent = 'Salvar';
    document.getElementById('prestadorNome').value = '';
    renderCnpjRows([]);
    document.getElementById('prestadorForm').classList.remove('hidden');
    document.getElementById('prestadorNome').focus();
}

function editPrestadorEntry(index) {
    editingPrestadorIndex = index;
    const p = prestadoresList[index];
    document.getElementById('prestadorFormTitle').textContent = 'Editar Prestador';
    document.getElementById('prestadorFormSaveText').textContent = 'Atualizar';
    document.getElementById('prestadorNome').value = p.nome || '';
    renderCnpjRows(p.cnpjs || []);
    document.getElementById('prestadorForm').classList.remove('hidden');
    document.getElementById('prestadorNome').focus();
}

function cancelPrestadorForm() {
    editingPrestadorIndex = null;
    document.getElementById('prestadorForm').classList.add('hidden');
}

async function savePrestadorEntry() {
    const nome  = document.getElementById('prestadorNome').value.trim();
    const cnpjs = getCnpjRowsData();
    if (!nome) { alert('Preencha o nome do motorista.'); return; }

    for (const { cnpj } of cnpjs) {
        if (!cnpj) continue;
        const norm = normalizeCnpj(cnpj);
        const dono = prestadoresList.find((p, i) =>
            i !== editingPrestadorIndex &&
            Array.isArray(p.cnpjs) && p.cnpjs.some(c => normalizeCnpj(c.cnpj) === norm)
        );
        if (dono) { alert(`CNPJ ${cnpj} já vinculado a "${dono.nome}".`); return; }
    }

    if (editingPrestadorIndex !== null) {
        const p = prestadoresList[editingPrestadorIndex];
        p.nome = nome; p.cnpjs = cnpjs; p.dataAtualizacao = new Date().toISOString();
        if (useSharePoint) { await savePrestadoresSP_internal(); }
        else { await savePrestadorLocal(p); }
    } else {
        const manualId = nextManualId();
        const novo = { vehicleId: manualId, nome, cd: '', time: '', inUse: true, cnpjs, contrato: null, fromSharePoint: useSharePoint, dataAdicao: new Date().toISOString() };
        if (useSharePoint) { prestadoresList.push(novo); await savePrestadoresSP_internal(); }
        else { const newId = await savePrestadorLocal(novo); novo.id = newId; prestadoresList.push(novo); }
    }
    sortPrestadores(); updatePrestadoresUI(); cancelPrestadorForm(); reprocessExtractedResults();
}

async function uploadContratoPrestador(index) {
    document.getElementById('contratoUploadInput').onchange = async (e) => {
        const file = e.target.files[0]; if (!file) return;
        const p = prestadoresList[index];
        if (useSharePoint) {
            const result = await uploadContratoSP(p.vehicleId || p.id || index, file);
            if (result) { p.contrato = result; await savePrestadoresSP_internal(); await loadPrestadoresFromSharePoint(); }
            else alert('Erro ao enviar contrato para o SharePoint.');
        } else {
            const reader = new FileReader();
            reader.onload = async (ev) => {
                p.contrato = { nomeArquivo: file.name, fileData: ev.target.result, dataUpload: new Date().toISOString() };
                await savePrestadorLocal(p); await loadPrestadoresLocal(); reprocessExtractedResults();
            };
            reader.readAsDataURL(file);
        }
        document.getElementById('contratoUploadInput').value = '';
    };
    document.getElementById('contratoUploadInput').click();
}

async function removeContratoPrestador(index) {
    if (!confirm('Remover o contrato deste prestador?')) return;
    const p = prestadoresList[index];
    if (useSharePoint && p.contrato && p.contrato.id) await deleteContratoSP(p.contrato.id);
    p.contrato = null;
    if (useSharePoint) { await savePrestadoresSP_internal(); await loadPrestadoresFromSharePoint(); }
    else { await savePrestadorLocal(p); await loadPrestadoresLocal(); }
    reprocessExtractedResults();
}

async function viewContratoPrestador(index) {
    const p = prestadoresList[index]; if (!p.contrato) return;
    if (useSharePoint && p.contrato.id) {
        try {
            const resp = await graphFetch(`${GRAPH_ITEMS}/${p.contrato.id}`);
            if (resp.ok) { const d = await resp.json(); if (d['@microsoft.graph.downloadUrl']) window.open(d['@microsoft.graph.downloadUrl'], '_blank'); }
        } catch(e) { alert('Erro ao abrir contrato.'); }
    } else if (p.contrato.fileData) {
        const a = document.createElement('a'); a.href = p.contrato.fileData; a.download = p.contrato.nomeArquivo; a.click();
    }
}

async function removePrestador(index) {
    const p = prestadoresList[index];
    if (!confirm(`Excluir "${p.nome}"?`)) return;
    if (useSharePoint) {
        if (p.contrato && p.contrato.id) await deleteContratoSP(p.contrato.id);
        prestadoresList.splice(index, 1);
        await savePrestadoresSP_internal(); await loadPrestadoresFromSharePoint();
    } else { await deletePrestadorLocal(p.id); await loadPrestadoresLocal(); }
    reprocessExtractedResults();
}

function updatePrestadoresUI() {
    const ativos = prestadoresList.filter(p => p.inUse !== false && p.tipo !== 'lider');
    document.getElementById('prestadoresCount').textContent = ativos.length;
    document.getElementById('prestadoresSavedCount').textContent = ativos.length;
    document.getElementById('lideresCount').textContent = getLideres().length;

    const bar = document.getElementById('prestadoresSelectAllBar');
    const container = document.getElementById('prestadoresListContainer');
    const query = (document.getElementById('prestadoresSearch')?.value || '').toLowerCase().trim();
    const filtroPix      = document.getElementById('filtroPix')?.checked || false;
    const filtroContrato = document.getElementById('filtroContrato')?.checked || false;

    const queryDigits = query.replace(/\D/g, '');
    const visiveis = ativos.filter(p => {
        if (p.tipo === 'lider') return false; // líderes ficam na aba Líderes
        if (query && !(
            (p.nome || '').toLowerCase().includes(query) ||
            (p.time || '').toLowerCase().includes(query) ||
            (p.cd || '').toLowerCase().includes(query) ||
            (p.vehicleId && String(p.vehicleId).toLowerCase().includes(query)) ||
            (queryDigits.length > 0 && Array.isArray(p.cnpjs) &&
                p.cnpjs.some(c => (c.cnpj||'').replace(/\D/g,'').includes(queryDigits)))
        )) return false;
        if (filtroPix && !(Array.isArray(p.cnpjs) && p.cnpjs.some(c => c.chavePix && c.chavePix.trim()))) return false;
        if (filtroContrato && !p.contrato) return false;
        return true;
    });

    if (ativos.length === 0) {
        container.innerHTML = '<p class="text-sm text-gray-400 text-center py-10">Nenhum motorista ativo. Importe a planilha base.</p>';
        bar.classList.add('hidden'); bar.classList.remove('flex');
        updatePrestadoresSelectionUI(); return;
    }
    bar.classList.remove('hidden'); bar.classList.add('flex');

    if (visiveis.length === 0) {
        container.innerHTML = '<p class="text-sm text-gray-400 text-center py-10">Nenhum resultado encontrado</p>';
        updatePrestadoresSelectionUI(); return;
    }

    container.innerHTML = visiveis.map(p => {
        const index = prestadoresList.indexOf(p);
        const timeBadge = p.time
            ? `<span class="text-xs bg-blue-50 text-blue-700 border border-blue-200 px-1.5 py-0.5 rounded">${p.time}</span>`
            : `<span class="text-xs text-gray-400">sem time</span>`;
        const cnpjsHtml = (Array.isArray(p.cnpjs) && p.cnpjs.length > 0)
            ? p.cnpjs.map(c => `<div class="flex items-center gap-1"><span class="text-xs font-mono text-gray-600 bg-gray-100 px-1 py-0.5 rounded truncate">${c.cnpj||'—'}</span>${c.chavePix ? `<span class="text-xs text-frexco truncate" title="${c.chavePix}">→ ${c.chavePix}</span>` : ''}</div>`).join('')
            : `<span class="text-xs text-orange-500 bg-orange-50 px-1.5 py-0.5 rounded border border-orange-200">⚠ sem CNPJ</span>`;
        const contratoHtml = p.contrato
            ? `<div class="flex items-center gap-1">
                <span class="text-xs text-green-700 bg-green-50 px-2 py-0.5 rounded-full border border-green-200 truncate max-w-[100px]" title="${p.contrato.nomeArquivo}">✓ ${p.contrato.nomeArquivo}</span>
                <button onclick="viewContratoPrestador(${index})" class="p-1 text-blue-600 hover:bg-blue-50 rounded" title="Ver"><svg xmlns="http://www.w3.org/2000/svg" class="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg></button>
                <button onclick="removeContratoPrestador(${index})" class="p-1 text-red-400 hover:bg-red-50 rounded" title="Remover"><svg xmlns="http://www.w3.org/2000/svg" class="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>
               </div>`
            : `<button onclick="uploadContratoPrestador(${index})" class="flex items-center gap-1 text-xs text-blue-600 hover:bg-blue-50 px-2 py-1 rounded-lg border border-blue-200 transition-colors">↑ Anexar PDF</button>`;
        return `
        <div class="grid grid-cols-12 gap-2 items-start p-3 bg-gray-50 hover:bg-gray-100 rounded-xl border border-gray-200 transition-colors text-sm">
            <div class="col-span-1 pt-1">
                <input type="checkbox" class="prestador-checkbox w-4 h-4 accent-frexco cursor-pointer" data-index="${index}" onchange="onPrestadorCheckChange()">
            </div>
            <div class="col-span-3 min-w-0">
                <p class="font-semibold text-gray-900 truncate" title="${p.nome}">${p.nome}</p>
                <div class="flex items-center gap-1 mt-0.5">${timeBadge}</div>
                ${p.vehicleId ? `<span class="text-xs text-gray-400">#${p.vehicleId}</span>` : ''}
            </div>
            <div class="col-span-3 min-w-0 space-y-0.5">${cnpjsHtml}</div>
            <div class="col-span-3 min-w-0">${contratoHtml}</div>
            <div class="col-span-2 flex items-start justify-end gap-1 pt-0.5">
                <button onclick="editPrestadorEntry(${index})" class="p-1.5 text-gray-500 hover:text-frexco hover:bg-frexco/10 rounded-lg" title="Editar"><svg xmlns="http://www.w3.org/2000/svg" class="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg></button>
                <button onclick="removePrestador(${index})" class="p-1.5 text-gray-400 hover:text-red-600 hover:bg-red-50 rounded-lg" title="Excluir"><svg xmlns="http://www.w3.org/2000/svg" class="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/><path d="M9 6V4h6v2"/></svg></button>
            </div>
        </div>`;
    }).join('');

    updatePrestadoresSelectionUI();
}

function onPrestadorCheckChange() {
    const all = document.querySelectorAll('.prestador-checkbox');
    const checked = document.querySelectorAll('.prestador-checkbox:checked');
    const selAll = document.getElementById('selectAllPrestadores');
    selAll.checked = all.length > 0 && checked.length === all.length;
    selAll.indeterminate = checked.length > 0 && checked.length < all.length;
    updatePrestadoresSelectionUI();
}

function toggleSelectAllPrestadores(checked) {
    document.querySelectorAll('.prestador-checkbox').forEach(cb => cb.checked = checked);
    updatePrestadoresSelectionUI();
}

function updatePrestadoresSelectionUI() {
    const checked = document.querySelectorAll('.prestador-checkbox:checked');
    const btn = document.getElementById('btnDeleteSelectedPrestadores');
    if (checked.length > 0) { btn.classList.remove('hidden'); document.getElementById('prestadoresSelectedCount').textContent = checked.length; }
    else { btn.classList.add('hidden'); }
}

async function deleteSelectedPrestadores() {
    const checked = document.querySelectorAll('.prestador-checkbox:checked');
    if (!checked.length || !confirm(`Excluir ${checked.length} prestador(es)?`)) return;
    const indices = Array.from(checked).map(cb => parseInt(cb.dataset.index)).sort((a,b) => b - a);
    for (const i of indices) {
        const p = prestadoresList[i];
        if (useSharePoint && p.contrato && p.contrato.id) await deleteContratoSP(p.contrato.id);
        if (!useSharePoint) await deletePrestadorLocal(p.id);
        prestadoresList.splice(i, 1);
    }
    if (useSharePoint) {
        // skipMerge=true: impede que o merge reinsira os prestadores que acabaram de ser excluídos
        await savePrestadoresSP_internal(true);
        updatePrestadoresUI();
    } else {
        await loadPrestadoresLocal();
    }
    reprocessExtractedResults();
}

function switchPrestadoresTab(tab) {
    _prestadoresTab = tab;
    const bodyMotoristas = document.getElementById('painel-motoristas');
    const bodyLideres    = document.getElementById('painel-lideres');
    const tabM = document.getElementById('tabMotoristas');
    const tabL = document.getElementById('tabLideres');
    if (tab === 'motoristas') {
        bodyMotoristas?.classList.remove('hidden');
        bodyLideres?.classList.add('hidden');
        tabM?.classList.add('border-frexco', 'text-frexco');
        tabM?.classList.remove('border-transparent', 'text-gray-500');
        tabL?.classList.add('border-transparent', 'text-gray-500');
        tabL?.classList.remove('border-purple-500', 'text-purple-600');
    } else {
        bodyMotoristas?.classList.add('hidden');
        bodyLideres?.classList.remove('hidden');
        tabL?.classList.add('border-purple-500', 'text-purple-600');
        tabL?.classList.remove('border-transparent', 'text-gray-500');
        tabM?.classList.add('border-transparent', 'text-gray-500');
        tabM?.classList.remove('border-frexco', 'text-frexco');
        renderLideresList();
        lucide.createIcons();
    }
}

function abrirAjudaCnpjPix() {
    document.getElementById('ajudaCnpjPixModal').classList.remove('hidden');
    lucide.createIcons();
}
function abrirAjudaContratos() {
    document.getElementById('ajudaContratosModal').classList.remove('hidden');
    lucide.createIcons();
}
