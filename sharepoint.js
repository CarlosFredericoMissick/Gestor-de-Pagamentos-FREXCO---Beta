// ========================================
// SHAREPOINT — LEITURA/ESCRITA VIA MICROSOFT GRAPH
// Arquivos na pasta da aplicação (approot):
//   prestadores.json, recebivel.json, notas/<arquivo>_<ts>.json, pdfs/<arquivo>_<ts>.pdf,
//   contrato_<id>_<arquivo>.pdf
// ========================================

function urlApproot(caminho, sufixo = '') {
    return `${GRAPH_APPROOT}/${caminho}${sufixo}`;
}

// Texto do erro para alertas; 401 (ou token que exige login de novo) orienta a reconectar.
function mensagemErroSP(status) {
    if (status === 401 || spExigeReconexao) {
        return `HTTP ${status} — sessão do SharePoint expirada; clique em "Conectar SharePoint" para reconectar e tente de novo`;
    }
    return `HTTP ${status}`;
}

// fetch autenticado no Graph. Renova o token antes da chamada e tenta de novo uma vez em 401.
async function graphFetch(url, opcoes = {}) {
    const comToken = (token) => ({ ...opcoes, headers: { ...(opcoes.headers || {}), 'Authorization': `Bearer ${token}` } });
    const token = await obterTokenSP();
    let resp = await fetch(url, comToken(token));
    if (resp.status === 401 && spAccount) {
        const novo = await obterTokenSP(true);
        if (novo) resp = await fetch(url, comToken(novo));
    }
    return resp;
}

// ========================================
// PRESTADORES JSON
// ========================================

async function loadPrestadoresFromSharePoint() {
    if (!spAccessToken) return;
    try {
        const resp = await graphFetch(urlApproot('prestadores.json', ':/content'));
        if (resp.ok) {
            const data = await resp.json();
            prestadoresList = migratePrestadoresData(data.prestadores || []);
            sortPrestadores();
            updatePrestadoresUI();
            console.log('✅ Prestadores SharePoint:', prestadoresList.length);
        } else if (resp.status === 404) {
            prestadoresList = [];
            await savePrestadoresSP();
        } else {
            throw new Error('Falha ao carregar prestadores');
        }
    } catch(e) {
        console.error('SharePoint prestadores:', e);
        useSharePoint = false;
        await loadPrestadoresLocal();
    }
}

async function savePrestadoresSP() { await savePrestadoresSP_internal(); }

async function savePrestadoresSP_internal(skipMerge = false) {
    if (!spAccessToken) return;

    // Merge: busca versão atual do SP, preserva campos que possam ter sido
    // editados por outro usuário desde o último load (last-write-wins por campo)
    // ATENÇÃO: skipMerge=true deve ser usado ao deletar para evitar que os
    // registros excluídos sejam reinseridos a partir da versão remota.
    if (!skipMerge) {
        try {
            const resp = await graphFetch(urlApproot('prestadores.json', ':/content'));
            if (resp.ok) {
                const spData = await resp.json();
                const spList = migratePrestadoresData(spData.prestadores || []);
                // Para cada prestador no SP que não está na nossa lista local,
                // verificamos se foi adicionado por outro usuário e preservamos
                spList.forEach(spP => {
                    const local = prestadoresList.find(p => String(p.vehicleId) === String(spP.vehicleId));
                    if (!local) {
                        // Adicionado por outro usuário — incluir
                        prestadoresList.push(spP);
                    }
                    // Se existe local, nossa versão prevalece (já tem as edições do usuário atual)
                });
                sortPrestadores();
            }
        } catch(e) {
            // Se falhar o merge, salva com a versão local mesmo
            console.warn('Merge SP falhou, salvando versão local:', e);
        }
    }

    const toSave = prestadoresList.map(p => ({...p, contrato: p.contrato ? {...p.contrato, fileData: null} : null}));
    const body = JSON.stringify({ prestadores: toSave, lastUpdate: new Date().toISOString(), lastModifiedBy: spAccount?.name || spAccount?.username || 'desconhecido' }, null, 2);
    let erro = null;
    try {
        const resp = await graphFetch(
            urlApproot('prestadores.json', ':/content'),
            { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body }
        );
        if (!resp.ok) erro = mensagemErroSP(resp.status);
    } catch(e) {
        console.error('Erro ao salvar prestadores SP:', e);
        erro = e.message || String(e);
    }
    if (erro) {
        alert(`❌ Os prestadores NÃO foram salvos no SharePoint (${erro}).\n\nAs alterações ficaram só neste navegador. Tente novamente.`);
        return false;
    }
    return true;
}

// ========================================
// CONTRATOS (arquivos PDF)
// ========================================

async function uploadContratoSP(prestadorId, file) {
    if (!spAccessToken) return null;
    const safeFileName = `contrato_${prestadorId}_${file.name}`;
    try {
        const buf = await file.arrayBuffer();
        const resp = await graphFetch(urlApproot(safeFileName, ':/content'), { method: 'PUT', body: buf });
        if (!resp.ok) return null;
        const data = await resp.json();
        return { id: data.id, nomeArquivo: file.name, spFileName: safeFileName, dataUpload: new Date().toISOString() };
    } catch(e) {
        console.error('Erro upload contrato SP:', e);
        return null;
    }
}

async function deleteContratoSP(spFileId) {
    if (!spAccessToken || !spFileId) return;
    try {
        await graphFetch(`${GRAPH_ITEMS}/${spFileId}`, { method: 'DELETE' });
    } catch(e) { console.error('Erro delete contrato SP:', e); }
}

// ========================================
// RECEBÍVEL JSON
// ========================================

async function loadRecebivelFromSharePoint() {
    if (!spAccessToken) return;
    try {
        const resp = await graphFetch(urlApproot('recebivel.json', ':/content'));
        if (!resp.ok) return; // arquivo não existe ainda
        const data = await resp.json();
        recebivelData = (data.recebivel || []).map(r => ({
            ...r,
            data:      new Date(r.data),
            valorRotas: r.valorRotas || 0,
            ajuste:     r.ajuste     || 0,
            adicional:  r.adicional  || 0,
            descarga:   r.descarga   || 0,
            suporte:    r.suporte    || 0,
            comissao:   r.comissao   || 0,
            rota:       r.rota       || '',
        }));
        console.log(`✅ Recebível SP: ${recebivelData.length} registros`);
        document.getElementById('recebivelStatus').textContent =
            `Recebível carregado: ${recebivelData.length} registros`;
    } catch(e) { console.error('Erro ao carregar recebível SP:', e); }
}

async function saveRecebivelSP() {
    const toSave = recebivelData.map(r => ({...r, data: r.data.toISOString()}));
    const body = JSON.stringify({ recebivel: toSave, lastUpdate: new Date().toISOString() }, null, 2);
    let erro = null;
    try {
        const resp = await graphFetch(
            urlApproot('recebivel.json', ':/content'),
            { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body }
        );
        if (!resp.ok) erro = mensagemErroSP(resp.status);
    } catch(e) {
        console.error('Erro ao salvar recebível SP:', e);
        erro = e.message || String(e);
    }
    if (erro) {
        alert(`❌ O recebível foi importado, mas NÃO foi salvo no SharePoint (${erro}).\n\nOs outros usuários não verão esta base. Importe novamente.`);
        return false;
    }
    console.log('✅ Recebível salvo no SharePoint');
    return true;
}

async function deleteRecebivelSP() {
    await graphFetch(urlApproot('recebivel.json'), { method: 'DELETE' });
}

// ========================================
// NOTAS — um arquivo JSON por nota em notas/
// ========================================

function notaFileName(nota) {
    const safe = (nota.arquivo || 'nota').replace(/[^a-zA-Z0-9_.-]/g, '_');
    const ts = nota.dataExtracao || Date.now();
    return `notas/${safe}_${ts}.json`;
}

// ---------------- Concorrência otimista ----------------
// Para cada nota carregada/gravada guardamos o eTag do arquivo e o conteúdo remoto
// correspondente (base). A gravação envia If-Match: se outro usuário gravou antes (412),
// a nota é relida, só os campos editados aqui são reaplicados sobre a versão remota e a
// gravação é repetida. Esses metadados ficam fora do JSON gravado (WeakMap).
const _versaoNotasSP = new WeakMap(); // nota → { eTag, base }

function registrarVersaoNotaSP(nota, eTag, conteudo) {
    if (!eTag) return;
    _versaoNotasSP.set(nota, { eTag, base: JSON.parse(JSON.stringify(conteudo)) });
}

function versaoNotaSP(nota) { return _versaoNotasSP.get(nota) || null; }

async function lerNotaSP(fileName) {
    const resp = await graphFetch(urlApproot(fileName));
    if (resp.status === 404) return null;
    if (!resp.ok) throw new Error(mensagemErroSP(resp.status));
    const item = await resp.json();
    const r = await fetch(item['@microsoft.graph.downloadUrl']);
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    return { nota: await r.json(), eTag: item.eTag };
}

const mesmoValor = (a, b) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);

// Grava a nota no SharePoint. Notas sem _spFileName (extraídas sem conexão) ganham um
// arquivo novo em notas/; notas novas (sem eTag) são gravadas sem If-Match.
// opcoes.campos: campos alterados por esta edição (reaplicados em caso de conflito).
//                Sem campos, em conflito a versão local inteira prevalece.
// opcoes.somenteCampos: grava a base remota + apenas `campos` (não leva valores derivados
//                da memória, como o responsável recalculado).
// opcoes.abortarSeRemotoTiver: nome de campo; se após um 412 a versão remota já tiver esse
//                campo definido, não grava e adota a versão remota em memória (abortado: true).
// Retorna { ok, abortado?, mensagem?, conflitos: [campos alterados também por outro usuário] }.
async function saveNotaSP(nota, opcoes = {}) {
    if (!spAccessToken) return { ok: false, mensagem: 'SharePoint não conectado', conflitos: [] };
    const conflitos = [];
    try {
        const fileName = nota._spFileName || notaFileName(nota);
        nota._spFileName = fileName;
        const campos = opcoes.campos || Object.keys(nota);

        for (let tentativa = 0; tentativa < 3; tentativa++) {
            const versao = versaoNotaSP(nota);
            const conteudo = opcoes.somenteCampos && versao
                ? { ...versao.base, ...Object.fromEntries(campos.map(c => [c, nota[c]])), _spFileName: fileName }
                : nota;
            const headers = { 'Content-Type': 'application/json' };
            if (versao) headers['If-Match'] = versao.eTag;
            const resp = await graphFetch(
                urlApproot(fileName, ':/content'),
                { method: 'PUT', headers, body: JSON.stringify(conteudo) }
            );
            if (resp.ok) {
                const item = await resp.json().catch(() => ({}));
                registrarVersaoNotaSP(nota, item.eTag, conteudo);
                return { ok: true, conflitos };
            }
            if (resp.status !== 412 || !versao) {
                console.error('Erro ao salvar nota SP:', resp.status, fileName);
                return { ok: false, mensagem: mensagemErroSP(resp.status), conflitos };
            }

            // 412: outro usuário gravou depois da nossa leitura → mesclar e tentar de novo
            const remoto = await lerNotaSP(fileName);
            if (!remoto) return { ok: false, mensagem: 'a nota foi excluída por outro usuário', conflitos };
            const campoGuarda = opcoes.abortarSeRemotoTiver;
            if (campoGuarda && remoto.nota[campoGuarda] !== undefined) {
                Object.keys(nota).forEach(k => { if (!(k in remoto.nota)) delete nota[k]; });
                Object.assign(nota, remoto.nota);
                registrarVersaoNotaSP(nota, remoto.eTag, remoto.nota);
                return { ok: true, abortado: true, conflitos };
            }
            const meus = Object.fromEntries(campos.map(c => [c, nota[c]]));
            campos.forEach(c => {
                if (!mesmoValor(remoto.nota[c], versao.base[c]) && !mesmoValor(remoto.nota[c], nota[c]) && !conflitos.includes(c)) {
                    conflitos.push(c);
                }
            });
            Object.keys(nota).forEach(k => { if (!(k in remoto.nota) && !campos.includes(k)) delete nota[k]; });
            Object.assign(nota, remoto.nota, meus);
            registrarVersaoNotaSP(nota, remoto.eTag, remoto.nota);
        }
        return { ok: false, mensagem: 'a nota está sendo alterada por outro usuário; tente de novo', conflitos };
    } catch(e) {
        console.error('Erro ao salvar nota SP:', e);
        return { ok: false, mensagem: e.message || String(e), conflitos };
    }
}

// Grava a classificação motoristasManual inferida nesta carga para notas antigas, só nas
// que ainda não têm a flag no SharePoint, sem levar valores derivados da memória.
// Se outro usuário já gravou a nota com a flag (ex.: editou os motoristas no modal depois da
// nossa carga), a inferência é descartada e a versão dele é adotada: nunca sobrescreve uma
// escolha feita por alguém.
async function persistirClassificacaoNotasLegadas(migradas) {
    let falhas = 0, adotadas = 0;
    for (const { nota, listaAlterada } of migradas) {
        const versao = versaoNotaSP(nota);
        if (!nota._spFileName || !versao || 'motoristasManual' in versao.base) continue;
        const campos = listaAlterada ? ['motoristasManual', 'motoristasNaNota'] : ['motoristasManual'];
        const res = await saveNotaSP(nota, { campos, somenteCampos: true, abortarSeRemotoTiver: 'motoristasManual' });
        if (!res.ok) falhas++;
        else if (res.abortado) adotadas++;
    }
    if (falhas) console.warn(`${falhas} nota(s) antiga(s) sem a classificação gravada; será refeita na próxima carga.`);
    // Notas adotadas do SharePoint já têm a flag: o reprocessamento só recalcula e redesenha.
    if (adotadas) reprocessExtractedResults();
}

async function deleteNotaSP(nota) {
    if (!nota?._spFileName) return;
    try {
        await graphFetch(urlApproot(nota._spFileName), { method: 'DELETE' });
    } catch(e) { console.error('Erro ao deletar nota SP:', e); }
}

// aguardar: gravação inicial da nota, para o re-save com a referência do PDF vir depois dela.
// Retorna true se o PDF e a nota foram gravados. A regravação leva a nota inteira (sem eTag,
// se a gravação inicial falhou), então true também significa que a nota está salva.
async function uploadPdfSP(file, nota, aguardar = null) {
    if (!spAccessToken) return false;
    try {
        const safe = (nota.arquivo || file.name).replace(/[^a-zA-Z0-9_.-]/g, '_');
        const ts = nota.dataExtracao || Date.now();
        const spPath = `pdfs/${safe}_${ts}.pdf`;
        const buf = await file.arrayBuffer();
        const resp = await graphFetch(
            urlApproot(spPath, ':/content'),
            { method: 'PUT', headers: { 'Content-Type': 'application/pdf' }, body: buf }
        );
        if (resp.ok) {
            const item = await resp.json();
            nota._pdfSpPath = spPath;
            nota._pdfSpId   = item.id;
            // Salva nota atualizada com referência ao PDF
            if (aguardar) await aguardar;
            const res = await saveNotaSP(nota, { campos: ['_pdfSpPath', '_pdfSpId'] });
            return res.ok;
        }
        return false;
    } catch(e) { console.error('Erro ao upload PDF SP:', e); return false; }
}

// Lista notas/ seguindo a paginação do Graph (@odata.nextLink): sem isso só as
// primeiras 200 notas eram carregadas e as demais sumiam para os outros usuários.
async function listarArquivosNotasSP() {
    const arquivos = [];
    let url = urlApproot('notas', ':/children');
    let primeiraPagina = true;
    while (url) {
        const resp = await graphFetch(url);
        if (!resp.ok) {
            if (primeiraPagina) return null; // pasta ainda não existe
            throw new Error(`Falha ao listar notas (HTTP ${resp.status})`);
        }
        const data = await resp.json();
        arquivos.push(...(data.value || []));
        url = data['@odata.nextLink'] || null;
        primeiraPagina = false;
    }
    return arquivos;
}

async function loadNotasFromSharePoint() {
    if (!spAccessToken) return;
    try {
        const lista = await listarArquivosNotasSP();
        if (!lista) return;
        const files = lista.filter(f => f.name.endsWith('.json'));
        const notas = [];
        for (const f of files) {
            try {
                const r = await fetch(f['@microsoft.graph.downloadUrl']);
                if (r.ok) {
                    const nota = await r.json();
                    registrarVersaoNotaSP(nota, f.eTag, nota);
                    notas.push(nota);
                }
            } catch(e) {}
        }
        const spFileNames = new Set(notas.map(n => n._spFileName));
        extractedResults = [
            ...notas,
            ...extractedResults.filter(n => !n._spFileName || !spFileNames.has(n._spFileName))
        ];
        extractedResults.sort((a,b) => (a.dataExtracao||0) - (b.dataExtracao||0));
        const tbody = document.getElementById('resultsBody');
        tbody.innerHTML = '';
        extractedResults.forEach(r => renderRow(r));
        atualizarOpcoesTime();
        updateCoverageStats();
        console.log(`✅ ${notas.length} notas carregadas do SharePoint`);
    } catch(e) { console.error('Erro ao carregar notas SP:', e); }
}
