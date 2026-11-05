// ========================================
// INDEXEDDB — prestadores quando o SharePoint não está conectado + migrações
// ========================================

// Versão 2: a versão anterior podia deixar o banco na v1 SEM a store 'prestadores' (a
// migração abria o banco sem onupgradeneeded e criava um banco vazio no primeiro acesso),
// e daí em diante toda abertura dava "One of the specified object stores was not found".
// Subir a versão faz o onupgradeneeded rodar nesses navegadores e criar a store que falta;
// bancos v1 com a store e dados passam para a v2 sem tocar nos dados.
const PRESTADORES_DB_NOME = 'FrexcoPrestadoresDB';
const PRESTADORES_DB_VERSAO = 2;

async function initPrestadoresDB() {
    return new Promise((resolve, reject) => {
        const req = indexedDB.open(PRESTADORES_DB_NOME, PRESTADORES_DB_VERSAO);
        req.onsuccess = () => {
            prestadoresDB = req.result;
            // Outra aba com versão mais nova: libera o banco para o upgrade dela
            prestadoresDB.onversionchange = () => { prestadoresDB.close(); prestadoresDB = null; };
            resolve();
        };
        req.onerror = () => reject(req.error);
        req.onblocked = () => {
            alert('O Gestor de Pagamentos está aberto em outra aba com uma versão anterior.\n\nFeche as outras abas do app e recarregue esta página.');
        };
        req.onupgradeneeded = (e) => {
            const db = e.target.result;
            if (!db.objectStoreNames.contains('prestadores')) {
                const store = db.createObjectStore('prestadores', { keyPath: 'id', autoIncrement: true });
                store.createIndex('cnpj', 'cnpj', { unique: false });
            }
        };
    });
}

// Abre um banco só se ele já existir (qualquer versão), sem criar banco vazio: se o open
// precisaria criar o banco, a transação de upgrade é abortada e o resultado é null.
// Quem chama deve fechar o banco (db.close()) para não bloquear upgrades futuros.
function abrirBancoExistente(nome) {
    return new Promise(resolve => {
        let criando = false;
        const req = indexedDB.open(nome);
        req.onupgradeneeded = (e) => {
            criando = true;
            e.target.transaction.abort();
        };
        req.onsuccess = () => {
            if (criando) { req.result.close(); resolve(null); return; }
            resolve(req.result);
        };
        req.onerror = (e) => { e?.preventDefault?.(); resolve(null); };
        req.onblocked = () => resolve(null);
    });
}

// Lê todos os registros de uma store de um banco já aberto (vazio se a store não existir).
function lerTodosDaStore(db, nomeStore) {
    return new Promise(resolve => {
        if (!db || !db.objectStoreNames.contains(nomeStore)) { resolve([]); return; }
        const r = db.transaction([nomeStore], 'readonly').objectStore(nomeStore).getAll();
        r.onsuccess = () => resolve(r.result || []);
        r.onerror = () => resolve([]);
    });
}

async function loadPrestadoresLocal() {
    if (!prestadoresDB) await initPrestadoresDB();
    return new Promise((resolve, reject) => {
        const tx = prestadoresDB.transaction(['prestadores'], 'readonly');
        const req = tx.objectStore('prestadores').getAll();
        req.onsuccess = () => {
            prestadoresList = migratePrestadoresData(req.result || []);
            sortPrestadores();
            updatePrestadoresUI();
            resolve();
        };
        req.onerror = () => reject(req.error);
    });
}

async function savePrestadorLocal(p) {
    if (!prestadoresDB) await initPrestadoresDB();
    return new Promise((resolve, reject) => {
        const tx = prestadoresDB.transaction(['prestadores'], 'readwrite');
        const store = tx.objectStore('prestadores');
        const req = p.id ? store.put(p) : store.add(p);
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
    });
}

async function deletePrestadorLocal(id) {
    if (!prestadoresDB) await initPrestadoresDB();
    return new Promise((resolve, reject) => {
        const tx = prestadoresDB.transaction(['prestadores'], 'readwrite');
        const req = tx.objectStore('prestadores').delete(id);
        req.onsuccess = () => resolve();
        req.onerror = () => reject(req.error);
    });
}

function migratePrestadoresData(data) {
    if (!data || data.length === 0) return [];
    return data.map(p => {
        if (Array.isArray(p.cnpjs)) return p;
        const cnpjs = [];
        if (p.cnpj) cnpjs.push({ cnpj: p.cnpj, chavePix: p.chavePix || '' });
        return {
            id: p.id || (Date.now() + Math.random()),
            vehicleId: p.vehicleId || ('M' + (p.id || Date.now()).toString()),
            nome: p.nome || p.nomeMotorista || '',
            cd: p.cd || '',
            time: p.time || '',
            inUse: p.inUse !== undefined ? p.inUse : true,
            cnpjs,
            contrato: p.contrato || null,
            fromSharePoint: p.fromSharePoint || false,
            dataAdicao: p.dataAdicao || new Date().toISOString()
        };
    });
}

// Migra o banco legado FrexcoPixDB para prestadores, se ainda não houver prestadores locais.
// Só lê bancos que já existem (abrirBancoExistente): o primeiro acesso não cria banco vazio.
async function migrateFromLegacyDBs() {
    const dbPrestadores = await abrirBancoExistente(PRESTADORES_DB_NOME);
    const prestadoresExistentes = await lerTodosDaStore(dbPrestadores, 'prestadores');
    dbPrestadores?.close();
    if (prestadoresExistentes.length > 0) return;

    const dbPix = await abrirBancoExistente('FrexcoPixDB');
    const legacyPix = await lerTodosDaStore(dbPix, 'motoristas');
    dbPix?.close();

    if (legacyPix.length === 0) return;
    for (const p of legacyPix) {
        const novo = {
            vehicleId: 'M' + (p.id || Date.now() + Math.random().toString(36).substr(2,4)), nome: p.nomeMotorista, cd: '', time: '', inUse: true,
            cnpjs: p.chavePix ? [{ cnpj: '', chavePix: p.chavePix }] : [],
            contrato: null, fromSharePoint: false,
            dataAdicao: p.dataAdicao || new Date().toISOString()
        };
        await savePrestadorLocal(novo);
    }
}
