// ========================================
// INDEXEDDB — prestadores quando o SharePoint não está conectado + migrações
// ========================================

async function initPrestadoresDB() {
    return new Promise((resolve, reject) => {
        const req = indexedDB.open('FrexcoPrestadoresDB', 1);
        req.onsuccess = () => { prestadoresDB = req.result; resolve(); };
        req.onerror = () => reject(req.error);
        req.onupgradeneeded = (e) => {
            const db = e.target.result;
            if (!db.objectStoreNames.contains('prestadores')) {
                const store = db.createObjectStore('prestadores', { keyPath: 'id', autoIncrement: true });
                store.createIndex('cnpj', 'cnpj', { unique: false });
            }
        };
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

async function migrateFromLegacyDBs() {
    const hasPrestadores = await new Promise(resolve => {
        const req = indexedDB.open('FrexcoPrestadoresDB', 1);
        req.onsuccess = () => {
            const db = req.result;
            if (!db.objectStoreNames.contains('prestadores')) { resolve(false); return; }
            const tx = db.transaction(['prestadores'], 'readonly');
            const r = tx.objectStore('prestadores').count();
            r.onsuccess = () => resolve(r.result > 0);
            r.onerror = () => resolve(false);
        };
        req.onerror = () => resolve(false);
    });
    if (hasPrestadores) return;

    const legacyPix = await new Promise(resolve => {
        const req = indexedDB.open('FrexcoPixDB', 1);
        req.onsuccess = () => {
            const db = req.result;
            if (!db.objectStoreNames.contains('motoristas')) { resolve([]); return; }
            const tx = db.transaction(['motoristas'], 'readonly');
            const r = tx.objectStore('motoristas').getAll();
            r.onsuccess = () => resolve(r.result || []);
            r.onerror = () => resolve([]);
        };
        req.onerror = () => resolve([]);
    });

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
