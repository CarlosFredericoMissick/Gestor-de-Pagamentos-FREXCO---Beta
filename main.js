// ========================================
// INICIALIZAÇÃO — listeners de upload e carga inicial (último script da página)
// ========================================

// Inputs de importação (prestadores e recebível)
document.addEventListener('DOMContentLoaded', () => {
    const baseInput = document.getElementById('baseMotoristaInput');
    if (baseInput) baseInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        await importarBaseMotoristasExcel(file);
        e.target.value = '';
    });

    const cnpjPixInput = document.getElementById('cnpjPixInput');
    if (cnpjPixInput) cnpjPixInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        await importarCnpjPixExcel(file);
        e.target.value = '';
    });

    const contratosZipInput = document.getElementById('contratosZipInput');
    if (contratosZipInput) contratosZipInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        await importarContratosZip(file);
        e.target.value = '';
    });

    const recInput = document.getElementById('recebivelInput');
    if (recInput) recInput.addEventListener('change', async (e) => {
        const file = e.target.files[0]; if (!file) return;
        await importarRecebivelExcel(file);
        document.getElementById('recebivelStatus').textContent =
            `Recebível carregado: ${recebivelData.length} registros`;
        e.target.value = '';
    });
});

// Área de upload de PDFs
const dropZone  = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('bg-primary/5', 'border-primary'); });
dropZone.addEventListener('dragleave', () => { dropZone.classList.remove('bg-primary/5', 'border-primary'); });
dropZone.addEventListener('drop', (e) => { e.preventDefault(); dropZone.classList.remove('bg-primary/5', 'border-primary'); handleFiles(e.dataTransfer.files); });
fileInput.addEventListener('change', (e) => handleFiles(e.target.files));

(async function init() {
    try {
        await migrateFromLegacyDBs();
        await loadPrestadoresLocal();
        // Notas vivem no SharePoint — serão carregadas ao conectar
        console.log('✅ Sistema inicializado');
    } catch (error) {
        console.error('Erro ao inicializar:', error);
        alert('Erro ao inicializar sistema');
    }
})();
