// ========================================
// UTILITÁRIOS DE UI
// ========================================

function formatarMoeda(v) {
    return (v || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

// Literal JS seguro para argumentos de onclick="..." gerados em template:
// preserva o tipo (número continua número, 'L3'/'M5' viram string) e escapa aspas.
function argJs(valor) {
    return JSON.stringify(valor ?? null).replace(/&/g, '&amp;').replace(/"/g, '&quot;');
}

// Loading helpers
function setButtonLoading(btnId, iconId, textId, loadingText) {
    const btn  = document.getElementById(btnId);
    const icon = document.getElementById(iconId);
    const text = document.getElementById(textId);
    if (btn)  btn.disabled = true;
    if (icon) icon.classList.add('spin');
    if (text) text.textContent = loadingText;
}
function resetButton(btnId, iconId, textId, originalText) {
    const btn  = document.getElementById(btnId);
    const icon = document.getElementById(iconId);
    const text = document.getElementById(textId);
    if (btn)  btn.disabled = false;
    if (icon) icon.classList.remove('spin');
    if (text) text.textContent = originalText;
}
