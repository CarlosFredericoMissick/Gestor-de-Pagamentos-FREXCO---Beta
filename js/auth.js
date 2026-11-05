// ========================================
// AUTENTICAÇÃO MSAL + CONEXÃO COM O SHAREPOINT
// ========================================

const msalInstance = new msal.PublicClientApplication(msalConfig);

// Token sempre renovado antes de cada chamada ao Graph: o MSAL devolve o token em cache
// se ainda for válido e renova silenciosamente se expirou (antes, gravações feitas entre
// a expiração e o próximo ciclo do polling de 5 min falhavam com 401 sem aviso).
async function obterTokenSP(forcarRenovacao = false) {
    if (!spAccount) return spAccessToken;
    try {
        const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: SP_SCOPES,
            account: spAccount,
            forceRefresh: forcarRenovacao
        });
        spAccessToken = tokenResponse.accessToken;
        spExigeReconexao = false;
    } catch (e) {
        console.warn('Não foi possível renovar o token do SharePoint:', e);
        // Sessão que só volta com login interativo: mensagemErroSP orienta a reconectar
        spExigeReconexao = (msal.InteractionRequiredAuthError && e instanceof msal.InteractionRequiredAuthError)
            || /interaction_required|login_required|consent_required/.test(e?.errorCode || '');
    }
    return spAccessToken;
}

async function toggleSharePoint() {
    if (useSharePoint) {
        useSharePoint = false;
        spAccessToken = null;
        spAccount = null;
        document.getElementById('spIndicator').classList.add('hidden');
        document.querySelector('#btnSharePoint span').textContent = 'Conectar SharePoint';
        const icon = document.getElementById('iconSharePoint');
        if (icon) { icon.classList.remove('spin'); }
        await loadPrestadoresLocal();
        reprocessExtractedResults();
    } else {
        // Remove banner de sessão expirada imediatamente ao tentar reconectar
        document.getElementById('spDisconnectedBanner')?.remove();

        const btn  = document.getElementById('btnSharePoint');
        const icon = document.getElementById('iconSharePoint');
        const span = btn?.querySelector('span');
        const setPhase = (txt) => { if (span) span.textContent = txt; };

        btn.disabled = true;
        if (icon) icon.classList.add('spin');
        setPhase('Autenticando...');

        try {
            await msalInstance.initialize();
            const loginResponse = await msalInstance.loginPopup({
                scopes: SP_SCOPES
            });
            spAccount = loginResponse.account;
            const tokenResponse = await msalInstance.acquireTokenSilent({
                scopes: SP_SCOPES,
                account: spAccount
            });
            spAccessToken = tokenResponse.accessToken;
            useSharePoint = true;

            document.getElementById('spIndicator').classList.remove('hidden');
            setPhase('Carregando prestadores...');

            await loadPrestadoresFromSharePoint();
            setPhase('Carregando notas...');

            await loadNotasFromSharePoint();
            setPhase('Carregando recebível...');

            await loadRecebivelFromSharePoint();
            reprocessExtractedResults();

            // Polling de expiração de token a cada 5 minutos
            startTokenExpiryCheck();

        } catch (error) {
            console.error('Erro login SharePoint:', error);
            useSharePoint = false;
            spAccessToken = null;
            if (error.errorCode !== 'user_cancelled') {
                alert('Erro ao conectar SharePoint: ' + (error.message || error));
            }
        } finally {
            btn.disabled = false;
            if (icon) icon.classList.remove('spin');
            if (useSharePoint) {
                setPhase('Desconectar');
            } else {
                setPhase('Conectar SharePoint');
            }
        }
    }
}

// Verifica silenciosamente se o token ainda é válido a cada 5 minutos
let _tokenCheckInterval = null;
function startTokenExpiryCheck() {
    if (_tokenCheckInterval) clearInterval(_tokenCheckInterval);
    _tokenCheckInterval = setInterval(async () => {
        if (!useSharePoint || !spAccount) return;
        try {
            const tokenResponse = await msalInstance.acquireTokenSilent({
                scopes: SP_SCOPES,
                account: spAccount
            });
            spAccessToken = tokenResponse.accessToken; // renova silenciosamente
        } catch(e) {
            // Token expirou e não foi possível renovar silenciosamente
            clearInterval(_tokenCheckInterval);
            _tokenCheckInterval = null;
            spAccessToken = null;
            useSharePoint = false;
            spAccount = null;

            // Atualizar UI para estado desconectado
            document.getElementById('spIndicator').classList.add('hidden');
            document.querySelector('#btnSharePoint span').textContent = 'Conectar SharePoint';

            // Notificação visual no header
            showSpDisconnectedBanner();
        }
    }, 5 * 60 * 1000); // a cada 5 minutos
}

function showSpDisconnectedBanner() {
    // Remove banner anterior se existir
    const old = document.getElementById('spDisconnectedBanner');
    if (old) old.remove();

    const banner = document.createElement('div');
    banner.id = 'spDisconnectedBanner';
    banner.className = 'fixed top-0 left-0 right-0 z-[100] bg-amber-500 text-white text-sm font-semibold flex items-center justify-between px-6 py-3 shadow-lg';
    banner.innerHTML = `
        <div class="flex items-center gap-2">
            <svg xmlns="http://www.w3.org/2000/svg" class="w-5 h-5 flex-shrink-0" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
            <span>Sessão do SharePoint expirada. Suas próximas alterações <strong>não serão salvas</strong> até reconectar.</span>
        </div>
        <div class="flex items-center gap-3">
            <button onclick="toggleSharePoint()" class="bg-white text-amber-600 px-3 py-1 rounded-lg text-xs font-bold hover:bg-amber-50 transition-colors">Reconectar</button>
            <button onclick="document.getElementById('spDisconnectedBanner').remove()" class="text-white/80 hover:text-white text-lg leading-none">✕</button>
        </div>`;
    document.body.prepend(banner);
}

async function refreshSharePoint() {
    if (!useSharePoint || !spAccessToken) return;
    const btn  = document.getElementById('btnAtualizar');
    const icon = btn?.querySelector('i');
    const text = document.getElementById('btnAtualizarText');
    if (btn)  btn.disabled = true;
    if (icon) icon.classList.add('spin');
    if (text) text.textContent = 'Atualizando...';
    try {
        await loadPrestadoresFromSharePoint();
        await loadNotasFromSharePoint();
        await loadRecebivelFromSharePoint();
        reprocessExtractedResults();
        console.log('✅ SharePoint atualizado');
    } catch(e) {
        alert('Erro ao atualizar dados do SharePoint.');
    } finally {
        if (btn)  btn.disabled = false;
        if (icon) icon.classList.remove('spin');
        if (text) text.textContent = 'Atualizar';
    }
}
