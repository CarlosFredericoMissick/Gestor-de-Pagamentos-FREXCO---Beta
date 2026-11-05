// ========================================
// MODAL — DIVIDIR RECEBÍVEL DE UM MOTORISTA ENTRE NOTAS (conciliação)
// A regra de divisão está em regras.js (dividirRecebivelEntreNotas). Aqui o usuário fixa o
// valor de cada nota; as alocações são gravadas nas próprias notas (alocacoesRecebivel) no
// SharePoint, uma por motorista e período de conciliação.
// ========================================

// { resultado, ids, vehicleId, notas, totalCentavos, periodo }
let _divisaoAtual = null;

function abrirDivisaoRecebivel(rowIdx) {
    const linha = _concilRows[rowIdx];
    if (!linha || !_concilCalculo || linha.resultado.compartilhados.length === 0) return;
    const { resultado } = linha;
    _divisaoAtual = { resultado, ids: resultado.compartilhados };

    const select = document.getElementById('divisaoMotorista');
    select.innerHTML = resultado.compartilhados
        .map(id => `<option value="${id}">${rotuloMotoristaDoResultado(resultado, id)}</option>`).join('');
    select.disabled = resultado.compartilhados.length === 1;

    const { periodo } = _concilCalculo.preparo;
    const fmt = (iso) => iso.split('-').reverse().join('/');
    document.getElementById('divisaoPeriodo').textContent = `Período ${fmt(periodo.inicio)} — ${fmt(periodo.fim)} · a divisão vale só para este período`;

    selecionarMotoristaDivisao(resultado.compartilhados[0]);
    document.getElementById('divisaoRecebivelModal').classList.remove('hidden');
    lucide.createIcons();
}

function fecharDivisaoRecebivel() {
    _divisaoAtual = null;
    document.getElementById('divisaoRecebivelModal').classList.add('hidden');
}

function selecionarMotoristaDivisao(vehicleId) {
    const { preparo } = _concilCalculo;
    const notas = preparo.notasPorMotorista.get(vehicleId) || [];
    const linhas = preparo.linhasPorMotorista.get(vehicleId) || [];
    Object.assign(_divisaoAtual, {
        vehicleId,
        notas,
        totalCentavos: linhas.reduce((s, l) => s + centavos(l.totalMotorista), 0),
        periodo: preparo.periodo
    });
    document.getElementById('divisaoTotal').textContent =
        `${formatarMoeda(_divisaoAtual.totalCentavos / 100)} · ${linhas.length} rota(s)`;

    document.getElementById('divisaoNotasBody').innerHTML = notas.map((n, i) => {
        const atual = alocacaoDaNota(n, vehicleId, preparo.periodo);
        return `<tr>
            <td class="px-3 py-2 truncate max-w-[220px]" title="${n.arquivo}">${n.arquivo}</td>
            <td class="px-3 py-2 text-center text-gray-500">${n.dataEmissao || '-'}</td>
            <td class="px-3 py-2 text-right font-mono">
                ${formatarMoeda(n.vLiquido || 0)}
                <button type="button" onclick="preencherAlocacaoComValorDaNota(${i})" class="ml-1 text-[10px] text-blue-600 hover:underline" title="Alocar o valor desta nota">usar</button>
            </td>
            <td class="px-3 py-2 text-right">
                <input type="number" step="0.01" min="0" data-divisao-idx="${i}" value="${atual ? atual.valor : ''}" placeholder="restante"
                    oninput="atualizarResumoDivisao()"
                    class="w-32 px-2 py-1 border border-gray-300 rounded-lg text-sm text-right font-mono focus:ring-2 focus:ring-gray-400 outline-none">
            </td>
            <td class="px-3 py-2 text-gray-600" id="divisaoResultado-${i}"></td>
        </tr>`;
    }).join('');
    atualizarResumoDivisao();
}

function preencherAlocacaoComValorDaNota(i) {
    const input = document.querySelector(`[data-divisao-idx="${i}"]`);
    if (input) input.value = (_divisaoAtual.notas[i].vLiquido || 0).toFixed(2);
    atualizarResumoDivisao();
}

// { valores, invalidos }: valor de cada nota em centavos (null = em branco, recebe o restante)
// e os índices com conteúdo que não é número. Campo type=number inválido chega como '' no
// value; validity.badInput distingue isso de campo realmente vazio.
function lerValoresDivisao() {
    const invalidos = [];
    const valores = _divisaoAtual.notas.map((_, i) => {
        const input = document.querySelector(`[data-divisao-idx="${i}"]`);
        const texto = (input?.value ?? '').trim();
        if (input?.validity?.badInput) { invalidos.push(i); return null; }
        if (texto === '') return null;
        const valor = Number(texto.replace(',', '.'));
        if (!Number.isFinite(valor)) { invalidos.push(i); return null; }
        return centavos(valor);
    });
    return { valores, invalidos };
}

function mensagemValidacaoDivisao(validacao, invalidos = []) {
    const total = formatarMoeda(_divisaoAtual.totalCentavos / 100);
    if (invalidos.length > 0) {
        const notas = invalidos.map(i => _divisaoAtual.notas[i].arquivo).join(', ');
        return `Valor inválido em: ${notas}. Corrija (use só números, ex.: 1500,50) ou deixe em branco para receber o restante.`;
    }
    switch (validacao.status) {
        case 'negativo':
            return 'Valores alocados não podem ser negativos.';
        case 'excede':
            return `A soma alocada (${formatarMoeda(validacao.somaCentavos / 100)}) passa do total do motorista no período (${total}). Reduza os valores.`;
        case 'semRestante':
            return 'Deixe pelo menos uma nota em branco: ela recebe o restante do recebível.';
        default:
            return null;
    }
}

function atualizarResumoDivisao() {
    if (!_divisaoAtual) return;
    const { valores, invalidos } = lerValoresDivisao();
    const v = validarAlocacoes(_divisaoAtual.totalCentavos, valores);
    const restante = formatarMoeda(v.restanteCentavos / 100);
    const total = formatarMoeda(_divisaoAtual.totalCentavos / 100);

    valores.forEach((valor, i) => {
        const el = document.getElementById(`divisaoResultado-${i}`);
        if (!el) return;
        if (invalidos.includes(i)) el.innerHTML = '<span class="text-red-600 font-semibold">valor inválido</span>';
        else if (invalidos.length > 0) el.textContent = '';
        else if (valor !== null) el.textContent = 'valor alocado';
        else if (v.status === 'ok') el.innerHTML = `<span class="text-green-700 font-semibold">recebe o restante: ${restante}</span>`;
        else if (v.status === 'restanteMultiplo') el.innerHTML = `<span class="text-amber-700">sem alocação: recebe o total (${total})</span>`;
        else el.textContent = '';
    });

    const resumo = document.getElementById('divisaoResumo');
    const erro = mensagemValidacaoDivisao(v, invalidos);
    let texto, classe;
    if (erro) {
        texto = erro; classe = 'bg-red-50 text-red-700 border border-red-200';
    } else if (v.status === 'restanteMultiplo') {
        texto = `Alocado ${formatarMoeda(v.somaCentavos / 100)} de ${total}. Há ${v.semAlocacao} notas em branco: aloque as demais notas deste motorista (até lá o total entra em cada uma delas).`;
        classe = 'bg-amber-50 text-amber-800 border border-amber-200';
    } else {
        texto = v.somaCentavos === 0
            ? 'Sem alocação: o total do motorista entra em cada nota. Informe valores para dividir.'
            : `Alocado ${formatarMoeda(v.somaCentavos / 100)} de ${total}. Restante ${restante} vai para a nota em branco.`;
        classe = 'bg-gray-50 text-gray-700 border border-gray-200';
    }
    resumo.className = `text-xs rounded-lg px-3 py-2 ${classe}`;
    resumo.textContent = texto;
}

async function salvarDivisaoRecebivel() {
    if (!_divisaoAtual) return;
    const { valores, invalidos } = lerValoresDivisao();
    // Valor inválido sempre bloqueia; sem nenhum valor = limpar; senão as regras de bloqueio valem
    if (invalidos.length > 0 || valores.some(v => v !== null)) {
        const erro = mensagemValidacaoDivisao(validarAlocacoes(_divisaoAtual.totalCentavos, valores), invalidos);
        if (erro) { alert(`Não foi possível salvar a divisão.\n\n${erro}`); return; }
    }
    await gravarDivisao(valores);
}

async function limparDivisaoRecebivel() {
    if (!_divisaoAtual) return;
    if (!confirm('Limpar a divisão do recebível deste motorista neste período? O total volta a entrar em cada nota.')) return;
    await gravarDivisao(_divisaoAtual.notas.map(() => null));
}

async function gravarDivisao(valores) {
    const { vehicleId, notas, periodo } = _divisaoAtual;
    const prestador = findPrestadorByVehicleId(vehicleId);
    const vehicleIdGravado = prestador ? prestador.vehicleId : vehicleId;
    const agora = new Date().toISOString();
    const editadoPor = spAccount?.name || spAccount?.username || null;

    const alteracoes = [];
    notas.forEach((nota, i) => {
        const anterior = alocacaoDaNota(nota, vehicleId, periodo);
        const nova = valores[i] === null ? null : {
            vehicleId: vehicleIdGravado, valor: valores[i] / 100,
            periodoInicio: periodo.inicio, periodoFim: periodo.fim, editadoEm: agora, editadoPor
        };
        if ((anterior?.valor ?? null) === (nova?.valor ?? null)) return;
        alteracoes.push({ nota, nova });
    });

    const btn = document.getElementById('btnSalvarDivisao');
    if (btn) { btn.disabled = true; btn.textContent = 'Salvando...'; }
    const salvou = await persistirAlocacoes(alteracoes, vehicleId, periodo);
    if (btn) { btn.disabled = false; btn.textContent = 'Salvar'; }
    if (!salvou) return; // modal continua aberto para nova tentativa
    fecharDivisaoRecebivel();
    executarConciliacao();
}

// Grava a alocação de um motorista em cada nota alterada (If-Match; em conflito só a entrada
// deste motorista/período é reaplicada sobre a versão remota). Retorna true se todas gravaram.
async function persistirAlocacoes(alteracoes, vehicleId, periodo) {
    if (alteracoes.length === 0) return true;
    const aplicar = (nota, nova) => {
        nota.alocacoesRecebivel = substituirAlocacao(nota.alocacoesRecebivel, vehicleId, periodo, nova);
    };
    if (!useSharePoint) {
        alteracoes.forEach(({ nota, nova }) => aplicar(nota, nova));
        alert('⚠️ SharePoint não conectado: a divisão ficou só neste navegador e será perdida ao recarregar.\n\nConecte o SharePoint e salve novamente para compartilhar com os outros usuários.');
        return true;
    }

    const falhas = [], conflitos = [];
    for (const { nota, nova } of alteracoes) {
        const anterior = nota.alocacoesRecebivel;
        aplicar(nota, nova);
        const res = await saveNotaSP(nota, {
            campos: ['alocacoesRecebivel'],
            mesclar: (remoto, base) => {
                const valorRemoto = alocacaoDaNota(remoto, vehicleId, periodo)?.valor ?? null;
                const valorBase   = alocacaoDaNota(base, vehicleId, periodo)?.valor ?? null;
                const conflito = valorRemoto !== valorBase && valorRemoto !== (nova?.valor ?? null);
                return {
                    valores: { alocacoesRecebivel: substituirAlocacao(remoto.alocacoesRecebivel, vehicleId, periodo, nova) },
                    conflitos: conflito ? ['alocacoesRecebivel'] : []
                };
            }
        });
        if (!res.ok) {
            if (anterior === undefined) delete nota.alocacoesRecebivel; else nota.alocacoesRecebivel = anterior;
            falhas.push(`${nota.arquivo} (${res.mensagem})`);
        } else if (res.conflitos.length > 0) {
            conflitos.push(nota.arquivo);
        }
    }
    if (conflitos.length > 0) {
        alert(`⚠️ Outro usuário alterou a divisão deste motorista enquanto você editava, nas notas:\n${conflitos.join('\n')}\n\nFicou a sua versão.`);
    }
    if (falhas.length > 0) {
        alert(`❌ A divisão NÃO foi salva no SharePoint em ${falhas.length} nota(s):\n${falhas.join('\n')}\n\nAs demais foram salvas, então a divisão pode estar incompleta. Salve de novo.`);
        return false;
    }
    return true;
}
