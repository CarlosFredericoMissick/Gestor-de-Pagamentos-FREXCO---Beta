// ========================================
// REGRAS DE NEGÓCIO PURAS
// Sem DOM e sem estado global: recebem prestadores/recebível por parâmetro.
// Carregado no navegador como script clássico e no Node pelos testes (testes/).
// ========================================

function normalizeCnpj(cnpj) { return (cnpj || '').replace(/\D/g, ''); }

// Chave de comparação de motorista: driver_id numérico vira "123" (aceita 123, "123", "0123");
// IDs de líder/manual ("L3", "M5") ficam como estão.
function idMotorista(vehicleId) {
    const s = String(vehicleId ?? '').trim();
    return /^\d+$/.test(s) ? String(parseInt(s, 10)) : s;
}

function buscarPrestadorPorCnpj(prestadores, cnpj) {
    if (!cnpj) return null;
    const norm = normalizeCnpj(cnpj);
    if (!norm) return null;
    // Apenas motoristas EM USO participam do cruzamento por CNPJ
    return prestadores.find(p =>
        p.inUse !== false &&
        Array.isArray(p.cnpjs) && p.cnpjs.some(c => normalizeCnpj(c.cnpj) === norm)
    ) || null;
}

function buscarPrestadorPorVehicleId(prestadores, vehicleId) {
    if (!vehicleId) return null;
    const vid = parseInt(vehicleId);
    // Busca direta pelo driverId (campo vehicleId após migração)
    const direto = prestadores.find(p => p.vehicleId == vehicleId);
    if (direto) return direto;
    // Fallback: busca no array vehicles para normalização de notas com vehicle_ids antigos
    return prestadores.find(p =>
        Array.isArray(p.vehicles) && p.vehicles.some(v => parseInt(v.vehicleId) === vid)
    ) || null;
}

// ========================================
// VÍNCULO DE MOTORISTAS NA NOTA (automático x manual)
//
// nota.motoristasManual:
//   true  → lista escolhida pelo usuário no modal "Motoristas na Nota"; nunca é
//           sobrescrita pelo reprocessamento (só IDs desatualizados são corrigidos).
//   false → lista automática, derivada do dono do CNPJ; acompanha o cadastro.
//   ausente → nota gravada por versão antiga; classificada por migrarNotaLegada.
// ========================================

// Lista explícita de uma nota de líder: o líder (comissão), os motoristasVinculados e os
// demais motoristas já na lista, sem repetir. Reproduz o que a regra antiga somava
// (comissão + vinculados + lista) e passa a ficar visível e editável no modal.
function listaExplicitaDoLider(lider, lista) {
    const resultado = [];
    const vistos = new Set();
    const add = (m) => {
        const id = idMotorista(m.vehicleId);
        if (!id || vistos.has(id)) return;
        vistos.add(id);
        resultado.push({ vehicleId: m.vehicleId, nome: m.nome });
    };
    add({ vehicleId: lider.vehicleId, nome: lider.nome });
    (lider.motoristasVinculados || []).forEach(add);
    (lista || []).forEach(add);
    return resultado;
}

// Notas sem CNPJ de líder: na gravação automática a lista sempre era [responsável] ou os
// motoristasVinculados do líder responsável; qualquer outra combinação veio do modal.
function inferirMotoristasManual(nota, prestadores) {
    const lista = Array.isArray(nota.motoristasNaNota) ? nota.motoristasNaNota : [];
    if (lista.length === 0) return false;
    const resp = nota.responsavel;
    if (!resp) return true;
    const ids = new Set(lista.map(m => idMotorista(m.vehicleId)));
    if (ids.size === 1 && ids.has(idMotorista(resp.vehicleId))) return false;
    const lider = buscarPrestadorPorVehicleId(prestadores, resp.vehicleId);
    if (lider && lider.tipo === 'lider') {
        const vinculados = new Set((lider.motoristasVinculados || []).map(m => idMotorista(m.vehicleId)));
        if ([...ids].every(id => vinculados.has(id))) return false;
    }
    return true;
}

// Classifica uma nota gravada por versão antiga (sem motoristasManual) sem mudar o valor
// devido que a regra antiga calculava. Nota de CNPJ de líder: a regra antiga sempre somava
// comissão + vinculados atuais + lista; se a lista ainda cabe no automático (líder e/ou
// vinculados) fica automática, senão vira manual com a lista explícita do líder.
// Retorna { migrada, listaAlterada }.
function migrarNotaLegada(nota, prestadores) {
    if (nota.motoristasManual !== undefined) return { migrada: false, listaAlterada: false };
    const lista = Array.isArray(nota.motoristasNaNota) ? nota.motoristasNaNota : [];
    const dono = buscarPrestadorPorCnpj(prestadores, nota.cnpj);
    if (dono && dono.tipo === 'lider') {
        const automaticos = new Set([dono, ...(dono.motoristasVinculados || [])].map(m => idMotorista(m.vehicleId)));
        if (lista.every(m => automaticos.has(idMotorista(m.vehicleId)))) {
            nota.motoristasManual = false;
            return { migrada: true, listaAlterada: false };
        }
        nota.motoristasManual = true;
        nota.motoristasNaNota = listaExplicitaDoLider(dono, lista);
        return { migrada: true, listaAlterada: true };
    }
    nota.motoristasManual = inferirMotoristasManual(nota, prestadores);
    return { migrada: true, listaAlterada: false };
}

// Recalcula responsável e motoristas da nota a partir do cadastro atual (muta a nota).
// Retorna o resultado de migrarNotaLegada (para persistir a classificação de notas antigas).
function reconciliarMotoristasDaNota(nota, prestadores) {
    const responsavelAnterior = nota.responsavel;
    const migracao = migrarNotaLegada(nota, prestadores);

    // Atualizar responsável pelo CNPJ
    const prestadorResp = buscarPrestadorPorCnpj(prestadores, nota.cnpj);
    nota.responsavel = prestadorResp
        ? { vehicleId: prestadorResp.vehicleId, nome: prestadorResp.nome }
        : null;

    if (!nota.motoristasManual) {
        // ── CASO 1: lista vazia → preencher automaticamente ──
        if ((!nota.motoristasNaNota || nota.motoristasNaNota.length === 0) && nota.responsavel) {
            if (prestadorResp?.tipo === 'lider' && (prestadorResp.motoristasVinculados || []).length > 0) {
                nota.motoristasNaNota = prestadorResp.motoristasVinculados.map(m => ({...m}));
            } else {
                nota.motoristasNaNota = [{ ...nota.responsavel }];
            }
        }
        // ── CASO 2: 1 motorista automático, responsável mudou → corrigir ──
        // (exceto líderes, que têm múltiplos motoristas vinculados)
        else if (
            prestadorResp &&
            prestadorResp.tipo !== 'lider' &&
            Array.isArray(nota.motoristasNaNota) &&
            nota.motoristasNaNota.length === 1 &&
            String(nota.motoristasNaNota[0].vehicleId) !== String(prestadorResp.vehicleId)
        ) {
            nota.motoristasNaNota = [{ vehicleId: prestadorResp.vehicleId, nome: prestadorResp.nome }];
        }
        // ── CASO 3: responsável foi EXCLUÍDO da base → remove só o vínculo dele ──
        else if (!prestadorResp && responsavelAnterior && Array.isArray(nota.motoristasNaNota)) {
            const vidAnterior = String(responsavelAnterior.vehicleId);
            nota.motoristasNaNota = nota.motoristasNaNota.filter(
                m => String(m.vehicleId) !== vidAnterior
            );
        }
    }

    // ── Normalização: atualiza IDs obsoletos e remove vínculos para prestadores excluídos ──
    // Em lista manual o vínculo é mantido mesmo sem cadastro (escolha do usuário; o
    // recebível continua sendo buscado pelo driver_id).
    if (Array.isArray(nota.motoristasNaNota)) {
        nota.motoristasNaNota = nota.motoristasNaNota
            .map(m => {
                const p = buscarPrestadorPorVehicleId(prestadores, m.vehicleId);
                if (!p) return nota.motoristasManual ? m : null;
                if (parseInt(p.vehicleId) !== parseInt(m.vehicleId)) {
                    return { vehicleId: p.vehicleId, nome: p.nome }; // ID desatualizado → atualiza
                }
                return m;
            })
            .filter(Boolean);
    }
    return migracao;
}

// ========================================
// VALOR DEVIDO DA NOTA (conciliação)
//
// Regra única usada pela tabela, pelo detalhe e pelas exportações:
// 1. Motoristas efetivos = nota.motoristasNaNota. Lista manual é a fonte de verdade,
//    mesmo vazia. Lista automática vazia cai para o responsável e, por fim, para o
//    dono do CNPJ.
// 2. Líderes na nota = registros tipo 'lider' presentes nos motoristas efetivos. Em nota
//    automática cujo CNPJ é de um líder, esse líder também conta (é o emissor e a lista
//    veio dos motoristasVinculados dele). Em nota manual o líder só conta se estiver na lista.
// 3. Motoristas somados (Total Motorista do recebível no período), sem repetição:
//    - nota com líder: todos os motoristas efetivos que não são líderes; em nota
//      automática de líder, também os motoristasVinculados atuais do líder;
//    - nota sem líder: só motoristas com total positivo no período.
// 4. Comissão: coluna Comissão do recebível do time de cada líder presente, uma vez por
//    nota. Na conciliação, a comissão do time vai só para a nota mais recente em que o
//    líder conta (ver POLÍTICA DE COMISSÃO abaixo).
// ========================================

function normalizarPeriodo(dataInicio, dataFim) {
    return {
        ini: new Date(dataInicio.getFullYear(), dataInicio.getMonth(), dataInicio.getDate()),
        fim: new Date(dataFim.getFullYear(),    dataFim.getMonth(),    dataFim.getDate())
    };
}

function normalizarTime(time) { return (time || '').toLowerCase().trim(); }

function obterMotoristasEfetivos(nota, prestadores) {
    const lista = Array.isArray(nota.motoristasNaNota) ? nota.motoristasNaNota : null;
    if (lista && (lista.length > 0 || nota.motoristasManual)) return lista;
    if (nota.responsavel) return [nota.responsavel];
    const p = nota.cnpj ? buscarPrestadorPorCnpj(prestadores, nota.cnpj) : null;
    return p ? [{ vehicleId: p.vehicleId, nome: p.nome }] : [];
}

function obterLideresDaNota(nota, motoristas, prestadores) {
    const lideres = [];
    const add = (p) => { if (p && p.tipo === 'lider' && !lideres.includes(p)) lideres.push(p); };
    if (!nota.motoristasManual) add(buscarPrestadorPorCnpj(prestadores, nota.cnpj));
    motoristas.forEach(m => add(buscarPrestadorPorVehicleId(prestadores, m.vehicleId)));
    return lideres;
}

function dataEmissaoParaDate(dataEmissao) {
    const m = String(dataEmissao || '').match(/(\d{2})\/(\d{2})\/(\d{4})/);
    return m ? new Date(+m[3], +m[2] - 1, +m[1]) : null;
}

// ========================================
// POLÍTICA DE COMISSÃO NA CONCILIAÇÃO
// Uma política responde, para cada nota com líder, se a comissão do time entra nela:
//   { atribuir(time, nota) → { atribuir: boolean, arquivoAnterior?: string }, ordenar(notas) → notas }
// As notas são calculadas na ordem de ordenar(); atribuir só é chamado para notas em que o
// líder do time conta (obterLideresDaNota).
// ========================================

// Comissão em todas as notas em que o líder está (regra da versão original).
function politicaComissaoPorNota() {
    return { ordenar: (notas) => notas, atribuir: () => ({ atribuir: true }) };
}

// Regra decidida pelo usuário: a comissão do time no período vai só para a nota MAIS RECENTE
// em que o líder conta. A nota não precisa ter sido emitida dentro do período (as notas
// costumam sair depois da apuração e a base é limpa antes de cada pagamento).
// Ordem: data de emissão mais recente, depois dataExtracao mais recente, depois arquivo (A→Z).
// Nota sem data de emissão válida fica por último.
function politicaComissaoNotaMaisRecente() {
    const atribuidas = new Map(); // time normalizado → arquivo da nota que recebeu
    const chave = (n) => {
        const d = dataEmissaoParaDate(n.dataEmissao);
        return { emissao: d ? d.getTime() : -Infinity, extracao: n.dataExtracao || 0, arquivo: n.arquivo || '' };
    };
    return {
        ordenar: (notas) => notas
            .map(n => ({ n, k: chave(n) }))
            .sort((a, b) =>
                (b.k.emissao - a.k.emissao || 0) ||
                b.k.extracao - a.k.extracao ||
                a.k.arquivo.localeCompare(b.k.arquivo))
            .map(x => x.n),
        atribuir: (time, nota) => {
            if (atribuidas.has(time)) return { atribuir: false, arquivoAnterior: atribuidas.get(time) };
            atribuidas.set(time, nota.arquivo || '');
            return { atribuir: true };
        }
    };
}

// Política usada pela tela de conciliação. Trocar aqui muda tabela, detalhe e exportações.
function politicaComissaoPadrao() {
    return politicaComissaoNotaMaisRecente();
}

// ctx: { prestadores, recebivel, dataInicio, dataFim }
// politica (opcional): política de comissão; sem ela a nota é calculada isoladamente.
function calcularValorDevidoNota(nota, ctx, politica = politicaComissaoPorNota()) {
    const { prestadores, recebivel } = ctx;
    const { ini, fim } = normalizarPeriodo(ctx.dataInicio, ctx.dataFim);
    const noPeriodo = (r) => r.data >= ini && r.data <= fim;

    const motoristas = obterMotoristasEfetivos(nota, prestadores);
    const lideres = obterLideresDaNota(nota, motoristas, prestadores);
    const isLider = lideres.length > 0;

    // Candidatos a soma de Total Motorista
    const candidatos = [];
    const addCandidato = (vehicleId) => {
        if (!vehicleId) return;
        const p = buscarPrestadorPorVehicleId(prestadores, vehicleId);
        if (p && p.tipo === 'lider') return;
        const id = idMotorista(vehicleId);
        if (id && !candidatos.includes(id)) candidatos.push(id);
    };
    if (isLider && !nota.motoristasManual) {
        const liderCnpj = buscarPrestadorPorCnpj(prestadores, nota.cnpj);
        if (liderCnpj && liderCnpj.tipo === 'lider') {
            (liderCnpj.motoristasVinculados || []).forEach(m => addCandidato(m.vehicleId));
        }
    }
    motoristas.forEach(m => addCandidato(m.vehicleId));

    const totalPorMotorista = new Map(candidatos.map(id => [id, 0]));
    recebivel.forEach(r => {
        const id = idMotorista(r.driverId);
        if (totalPorMotorista.has(id) && noPeriodo(r)) {
            totalPorMotorista.set(id, totalPorMotorista.get(id) + (r.totalMotorista || 0));
        }
    });
    const driverIds = isLider
        ? candidatos
        : candidatos.filter(id => totalPorMotorista.get(id) > 0);
    const idsSomados = new Set(driverIds);
    const linhasMotoristas = recebivel.filter(r => idsSomados.has(idMotorista(r.driverId)) && noPeriodo(r));
    const valorMotoristas = linhasMotoristas.reduce((s, r) => s + (r.totalMotorista || 0), 0);

    // Comissão: uma vez por time na nota; entre notas, conforme a política
    const timesComissao = [];
    const comissaoEmOutraNota = [];
    lideres.forEach(l => {
        const time = normalizarTime(l.timeVinculado);
        if (!time || timesComissao.includes(time)) return;
        const decisao = politica.atribuir(time, nota);
        if (!decisao.atribuir) {
            comissaoEmOutraNota.push({ time: l.timeVinculado, arquivo: decisao.arquivoAnterior });
            return;
        }
        timesComissao.push(time);
    });
    const linhasComissao = recebivel.filter(r => timesComissao.includes(normalizarTime(r.time)) && noPeriodo(r));
    const comissao = linhasComissao.reduce((s, r) => s + (r.comissao || 0), 0);

    const valorDevido = valorMotoristas + comissao;
    return {
        motoristas,
        lideres,
        isLider,
        driverIds,
        linhasMotoristas,
        linhasComissao,
        timesComissao,
        comissaoEmOutraNota,
        valorMotoristas,
        comissao,
        valorDevido,
        temDados: isLider ? valorDevido > 0 : driverIds.length > 0
    };
}

// Calcula todas as notas da conciliação com a mesma política de comissão.
// Retorna Map nota → resultado de calcularValorDevidoNota.
function calcularConciliacaoNotas(notas, ctx, politica = politicaComissaoPadrao()) {
    const resultados = new Map();
    politica.ordenar(notas).forEach(nota => {
        resultados.set(nota, calcularValorDevidoNota(nota, ctx, politica));
    });
    return resultados;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        normalizeCnpj,
        idMotorista,
        buscarPrestadorPorCnpj,
        buscarPrestadorPorVehicleId,
        inferirMotoristasManual,
        migrarNotaLegada,
        listaExplicitaDoLider,
        reconciliarMotoristasDaNota,
        obterMotoristasEfetivos,
        obterLideresDaNota,
        politicaComissaoPorNota,
        politicaComissaoNotaMaisRecente,
        politicaComissaoPadrao,
        calcularValorDevidoNota,
        calcularConciliacaoNotas
    };
}
