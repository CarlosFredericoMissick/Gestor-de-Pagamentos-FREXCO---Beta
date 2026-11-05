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
// 3b. Motorista em 2 ou mais notas com recebível dividido (nota.alocacoesRecebivel): cada
//    nota soma só as rotas atribuídas a ela (ver DIVISÃO DO RECEBÍVEL abaixo).
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

// Composição da nota para a conciliação: motoristas efetivos, líderes que contam e os
// motoristas (não líderes) cujo Total Motorista pode entrar no valor devido.
function obterComposicaoNota(nota, prestadores) {
    const motoristas = obterMotoristasEfetivos(nota, prestadores);
    const lideres = obterLideresDaNota(nota, motoristas, prestadores);
    const isLider = lideres.length > 0;
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
    return { motoristas, lideres, isLider, candidatos };
}

// ctx: { prestadores, recebivel, dataInicio, dataFim }
// politica (opcional): política de comissão; sem ela a nota é calculada isoladamente.
// preparo (opcional): resultado de prepararDivisoes; motorista com recebível dividido entre
//   notas entra só com as rotas atribuídas a esta nota.
function calcularValorDevidoNota(nota, ctx, politica = politicaComissaoPorNota(), preparo = null) {
    const { prestadores, recebivel } = ctx;
    const { ini, fim } = normalizarPeriodo(ctx.dataInicio, ctx.dataFim);
    const noPeriodo = (r) => r.data >= ini && r.data <= fim;

    const { motoristas, lideres, isLider, candidatos } = obterComposicaoNota(nota, prestadores);

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

    // Motoristas com recebível dividido entre notas: só as rotas (ou partes) desta nota
    const divisao = [];
    const pecasPorIdx = new Map(); // posição no recebível → partes atribuídas a esta nota
    const divididos = new Set();
    driverIds.forEach(id => {
        const resultado = preparo?.divisoes.get(id);
        const daNota = resultado?.porNota.get(nota);
        if (!daNota) return;
        divididos.add(id);
        daNota.linhas.forEach(l => {
            if (!pecasPorIdx.has(l._idx)) pecasPorIdx.set(l._idx, []);
            pecasPorIdx.get(l._idx).push(l);
        });
        divisao.push({
            vehicleId: id,
            status: resultado.status,
            totalCentavos: resultado.totalCentavos,
            valorCentavos: daNota.centavos,
            alocadoCentavos: daNota.alocadoCentavos,
            recebeRestante: daNota.recebeRestante,
            qtdNotas: resultado.porNota.size
        });
    });
    const linhasMotoristas = [];
    recebivel.forEach((r, idx) => {
        const id = idMotorista(r.driverId);
        if (!idsSomados.has(id) || !noPeriodo(r)) return;
        if (divididos.has(id)) (pecasPorIdx.get(idx) || []).forEach(p => linhasMotoristas.push(p));
        else linhasMotoristas.push(r);
    });
    const valorMotoristas = linhasMotoristas.reduce((s, r) => s + centavos(r.totalMotorista), 0) / 100;

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

    // Motoristas desta nota que também estão em outras notas (habilita "Dividir recebível")
    const compartilhados = preparo
        ? candidatos.filter(id => (preparo.notasPorMotorista.get(id) || []).length >= 2)
        : [];
    const alocacaoOutroPeriodo = preparo
        ? candidatos.filter(id => temAlocacaoDeOutroPeriodo(nota, id, preparo.periodo))
        : [];

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
        divisao,
        compartilhados,
        alocacaoOutroPeriodo,
        temDados: isLider ? valorDevido > 0 : driverIds.length > 0
    };
}

// Calcula todas as notas da conciliação com a mesma política de comissão e a mesma divisão
// de recebível. Retorna Map nota → resultado de calcularValorDevidoNota.
function calcularConciliacaoNotas(notas, ctx, politica = politicaComissaoPadrao(), preparo = prepararDivisoes(notas, ctx)) {
    const resultados = new Map();
    politica.ordenar(notas).forEach(nota => {
        resultados.set(nota, calcularValorDevidoNota(nota, ctx, politica, preparo));
    });
    return resultados;
}

// ========================================
// DIVISÃO DO RECEBÍVEL DE UM MOTORISTA ENTRE NOTAS
//
// Quando o mesmo motorista está em 2 ou mais notas, o usuário pode fixar quanto do recebível
// dele no período vai para cada nota (nota.alocacoesRecebivel). Regra:
// - As rotas (linhas do recebível do motorista no período) são ordenadas por data e depois
//   pela posição na planilha importada.
// - As notas com alocação são preenchidas em ordem estável (emissão, extração, arquivo) até o
//   valor alocado; a rota que ultrapassa o alvo é fracionada (parte "parcial" na nota, o
//   resto vai para a nota que recebe o restante). Cada nota alocada fica com no máximo uma
//   rota fracionada, exceto se só sobrarem partes de rotas já fracionadas para completar o alvo.
// - Exatamente 1 nota sem alocação: recebe todas as rotas e partes que sobraram.
// - 2 ou mais notas sem alocação: elas continuam com o total do motorista (comportamento
//   anterior) e a linha avisa "aloque as demais notas deste motorista".
// - Soma alocada maior que o total, ou nenhuma nota sem alocação: divisão inválida, todas as
//   notas ficam com o total (o modal bloqueia esses casos ao salvar).
// Valores em centavos inteiros: a soma das partes é sempre o total do motorista.
// ========================================

function centavos(valor) { return Math.round((valor || 0) * 100); }

// Componentes exibidos no detalhe; são repartidos na mesma proporção da parte da rota.
const COMPONENTES_ROTA = ['valorRotas', 'ajuste', 'adicional', 'descarga', 'suporte'];

// Divide uma linha em [parte, sobra], com centavosParte na parte. A soma de cada componente
// entre parte e sobra é exatamente o valor original.
function fracionarLinha(linha, centavosParte) {
    const total = centavos(linha.totalMotorista);
    const totalOriginal = linha.totalOriginal ?? linha.totalMotorista;
    const parte = { ...linha, parcial: true, totalOriginal, totalMotorista: centavosParte / 100 };
    const sobra = { ...linha, parcial: true, totalOriginal, totalMotorista: (total - centavosParte) / 100 };
    COMPONENTES_ROTA.forEach(c => {
        const original = centavos(linha[c]);
        const p = total !== 0 ? Math.round(original * centavosParte / total) : 0;
        parte[c] = p / 100;
        sobra[c] = (original - p) / 100;
    });
    return [parte, sobra];
}

// Situação das alocações de um motorista (usada pelo modal e pela divisão).
// valoresCentavos: um item por nota; null = nota sem alocação.
function validarAlocacoes(totalCentavos, valoresCentavos) {
    const alocados = valoresCentavos.filter(v => v !== null && v !== undefined);
    const somaCentavos = alocados.reduce((s, v) => s + v, 0);
    const semAlocacao = valoresCentavos.length - alocados.length;
    let status = 'ok';
    if (alocados.some(v => v < 0)) status = 'negativo';
    else if (somaCentavos > totalCentavos) status = 'excede';
    else if (semAlocacao === 0) status = 'semRestante';
    else if (semAlocacao > 1) status = 'restanteMultiplo';
    return { status, somaCentavos, restanteCentavos: totalCentavos - somaCentavos, semAlocacao };
}

// linhas: linhas do recebível do motorista no período, cada uma com _idx (posição na planilha)
// alocacoes: Map nota → valor alocado em reais (só notas com alocação)
// notas: todas as notas que contêm o motorista, em ordem estável
// Retorna { status, totalCentavos, restanteCentavos,
//           porNota: Map nota → { linhas, centavos, alocadoCentavos|null, recebeRestante } }
function dividirRecebivelEntreNotas(linhas, alocacoes, notas) {
    const ordenadas = [...linhas].sort((a, b) => a.data - b.data || a._idx - b._idx);
    const totalCentavos = ordenadas.reduce((s, l) => s + centavos(l.totalMotorista), 0);
    const alvo = new Map([...alocacoes].filter(([n]) => notas.includes(n)).map(([n, v]) => [n, centavos(v)]));
    const { status, restanteCentavos } = validarAlocacoes(totalCentavos, notas.map(n => alvo.has(n) ? alvo.get(n) : null));
    const porNota = new Map(notas.map(n => [n, {
        linhas: [], centavos: 0, alocadoCentavos: alvo.has(n) ? alvo.get(n) : null, recebeRestante: false
    }]));
    const somaLinhas = (ls) => ls.reduce((s, l) => s + centavos(l.totalMotorista), 0);
    const totalEmCada = (lista) => lista.forEach(n => Object.assign(porNota.get(n), { linhas: ordenadas, centavos: totalCentavos }));

    if (status !== 'ok' && status !== 'restanteMultiplo') {
        totalEmCada(notas);
        return { status, totalCentavos, restanteCentavos, porNota };
    }

    const pool = [...ordenadas];
    const sobras = [];
    notas.filter(n => alvo.has(n)).forEach(n => {
        const destino = porNota.get(n);
        let falta = alvo.get(n);
        const consumir = (fila, guardarSobra) => {
            while (falta > 0 && fila.length > 0) {
                const linha = fila.shift();
                const valor = centavos(linha.totalMotorista);
                if (valor <= falta) {
                    destino.linhas.push(linha);
                    falta -= valor;
                } else {
                    const [parte, sobra] = fracionarLinha(linha, falta);
                    destino.linhas.push(parte);
                    guardarSobra(sobra);
                    falta = 0;
                }
            }
        };
        consumir(pool, s => sobras.push(s));
        consumir(sobras, s => sobras.unshift(s));
        destino.centavos = somaLinhas(destino.linhas);
    });

    const semAlocacao = notas.filter(n => !alvo.has(n));
    if (status === 'ok') {
        const resto = [...pool, ...sobras].sort((a, b) => a.data - b.data || a._idx - b._idx);
        Object.assign(porNota.get(semAlocacao[0]), { linhas: resto, centavos: somaLinhas(resto), recebeRestante: true });
    } else {
        totalEmCada(semAlocacao);
    }
    return { status, totalCentavos, restanteCentavos, porNota };
}

// ---------------- Alocações gravadas nas notas ----------------
// nota.alocacoesRecebivel = [{ vehicleId, valor, periodoInicio, periodoFim, editadoEm, editadoPor }]
// Uma alocação vale só para o período (datas da conciliação) em que foi feita.

function dataISO(d) {
    const p = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function periodoDaConciliacao(dataInicio, dataFim) {
    return { inicio: dataISO(dataInicio), fim: dataISO(dataFim) };
}

function mesmaAlocacao(a, vehicleId, periodo) {
    return idMotorista(a.vehicleId) === idMotorista(vehicleId)
        && a.periodoInicio === periodo.inicio && a.periodoFim === periodo.fim;
}

function alocacaoDaNota(nota, vehicleId, periodo) {
    return (nota.alocacoesRecebivel || []).find(a => mesmaAlocacao(a, vehicleId, periodo)) || null;
}

function temAlocacaoDeOutroPeriodo(nota, vehicleId, periodo) {
    return (nota.alocacoesRecebivel || []).some(a =>
        idMotorista(a.vehicleId) === idMotorista(vehicleId) && !mesmaAlocacao(a, vehicleId, periodo));
}

// Nova lista de alocações da nota trocando a do motorista no período (nova = null remove).
function substituirAlocacao(lista, vehicleId, periodo, nova) {
    const outras = (lista || []).filter(a => !mesmaAlocacao(a, vehicleId, periodo));
    return nova ? [...outras, nova] : outras;
}

// Remove alocações de motoristas que não estão mais na nota.
function descartarAlocacoesSemMotorista(lista, motoristas) {
    const ids = new Set((motoristas || []).map(m => idMotorista(m.vehicleId)));
    return (lista || []).filter(a => ids.has(idMotorista(a.vehicleId)));
}

// Ordem estável das notas de um motorista: emissão, extração, arquivo (mais antiga primeiro).
function ordenarNotasDoMotorista(notas) {
    const chave = (n) => {
        const d = dataEmissaoParaDate(n.dataEmissao);
        return { emissao: d ? d.getTime() : Infinity, extracao: n.dataExtracao || 0, arquivo: n.arquivo || '' };
    };
    return notas
        .map(n => ({ n, k: chave(n) }))
        .sort((a, b) => (a.k.emissao - b.k.emissao || 0) || a.k.extracao - b.k.extracao || a.k.arquivo.localeCompare(b.k.arquivo))
        .map(x => x.n);
}

// Prepara a divisão de todos os motoristas que aparecem em 2 ou mais notas (todas as notas,
// independente do filtro de time) e têm recebível positivo no período.
// Retorna { periodo, notasPorMotorista: Map id → notas, divisoes: Map id → dividirRecebivelEntreNotas }
function prepararDivisoes(notas, ctx) {
    const { ini, fim } = normalizarPeriodo(ctx.dataInicio, ctx.dataFim);
    const periodo = periodoDaConciliacao(ctx.dataInicio, ctx.dataFim);
    const linhasPorMotorista = new Map();
    ctx.recebivel.forEach((r, idx) => {
        if (!(r.data >= ini && r.data <= fim)) return;
        const id = idMotorista(r.driverId);
        if (!linhasPorMotorista.has(id)) linhasPorMotorista.set(id, []);
        linhasPorMotorista.get(id).push({ ...r, _idx: idx });
    });

    const notasPorMotorista = new Map();
    notas.forEach(nota => {
        obterComposicaoNota(nota, ctx.prestadores).candidatos.forEach(id => {
            const linhas = linhasPorMotorista.get(id) || [];
            if (linhas.reduce((s, l) => s + centavos(l.totalMotorista), 0) <= 0) return;
            if (!notasPorMotorista.has(id)) notasPorMotorista.set(id, []);
            notasPorMotorista.get(id).push(nota);
        });
    });

    const divisoes = new Map();
    notasPorMotorista.forEach((lista, id) => {
        if (lista.length < 2) return;
        const ordenadas = ordenarNotasDoMotorista(lista);
        notasPorMotorista.set(id, ordenadas);
        const alocacoes = new Map();
        ordenadas.forEach(n => {
            const a = alocacaoDaNota(n, id, periodo);
            if (a) alocacoes.set(n, a.valor);
        });
        divisoes.set(id, dividirRecebivelEntreNotas(linhasPorMotorista.get(id), alocacoes, ordenadas));
    });
    return { periodo, notasPorMotorista, linhasPorMotorista, divisoes };
}

// ========================================
// RÓTULOS DE EXIBIÇÃO (não gravados no JSON)
// Líder aparece como "<nome> Comissão". Motorista com o mesmo nome de um líder presente na
// mesma lista aparece como "<nome> Rotas"; sem ambiguidade, só o nome.
// ========================================

function normalizarNome(nome) {
    return (nome || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function ehLider(m, prestadores) {
    const p = buscarPrestadorPorVehicleId(prestadores, m.vehicleId);
    return p ? p.tipo === 'lider' : /^L\d+$/.test(String(m.vehicleId ?? ''));
}

function nomesDeLideres(lista, prestadores) {
    return new Set((lista || []).filter(m => ehLider(m, prestadores)).map(m => normalizarNome(m.nome)));
}

function rotuloMotorista(m, prestadores, nomesLideres = new Set()) {
    const nome = m.nome || '';
    if (ehLider(m, prestadores)) return `${nome} Comissão`;
    if (nomesLideres.has(normalizarNome(nome))) return `${nome} Rotas`;
    return nome;
}

function rotulosMotoristas(lista, prestadores) {
    const nomesLideres = nomesDeLideres(lista, prestadores);
    return (lista || []).map(m => rotuloMotorista(m, prestadores, nomesLideres));
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
        obterComposicaoNota,
        calcularValorDevidoNota,
        calcularConciliacaoNotas,
        centavos,
        fracionarLinha,
        validarAlocacoes,
        dividirRecebivelEntreNotas,
        periodoDaConciliacao,
        alocacaoDaNota,
        temAlocacaoDeOutroPeriodo,
        substituirAlocacao,
        descartarAlocacoesSemMotorista,
        ordenarNotasDoMotorista,
        prepararDivisoes,
        rotuloMotorista,
        rotulosMotoristas,
        nomesDeLideres
    };
}
