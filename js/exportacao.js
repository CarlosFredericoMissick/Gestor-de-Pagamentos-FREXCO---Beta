// ========================================
// EXPORTAÇÕES — Excel de notas, ZIP de notas, prestadores, contratos, conciliação e detalhe
// ========================================

function exportToExcel() {
    const filtradas = getNotasFiltradas();
    if (filtradas.length === 0) {
        alert("Nenhum dado para exportar!");
        return;
    }
    
    try {
        const data = filtradas.map(row => {
            let dateObj = null;
            if (row.dataEmissao && row.dataEmissao.includes('/')) {
                let datePart = row.dataEmissao.match(/(\d{2}\/\d{2}\/\d{4})/);
                if (datePart) {
                    const parts = datePart[0].split('/');
                    dateObj = new Date(parts[2], parts[1] - 1, parts[0], 12, 0, 0);
                }
            }
            
            let nfNum = row.numeroNF;
            const nfStr = String(row.numeroNF);
            if (!isNaN(parseInt(nfStr)) && !nfStr.includes("/")) {
                nfNum = parseInt(nfStr);
            }

            return {
                'Responsável': row.responsavel?.nome || row.prestador || '', // A ← alterado
                'Razão Social': row.prestador,       // B
                'CNPJ Prestador': row.cnpj,          // C
                'Valor Líquido': row.vLiquido,       // D
                'Valor Serviço': row.vServico,       // E
                ' ': '',                             // F (vazia)
                'Chave PIX': row.chavePix || '',     // G
                '  ': '',                            // H (vazia)
                '   ': '',                           // I (vazia)
                'Data Emissão': dateObj,             // J
                'Número NF': nfNum,                  // K
                'Tem Contrato': row.temContrato ? 'Sim' : 'Não', // L
                'Nome Contrato': row.nomeContrato || '-',         // M
                'Modelo': row.modelo,                             // N
                'Método': row.method || 'Geométrico'              // O
            };
        });
        
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "NFSe Extraídas");
        
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'Relatorio_NFSe_v32.xlsx';
        document.body.appendChild(a);
        a.click();
        setTimeout(() => {
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
        }, 100);
        
        console.log('Excel exportado com sucesso!');
    } catch (error) {
        console.error("Erro ao exportar:", error);
        alert("Erro ao exportar. Verifique o console.");
    }
}

async function gerarZipNotas() {
    const notas = getNotasFiltradas().filter(n => n._pdfSpPath || n._pdfSpId);
    if (notas.length === 0) {
        alert('Nenhuma nota com PDF armazenado no SharePoint. Processe as notas enquanto conectado ao SharePoint para habilitar esta função.');
        return;
    }
    if (!spAccessToken) { alert('Conecte o SharePoint para gerar o ZIP.'); return; }

    const btn  = document.getElementById('btnGerarZip');
    const span = btn?.querySelector('span');
    if (btn) btn.disabled = true;
    if (span) span.textContent = 'Gerando ZIP...';

    try {
        const zip = new JSZip();
        let ok = 0, erros = 0;

        for (let i = 0; i < notas.length; i++) {
            const nota = notas[i];
            if (span) span.textContent = `Baixando ${i+1}/${notas.length}...`;
            try {
                // Tenta pelo caminho direto
                const url = urlApproot(nota._pdfSpPath, ':/content');
                const resp = await graphFetch(url);
                if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
                const buf = await resp.arrayBuffer();

                // Nome do arquivo: baseado nos Motoristas na Nota
                const motoristas = (nota.motoristasNaNota && nota.motoristasNaNota.length > 0)
                    ? nota.motoristasNaNota
                    : (nota.responsavel ? [nota.responsavel] : []);
                const nomesMotoristas = motoristas
                    .map(m => (m.nome || '').replace(/[/\\:*?"<>|]/g, '_').trim())
                    .filter(Boolean)
                    .join(', ');
                const nomeBase = nomesMotoristas || nota.prestador || nota.arquivo || 'sem_motorista';
                const nf = nota.numeroNF && nota.numeroNF !== '?' ? `_NF${nota.numeroNF}` : '';
                const zipName = `${nomeBase}${nf}.pdf`;
                zip.file(zipName, buf);
                ok++;
            } catch(e) {
                console.error(`Erro ao baixar PDF ${nota.arquivo}:`, e);
                erros++;
            }
        }

        if (ok === 0) {
            alert('Nenhum PDF foi baixado. Verifique a conexão com o SharePoint.');
            return;
        }

        if (span) span.textContent = 'Compactando...';
        const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = `notas_renomeadas_${new Date().toISOString().slice(0,10)}.zip`;
        a.click();
        URL.revokeObjectURL(a.href);

        let msg = `✅ ZIP gerado com ${ok} nota(s).`;
        if (erros > 0) msg += `\n⚠️ ${erros} nota(s) não encontrada(s) no SharePoint.`;
        alert(msg);
    } catch(e) {
        console.error('Erro ao gerar ZIP:', e);
        alert('Erro ao gerar ZIP: ' + e.message);
    } finally {
        if (btn) btn.disabled = false;
        if (span) span.textContent = 'Baixar ZIP';
    }
}

function exportPrestadoresExcel() {
    const rows = prestadoresList.filter(p => p.inUse !== false).flatMap(p =>
        (p.cnpjs && p.cnpjs.length > 0)
            ? p.cnpjs.map(c => [p.vehicleId||'', p.nome, p.time||'', p.cd||'', c.cnpj, c.chavePix, p.contrato ? p.contrato.nomeArquivo : ''])
            : [[p.vehicleId||'', p.nome, p.time||'', p.cd||'', '', '', p.contrato ? p.contrato.nomeArquivo : '']]
    );
    const ws = XLSX.utils.aoa_to_sheet([['ID','Nome','Time','CD','CNPJ','Chave PIX','Contrato'], ...rows]);
    ws['!cols'] = [{wch:8},{wch:30},{wch:18},{wch:25},{wch:22},{wch:30},{wch:35}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Prestadores');
    XLSX.writeFile(wb, 'prestadores_frexco.xlsx');
}

async function downloadAllContratosZip() {
    const comContrato = prestadoresList.filter(p => p.contrato);
    if (comContrato.length === 0) { alert('Nenhum contrato para baixar.'); return; }
    const zip = new JSZip();
    for (const p of comContrato) {
        if (!p.contrato) continue;

        // Nome padronizado para reimportação: contrato_{driverId}_{nome}.pdf
        const safeName = (p.nome || 'sem_nome').replace(/[/\\:*?"<>|]/g, '_').trim();
        const zipFileName = `contrato_${p.vehicleId}_${safeName}.pdf`;

        if (useSharePoint) {
            // Tenta buscar pelo spFileName (caminho) primeiro, depois pelo id
            const fetchUrl = p.contrato.spFileName
                ? urlApproot(p.contrato.spFileName, ':/content')
                : p.contrato.id
                    ? `${GRAPH_ITEMS}/${p.contrato.id}/content`
                    : null;

            if (!fetchUrl) continue;
            try {
                const resp = await graphFetch(fetchUrl);
                if (resp.ok) {
                    const buf = await resp.arrayBuffer();
                    zip.file(zipFileName, buf);
                }
            } catch(e) { console.error('Zip SP contrato:', p.nome, e); }
        } else if (p.contrato.fileData) {
            const b64 = p.contrato.fileData.split(',')[1];
            zip.file(zipFileName, b64, { base64: true });
        }
    }
    const blob = await zip.generateAsync({ type: 'blob' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = 'contratos_frexco.zip'; a.click();
    URL.revokeObjectURL(url);
}

function exportarConciliacaoExcel() {
    if (_concilRows.length === 0) { alert('Rode a conciliação primeiro.'); return; }
    const wb = XLSX.utils.book_new();
    const data = [['Arquivo','Responsável','Motoristas','Data Emissão','Valor Nota','Valor Devido','Diferença','Dif %','Status']];
    _concilRows.forEach(r => {
        const cells = r.tr.querySelectorAll('td');
        data.push([...cells].map(td => td.textContent.trim()));
    });
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = data[0].map(() => ({ wch: 18 }));
    XLSX.utils.book_append_sheet(wb, ws, 'Conciliação');
    XLSX.writeFile(wb, `conciliacao_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// ========================================
// EXPORTAR EXCEL — DETALHE MOTORISTA
// ========================================

// Usa as mesmas linhas do recebível que o detalhe na tela (resultado de calcularValorDevidoNota).
function exportarDetalheMotoristaExcel() {
    if (!_concilDetalheAtual) return;
    const { nome, isLider, resultado } = _concilDetalheAtual;
    const fmtDate = d => (d instanceof Date ? d : new Date(d)).toLocaleDateString('pt-BR');
    const wb = XLSX.utils.book_new();
    const linhaMotorista = r => [
        fmtDate(r.data), r.nome||'', r.rota||'',
        r.valorRotas, r.ajuste, r.adicional, r.descarga||0, r.suporte, r.totalMotorista
    ];

    if (isLider) {
        // Aba 1: Motoristas na nota
        const dataM = [['Data','Veículo','Rota','Valor Rotas','Ajuste','Adicional','Descarga','Suporte','Total']];
        resultado.linhasMotoristas.forEach(r => dataM.push(linhaMotorista(r)));
        const wsM = XLSX.utils.aoa_to_sheet(dataM);
        wsM['!cols'] = [{wch:12},{wch:20},{wch:10},{wch:14},{wch:12},{wch:14},{wch:12},{wch:14}];
        XLSX.utils.book_append_sheet(wb, wsM, 'Motoristas na Nota');

        // Aba 2: Comissão do time
        const dataC = [['Data','Motorista','Comissão']];
        [...resultado.linhasComissao].sort((a,b) => a.data - b.data)
            .forEach(r => dataC.push([fmtDate(r.data), r.nome||'', r.comissao||0]));
        resultado.comissaoEmOutraNota.forEach(c =>
            dataC.push([`Comissão do time ${c.time} no período entra só na nota mais recente do líder: ${c.arquivo}`, '', 0]));
        const wsC = XLSX.utils.aoa_to_sheet(dataC);
        wsC['!cols'] = [{wch:12},{wch:24},{wch:14}];
        const times = resultado.lideres.map(l => l.timeVinculado).join(', ');
        XLSX.utils.book_append_sheet(wb, wsC, `Comissão ${times}`.slice(0,31));

    } else {
        // Motorista normal — aba única
        const data = [['Data','Veículo','Rota','Valor Rotas','Ajuste','Adicional','Descarga','Suporte','Total Motorista']];
        resultado.linhasMotoristas.forEach(r => data.push(linhaMotorista(r)));
        const ws = XLSX.utils.aoa_to_sheet(data);
        ws['!cols'] = [{wch:12},{wch:20},{wch:10},{wch:14},{wch:12},{wch:14},{wch:12},{wch:16}];
        XLSX.utils.book_append_sheet(wb, ws, (nome||'Detalhe').slice(0,31));
    }

    XLSX.writeFile(wb, `detalhe_${(nome||'motorista').replace(/\s/g,'_')}_${new Date().toISOString().slice(0,10)}.xlsx`);
}


function baixarExemploCnpjPix() {
    const wb = XLSX.utils.book_new();
    const data = [
        ['ID Motorista', 'Nome', 'CNPJ', 'Chave PIX'],
        [42, 'Ademário', '12.345.678/0001-90', '12.345.678/0001-90'],
        [228, 'Adilson Costa', '37.148.610/0001-47', '1998706-0403'],
        [672, 'Adriano Souza', '49.457.157/0001-30', 'adriano@email.com'],
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{wch:16},{wch:24},{wch:22},{wch:24}];
    // Bold header style hint via cell format
    XLSX.utils.book_append_sheet(wb, ws, 'CNPJs_PIX');
    XLSX.writeFile(wb, 'modelo_cnpjs_pix.xlsx');
}
