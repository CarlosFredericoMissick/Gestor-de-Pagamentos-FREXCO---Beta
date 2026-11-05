// ========================================
// EXTRAÇÃO DE NOTAS — fila de PDFs, leitura geométrica/OCR e parsers por prefeitura
// ========================================

function handleFiles(files) {
    filesQueue = [...filesQueue, ...Array.from(files).filter(f => f.type === 'application/pdf')];
    document.getElementById('fileListContainer').innerText = `${filesQueue.length} arquivos na fila`;
    document.getElementById('btnProcess').disabled = false;
    document.getElementById('btnProcess').classList.replace('bg-gray-300', 'bg-primary');
    document.getElementById('btnProcess').classList.replace('cursor-not-allowed', 'hover:bg-primary-dark');
}

async function startProcessing() {
    if (isProcessing) return;
    isProcessing = true;
    document.getElementById('statusPanel').classList.remove('hidden');

    const total = filesQueue.length;
    const worker = await Tesseract.createWorker('por');
    // Gravações no SharePoint seguem em paralelo à leitura; o resultado é avisado no fim do lote
    const gravacoes = [];

    for (let i = 0; i < total; i++) {
        const file = filesQueue[i];
        updateStatus(`Lendo ${i+1}/${total}: ${file.name}`, (i/total)*100);

        try {
            let lines = await extractLinesFromPDF(file);
            let method = "Geométrico";
            
            if (lines.length < 5) { 
                 updateStatus(`OCR em ${file.name}...`, (i/total)*100);
                 const ocrText = await extractTextWithOCR(file, worker);
                 lines = ocrText.split('\n').map(l => l.trim()).filter(l => l.length > 0);
                 method = "OCR";
            }

            let data = {};
            const fullTextUpper = lines.join(" ").toUpperCase();

            if (fullTextUpper.includes("DANFSE") || fullTextUpper.includes("DOCUMENTO AUXILIAR DA NFS-E")) {
                data = extractDanfseNacional(lines, file.name);
            }
            else if (fullTextUpper.includes("PREFEITURA MUNICIPAL CAMPINAS")) {
                data = extractCampinas(lines, file.name);
            }
            else if (fullTextUpper.includes("PREFEITURA MUNICIPAL DE PIEDADE")) {
                data = extractPiedade(lines, file.name);
            }
            else if (fullTextUpper.includes("PREFEITURA DO MUNICÍPIO DE SÃO PAULO") || fullTextUpper.includes("PREFEITURA DO MUNICIPIO DE SAO PAULO")) {
                data = extractSaoPaulo(lines, file.name);
            }
            else {
                data = extractStandardV7(lines, file.name);
            }

            data.method = method;
            data.dataExtracao = Date.now();
            data._id = `nota_${Date.now()}_${Math.random().toString(36).substr(2,6)}`;

            // Responsável — prestador resolvido pelo CNPJ da nota
            const prestadorResp = findPrestadorByCnpj(data.cnpj);
            data.responsavel = prestadorResp
                ? { vehicleId: prestadorResp.vehicleId, nome: prestadorResp.nome }
                : null;

            // Motoristas na nota — padrão = responsável
            data.motoristasNaNota = data.responsavel ? [{ ...data.responsavel }] : [];
            data.motoristasManual = false;

            // Verifica contrato (CNPJ-first)
            // motoristasNaNota set just above — use full check
            const contratoInfo0 = verificarContratoNota(data);
            data.temContrato       = contratoInfo0.status !== 'nao';
            data.temContratoStatus = contratoInfo0.status;
            data.nomeContrato      = contratoInfo0.contrato;

            // Busca chave PIX (CNPJ-first)
            data.chavePix = findPixByFileName(file.name, data.cnpj);
            
            extractedResults.push(data);
            renderRow(data);
            atualizarOpcoesTime();
            if (useSharePoint) {
                const gravacaoNota = saveNotaSP(data);
                // Upload do PDF original para o SharePoint (para geração de ZIP posterior)
                const gravacaoPdf = uploadPdfSP(file, data, gravacaoNota);
                gravacoes.push(Promise.all([gravacaoNota, gravacaoPdf])
                    // Estado final: a regravação com a referência do PDF grava a nota inteira e
                    // supre uma falha da gravação inicial.
                    .then(([nota, pdf]) => ({ arquivo: data.arquivo, notaOk: nota.ok || pdf, pdfOk: pdf })));
            }

        } catch (e) {
            console.error(e);
            const errData = { arquivo: file.name, prestador: "ERRO", status: "Erro", method: "Erro", temContrato: false, numeroNF: "?" };
            extractedResults.push(errData);
            renderRow(errData);
        }
    }
    
    await worker.terminate();

    if (gravacoes.length > 0) {
        updateStatus("Salvando no SharePoint...", 100);
        const resultados = await Promise.all(gravacoes);
        const semNota = resultados.filter(r => !r.notaOk);
        const semPdf  = resultados.filter(r => r.notaOk && !r.pdfOk);
        if (semNota.length || semPdf.length) {
            const lista = (rs) => rs.slice(0, 10).map(r => r.arquivo).join('\n');
            let msg = '';
            if (semNota.length) msg += `❌ ${semNota.length} nota(s) não foram salvas no SharePoint (os outros usuários não as verão):\n${lista(semNota)}\n\n`;
            if (semPdf.length)  msg += `⚠️ ${semPdf.length} PDF(s) não foram enviados (ficam fora do "Baixar ZIP"):\n${lista(semPdf)}\n\n`;
            msg += spExigeReconexao ? 'A sessão do SharePoint expirou: reconecte e processe esses arquivos de novo.' : 'Processe esses arquivos de novo.';
            alert(msg);
        }
    }
    updateStatus("Concluído", 100);
    updateCoverageStats();
    document.getElementById('btnExport').classList.remove('hidden');
    document.getElementById('btnExport').disabled = false;
    
    // Limpar fila após conclusão
    filesQueue = [];
    document.getElementById('fileListContainer').innerText = '0 arquivos na fila';
    
    isProcessing = false;
}

function updateStatus(msg, pct) {
    document.getElementById('statusText').innerText = msg;
    document.getElementById('progressBar').style.width = `${pct}%`;
}

async function extractLinesFromPDF(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const page = await pdf.getPage(1); 
    const textContent = await page.getTextContent();
    const items = textContent.items.map(item => ({ text: item.str, y: item.transform[5], x: item.transform[4] }));
    items.sort((a, b) => b.y - a.y || a.x - b.x);
    const lines = [];
    let currentLine = [];
    let lastY = items[0]?.y || 0;
    for (const item of items) {
        if (Math.abs(item.y - lastY) > 5) {
            if (currentLine.length > 0) lines.push(currentLine.join(" ").trim());
            currentLine = [];
            lastY = item.y;
        }
        currentLine.push(item.text);
    }
    if (currentLine.length > 0) lines.push(currentLine.join(" ").trim());
    return lines;
}

async function extractTextWithOCR(file, worker) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const page = await pdf.getPage(1);
    const viewport = page.getViewport({ scale: 3.0 });
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    canvas.height = viewport.height;
    canvas.width = viewport.width;
    
    await page.render({ canvasContext: context, viewport: viewport }).promise;
    
    const imageData = context.getImageData(0, 0, canvas.width, canvas.height);
    const data = imageData.data;
    
    for (let i = 0; i < data.length; i += 4) {
        const avg = (data[i] + data[i + 1] + data[i + 2]) / 3;
        const newVal = avg > 128 ? 255 : 0;
        data[i] = newVal;
        data[i + 1] = newVal;
        data[i + 2] = newVal;
    }
    context.putImageData(imageData, 0, 0);
    
    const { data: { text } } = await worker.recognize(canvas);
    return text;
}

function extractCampinas(lines, filename) {
    let data = { arquivo: filename, modelo: "Campinas", status: "Sucesso", prestador: "Não identificado", cnpj: "", vLiquido: 0, vServico: 0, dataEmissao: "", numeroNF: "?" };
    const fullText = lines.join("\n");

    for(let i=0; i<lines.length; i++) {
        if(lines[i].includes("Número / Série") || lines[i].includes("Número/Série")) {
            for(let k=1; k<=3; k++) {
                if(i+k < lines.length) {
                    let match = lines[i+k].match(/(\d+)\s*\/\s*[A-Z]/);
                    if(match) {
                        data.numeroNF = match[1];
                        break;
                    }
                }
            }
            if(data.numeroNF !== "?") break;
        }
        
        if(lines[i].includes("Competência") && data.numeroNF === "?") {
            let match = lines[i].match(/(\d+)\s*\/\s*[A-Z]/);
            if(match) {
                data.numeroNF = match[1];
                break;
            }
        }
    }

    for(let i=0; i<lines.length; i++) {
        if(lines[i].includes("Nome / Nome Empresarial") || lines[i].includes("Nome/Nome Empresarial")) {
            if(i+1 < lines.length && !lines[i+1].toUpperCase().includes("FREXCO")) {
                data.prestador = lines[i+1].replace(/[\w._%+-]+@[\w.-]+\.[a-zA-Z]{2,}/g, '').trim();
                break;
            }
        }
    }

     for(let i=0; i<lines.length; i++) {
        if(lines[i].includes("CPF / CNPJ / NIF") || lines[i].includes("CPF/CNPJ/NIF")) {
            if(i+1 < lines.length) {
                 const match = lines[i+1].match(/\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}/);
                 if(match) { data.cnpj = match[0]; break; }
            }
        }
    }

    const liqMatch = fullText.match(/Valor\s*L[íi]quido\s*da\s*NFSe\s*Campinas\s*\(R\$\)[\s\S]*?([\d.,]+)/i);
    if(liqMatch) {
        let cleanVal = liqMatch[1].trim();
        if (cleanVal) data.vLiquido = parseFloat(cleanVal.replace(/\./g,'').replace(',','.'));
    }
    
    const totalMatch = fullText.match(/Valor\s*total\s*da\s*NFSe\s*Campinas\s*\(R\$\)[\s\S]*?([\d.,]+)/i);
    if(totalMatch) data.vServico = parseFloat(totalMatch[1].replace(/\./g,'').replace(',','.'));
    
    if(!data.vServico && data.vLiquido) data.vServico = data.vLiquido;

    for(let i=0; i<lines.length; i++) {
        if(lines[i].toLowerCase().includes("data e hora de emissão")) {
            if(i+1 < lines.length) {
                const dateMatch = lines[i+1].match(/(\d{2}\/\d{2}\/\d{4})/);
                if(dateMatch) data.dataEmissao = dateMatch[1];
            }
        }
    }

    return data;
}

function extractPiedade(lines, filename) {
        let data = { arquivo: filename, modelo: "Piedade", status: "Sucesso", prestador: "Não identificado", cnpj: "", vLiquido: 0, vServico: 0, dataEmissao: "", numeroNF: "?" };
        const fullText = lines.join("\n");

        function extractMoney(str) {
                const m = str.match(/([\d]{1,3}(?:\.[\d]{3})*,\d{2})/);
                return m ? parseFloat(m[1].replace(/\./g,'').replace(',','.')) : null;
        }

        // --- Número da NFS-e ---
        // Layout 2026 (giss): o número aparece como linha isolada logo antes de "Código de Verificação"
        for (let i = 0; i < lines.length; i++) {
                if (/C[óo]digo\s+de\s+Verifica[çc][ãa]o/i.test(lines[i])) {
                        for (let k = i - 1; k >= 0 && k >= i - 4; k--) {
                                const cand = lines[k].trim();
                                if (/^\d{1,7}$/.test(cand)) { data.numeroNF = cand; break; }
                                const m = cand.match(/\bNFS-?e\b[^\d]*(\d{1,7})\b/i);
                                if (m) { data.numeroNF = m[1]; break; }
                        }
                        if (data.numeroNF !== "?") break;
                }
        }
        // Fallback layout antigo: "Número da ... NFS-e ... dígitos"
        if (data.numeroNF === "?") {
                const numMatch = fullText.match(/Número\s*da[\s\S]*?NFS-?e[\s\S]*?(\d{1,6})/i);
                if (numMatch) data.numeroNF = numMatch[1];
        }

        // --- Prestador (Razão Social) ---
        // Layout 2026: "Nome/Razão Social: SP TRANSPORTES LTDA" (pega o 1º que não é FREXCO/tomador)
        for (let i = 0; i < lines.length; i++) {
                if (/Nome\s*\/\s*Raz[ãa]o\s*Social/i.test(lines[i])) {
                        let name = lines[i].split(/:/).slice(1).join(":").trim();
                        if (!name && i + 1 < lines.length) name = lines[i+1].trim();
                        name = name.replace(/[\w._%+-]+@[\w.-]+\.[a-zA-Z]{2,}/g, '').trim();
                        if (name.length > 3 && !name.toUpperCase().includes("FREXCO")) { data.prestador = name; break; }
                }
        }
        // Fallback layout antigo: "Razão Social/Nome"
        if (data.prestador === "Não identificado") {
                for (let i = 0; i < lines.length; i++) {
                        if (lines[i].includes("Razão Social/Nome")) {
                                let name = lines[i].replace("Razão Social/Nome", "").trim();
                                if (name.length > 3) data.prestador = name;
                                else if (lines[i+1]) data.prestador = lines[i+1].trim();
                                break;
                        }
                }
        }

        // --- CNPJ do prestador (ignora o CNPJ da Frexco/tomador) ---
        const cnpjs = fullText.match(/\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}/g) || [];
        const prestCnpj = cnpjs.find(c => !c.includes("33.935.782"));
        data.cnpj = prestCnpj || cnpjs[0] || "";

        // --- Valor dos Serviços ---
        for (let i = 0; i < lines.length; i++) {
                if (/ervi[çc]/i.test(lines[i]) && /Valor/i.test(lines[i])) {
                        const val = extractMoney(lines[i]);
                        if (val) { data.vServico = val; break; }
                        if (i+1 < lines.length) { const v2 = extractMoney(lines[i+1]); if (v2) { data.vServico = v2; break; } }
                }
        }
        if (!data.vServico) {
                const sm = fullText.match(/Valor\s*d[eo]s?\s*Servi[\s\S]{0,20}?([\d]{1,3}(?:\.[\d]{3})*,\d{2})/i);
                if (sm) data.vServico = parseFloat(sm[1].replace(/\./g,'').replace(',','.'));
        }

        // --- Valor Líquido ---
        for (let i = 0; i < lines.length; i++) {
                if (/quido/i.test(lines[i])) {
                        const val = extractMoney(lines[i]);
                        if (val) { data.vLiquido = val; break; }
                        if (i+1 < lines.length) { const v2 = extractMoney(lines[i+1]); if (v2) { data.vLiquido = v2; break; } }
                }
        }
        if (!data.vLiquido) {
                const lm = fullText.match(/quido[\s\S]{0,40}?([\d]{1,3}(?:\.[\d]{3})*,\d{2})/i);
                if (lm) data.vLiquido = parseFloat(lm[1].replace(/\./g,'').replace(',','.'));
        }
        if (!data.vLiquido && data.vServico) data.vLiquido = data.vServico;

        // --- Data de emissão ---
        const dataMatch = fullText.match(/Emiss[ãa]o\s*da\s*NFS-?e[\s\S]*?(\d{2}\/\d{2}\/\d{4})/i);
        if (dataMatch) data.dataEmissao = dataMatch[1];
        if (!data.dataEmissao) {
                const dm2 = fullText.match(/Data\s*e\s*Hora\s*d[ae]\s*Emiss[ãa]o[\s\S]*?(\d{2}\/\d{2}\/\d{4})/i);
                if (dm2) data.dataEmissao = dm2[1];
        }
        if (!data.dataEmissao) {
                for (let i=0; i<lines.length; i++) {
                        if (/Emiss[ãa]o/i.test(lines[i])) {
                                let dm = lines[i].match(/(\d{2}\/\d{2}\/\d{4})/) || (lines[i+1]||"").match(/(\d{2}\/\d{2}\/\d{4})/);
                                if (dm) { data.dataEmissao = dm[1]; break; }
                        }
                }
        }

        return data;
}

function extractDanfseNacional(lines, filename) {
        let data = { arquivo: filename, modelo: "Nacional (DANFSe)", status: "Sucesso", prestador: "Não identificado", cnpj: "", vLiquido: 0, vServico: 0, dataEmissao: "", numeroNF: "?" };
        const fullText = lines.join("\n");
        const money = (s) => { const m = (s||"").match(/([\d]{1,3}(?:\.[\d]{3})*,\d{2})/); return m ? parseFloat(m[1].replace(/\./g,'').replace(',','.')) : null; };

        // --- Número da NFS-e (não confundir com Número da DPS) ---
        for (let i = 0; i < lines.length; i++) {
                const up = lines[i].toUpperCase();
                if (up.includes("NÚMERO DA NFS-E") && !up.includes("DPS")) {
                        for (let k = 1; k <= 2 && i + k < lines.length; k++) {
                                const m = lines[i+k].match(/^\s*(\d{1,9})\b/);
                                if (m) { data.numeroNF = m[1]; break; }
                        }
                        if (data.numeroNF !== "?") break;
                }
        }
        // Fallback: rodapé "38 / <chave de acesso 20+ dígitos>"
        if (data.numeroNF === "?") {
                const m = fullText.match(/(\d{1,9})\s*\/\s*\d{20,}/);
                if (m) data.numeroNF = m[1];
        }

        // --- CNPJ do prestador (primeiro CNPJ que não é da Frexco/tomador) ---
        const cnpjs = fullText.match(/\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}/g) || [];
        const prestCnpj = cnpjs.find(c => !c.includes("33.935.782"));
        data.cnpj = prestCnpj || cnpjs[0] || "";

        // --- Nome / Razão Social do prestador ---
        // No bloco PRESTADOR/FORNECEDOR, a linha após "Nome / Nome Empresarial"
        // mescla "Nome  Município/UF  IBGE/CEP". Corta no município (Title Case).
        let prestadorIdx = lines.findIndex(l => /PRESTADOR\s*\/\s*FORNECEDOR/i.test(l));
        if (prestadorIdx === -1) prestadorIdx = 0;
        for (let i = prestadorIdx; i < lines.length; i++) {
                if (/Nome\s*\/\s*Nome\s*Empresarial/i.test(lines[i]) && i + 1 < lines.length) {
                        let raw = lines[i+1].trim();
                        let name = raw.replace(/\s+[A-ZÀ-Ú][a-zà-ú][^\/]*\/\s*[A-Z]{2}(\s+\d[\d.\/\s-]*)?\s*$/, '').trim();
                        name = name.replace(/\s+\d{2}\.\d{4,7}\s*\/\s*[\d.-]{4,}\s*$/, '').trim();
                        if (name.length > 2) data.prestador = name;
                        break;
                }
        }

        // --- Valor do serviço: "VALOR DA OPERAÇÃO / SERVIÇO" ---
        for (let i = 0; i < lines.length; i++) {
                if (/VALOR\s+DA\s+OPERA[ÇC][ÃA]O\s*\/\s*SERVI[ÇC]O/i.test(lines[i])) {
                        for (let k = 1; k <= 2 && i + k < lines.length; k++) {
                                const v = money(lines[i+k]); if (v !== null) { data.vServico = v; break; }
                        }
                        if (data.vServico) break;
                }
        }

        // --- Valor líquido: "VALOR LÍQUIDO DA NFS-e" (não o "+ IBS/CBS") ---
        for (let i = 0; i < lines.length; i++) {
                const before = (lines[i].split(/VALOR\s+L[ÍI]QUIDO/i)[0]) || "";
                if (/VALOR\s+L[ÍI]QUIDO\s+DA\s+NFS-?E/i.test(lines[i]) && !/\+\s*IBS/i.test(before)) {
                        for (let k = 1; k <= 2 && i + k < lines.length; k++) {
                                const v = money(lines[i+k]); // primeiro valor monetário da linha de valores (pula "-")
                                if (v !== null) { data.vLiquido = v; break; }
                        }
                        if (data.vLiquido) break;
                }
        }
        if (!data.vLiquido && data.vServico) data.vLiquido = data.vServico;
        if (!data.vServico && data.vLiquido) data.vServico = data.vLiquido;

        // --- Data e hora de emissão da NFS-e ---
        let de = fullText.match(/DATA\s*E\s*HORA\s*D[AE]\s*EMISS[ÃA]O\s*DA\s*NFS-?E[\s\S]*?(\d{2}\/\d{2}\/\d{4})\s+\d{2}:\d{2}/i);
        if (de) data.dataEmissao = de[1];
        if (!data.dataEmissao) { de = fullText.match(/(\d{2}\/\d{2}\/\d{4})\s+\d{2}:\d{2}/); if (de) data.dataEmissao = de[1]; }
        if (!data.dataEmissao) { const d = fullText.match(/(\d{2}\/\d{2}\/\d{4})/); if (d) data.dataEmissao = d[1]; }

        return data;
}

function extractSaoPaulo(lines, filename) {
    let data = { arquivo: filename, modelo: "São Paulo", status: "Sucesso", prestador: "Não identificado", cnpj: "", vLiquido: 0, vServico: 0, dataEmissao: "", numeroNF: "?" };
    const fullText = lines.join("\n");

    for(let i=0; i<lines.length; i++) {
        if(lines[i].includes("Número da Nota") && i+1 < lines.length) {
            const match = lines[i+1].trim().match(/(\d+)/);
            if(match) data.numeroNF = parseInt(match[1]); 
        }
    }

    for(let line of lines) {
        if(line.includes("Nome/Razão Social:")) {
            let name = line.split(":")[1].trim();
            if(!name.toUpperCase().includes("FREXCO")) {
                data.prestador = name;
                break;
            }
        }
    }

    for(let i=0; i<lines.length; i++) {
        if(lines[i].includes("CPF/CNPJ:")) {
             let match = lines[i].match(/(\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2})/);
             if(match) { data.cnpj = match[1]; break; }
             if(i+1 < lines.length) {
                 match = lines[i+1].match(/(\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2})/);
                 if(match) { data.cnpj = match[1]; break; }
             }
        }
    }

    const valMatch = fullText.match(/VALOR\s*TOTAL\s*DO\s*SERVIÇO\s*=\s*R\$\s*([\d.,]+)/i);
    if(valMatch) {
        data.vServico = parseFloat(valMatch[1].replace(/\./g,'').replace(',','.'));
        data.vLiquido = data.vServico;
    }

    for(let i=0; i<lines.length; i++) {
        if((lines[i].includes("Data e Hora de Emissão") || lines[i].includes("Data e Hora da Emissão")) && i+1 < lines.length) {
            const dateMatch = lines[i+1].match(/(\d{2}\/\d{2}\/\d{4})/);
            if(dateMatch) data.dataEmissao = dateMatch[1];
        }
    }
    
    if(!data.dataEmissao) {
        for(let i=0; i<lines.length; i++) {
            const lineUpper = lines[i].toUpperCase();
            if(lineUpper.includes("DATA") && (lineUpper.includes("HORA") || lineUpper.includes("EMISS"))) {
                for(let k=0; k<=2; k++) {
                    if(i+k < lines.length) {
                        const dateMatch = lines[i+k].match(/(\d{2}\/\d{2}\/\d{4})/);
                        if(dateMatch) {
                            data.dataEmissao = dateMatch[1];
                            break;
                        }
                    }
                }
                if(data.dataEmissao) break;
            }
        }
    }
    
    if(!data.dataEmissao) {
        const firstDate = fullText.match(/(\d{2}\/\d{2}\/\d{4})/);
        if(firstDate) data.dataEmissao = firstDate[1];
    }

    return data;
}

function extractStandardV7(lines, filename) {
    let data = { arquivo: filename, modelo: "Nacional", status: "Sucesso", prestador: "Não identificado", cnpj: "", vLiquido: 0, vServico: 0, dataEmissao: "", numeroNF: "?" };
    const fullText = lines.join("\n");

    const cnpjs = fullText.match(/\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}/g) || [];
    data.cnpj = cnpjs[0] || "";
    if (data.cnpj.includes("33.935.782") && cnpjs.length > 1) data.cnpj = cnpjs[1];

    const moneyRegex = /R\$\s?(\d{1,3}(?:\.\d{3})*,\d{2})/g;
    let amounts = [];
    let match;
    while ((match = moneyRegex.exec(fullText)) !== null) amounts.push(parseFloat(match[1].replace(/\./g, '').replace(',', '.')));
    if (amounts.length > 0) { amounts.sort((a, b) => b - a); data.vServico = amounts[0]; data.vLiquido = amounts[0]; }

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        const lineUpper = line.toUpperCase();

        if (lineUpper.includes("NÚMERO DA NFS-E") || lineUpper.includes("NÚMERO DA NOTA")) {
            const matchSame = line.match(/[:.]?\s*(\d{1,9})/);
            if (matchSame && matchSame[1].length < 9) data.numeroNF = matchSame[1];
            else if (i + 1 < lines.length) {
                const nextLine = lines[i+1].trim();
                const matchNext = nextLine.match(/^(\d{1,9})/);
                if (matchNext && matchNext[1] !== "2024" && matchNext[1] !== "2025") data.numeroNF = matchNext[1];
            }
        }

        if (lineUpper.includes("PRESTADOR DO SERVIÇO")) {
            for (let j = 1; j <= 4; j++) {
                if (i + j >= lines.length) break;
                let cand = lines[i + j].trim();
                if (!cand.toUpperCase().includes("NOME") && !cand.toUpperCase().includes("CNPJ") && cand.length > 3) {
                    let name = cand.replace(/\d{14}/,"").trim();
                    name = name.replace(/[\w._%+-]+@[\w.-]+\.[a-zA-Z]{2,}/g, '').trim();
                    name = name.replace(/\s+-\s*$/g, '').trim();
                    data.prestador = name;
                    break;
                }
            }
        }

        if ((lineUpper.includes("EMISSÃO") || lineUpper.includes("COMPETÊNCIA")) && !data.dataEmissao) {
            const dateRegex = /(\d{2}\/\d{2}\/\d{4})/;
            let match = line.match(dateRegex);
            if (match) data.dataEmissao = match[1];
            else if (i + 1 < lines.length) {
                match = lines[i+1].match(dateRegex);
                if (match) data.dataEmissao = match[1];
            }
        }
    }
    return data;
}
