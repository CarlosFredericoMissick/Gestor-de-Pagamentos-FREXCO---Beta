// ========================================
// ESTADO GLOBAL COMPARTILHADO ENTRE OS MÓDULOS
// (scripts clássicos: `let` de topo é visível em todos os scripts carregados depois)
// ========================================

// Extração de notas
let filesQueue = [];
let extractedResults = []; // notas em memória; fonte de verdade é o SharePoint (notas/*.json)
let isProcessing = false;

// ========================================
// PRESTADORES — DADOS UNIFICADOS
// Estrutura: {
//   id, vehicleId (= driver_id após migração), nome, cd, time, inUse,
//   vehicles: [{ vehicleId, nome }],  ← todos os veículos do motorista
//   cnpjs: [{ cnpj, chavePix }],
//   contrato: { nomeArquivo, fileData, dataUpload } | null,
//   fromSharePoint, dataAdicao
// }
// Líderes ficam na mesma lista com tipo: 'lider', timeVinculado e motoristasVinculados.
// Regras: CNPJ único global, 1 contrato por motorista, 1 registro por driver_id
// ========================================
let prestadoresDB = null;
let prestadoresList = []; // fonte única de verdade
let recebivelData = [];   // planilha de recebível importada
let _prestadoresTab = 'motoristas'; // aba ativa no modal

// SharePoint
let useSharePoint = false;
let spAccessToken = null;
let spAccount = null;
let spExigeReconexao = false; // renovação silenciosa do token pediu login interativo

// Filtros de time (vazio = todos)
let _filtroTimesAtivos = new Set();
let _concilTimesAtivos = new Set();
