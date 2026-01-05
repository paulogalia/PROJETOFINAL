/**
 * ============================================================================
 * EUM MANAGER 6.6 - CLEAN DATA EDITION
 * Inclui:
 * 1. Limpeza Adaptativa (Sua ideia): Remove datas/horas misturadas nos valores.
 * 2. Smart Dates: Lê datas ISO (YYYY-MM-DD) e BR (DD/MM/YYYY).
 * 3. Engine 6.0: Suporta campos virtuais e Oráculo.
 * ============================================================================
 */

const GEMINI_API_KEY = "AIzaSyDI_74u6OX5--6ZjOUlbJuqQUf5f46pP-8"; 
const MASTER_DB_ID = "1zKqYVR9seTPy3eyX5CR2xuXZtqSwwvERQn_KX3GK5JM"; // Base HCPA
// Lista de segurança caso a API de descoberta falhe
const FALLBACK_MODELS = ["gemini-2.5-flash","gemini-2.0-flash-lite", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"];

// --- UI & NAVEGAÇÃO ---
function onOpen() { SpreadsheetApp.getUi().createMenu('🚀 EUM App').addItem('Abrir Painel', 'abrirDashboard').addToUi(); }

function abrirDashboard() { 
  var html = HtmlService.createTemplateFromFile('LabHome')
      .evaluate()
      .setTitle('EUM Manager 6.6')
      .setWidth(1200).setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'EUM Manager');
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

// --- CONFIGURAÇÃO & ESTADO ---
function apiGetInitialState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(s => s.getName()).filter(n => !['Exames_Referencia','Doses_Referencia','Config','Dashboard'].includes(n));
  const savedConfig = JSON.parse(PropertiesService.getScriptProperties().getProperty('EUM_CONFIG_MASTER') || '{}');
  const hasRef = !!ss.getSheetByName("Exames_Referencia");
  return { 
    sheets: sheets, 
    config: savedConfig, 
    status: { hasCore: !!savedConfig.core?.abaFatos, hasExames: !!savedConfig.exames?.active, hasReferencias: hasRef } 
  };
}

function apiSaveConfig(newConfig) {
  try {
    PropertiesService.getScriptProperties().setProperty('EUM_CONFIG_MASTER', JSON.stringify(newConfig));
    return { sucesso: true };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

// ============================================================================
// ★ MÓDULO 1: WIZARD IA (Mapeamento + Design de Dashboard)
// ============================================================================
function apiMagicSetup(pdfBase64, matrixDados, nomeArquivo, abaExistente) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let headers = [], sheetName = "";

  if (abaExistente) {
    const sheet = ss.getSheetByName(abaExistente);
    if (!sheet) return { erro: "Aba não encontrada." };
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    sheetName = abaExistente;
  } else if (matrixDados && matrixDados.length > 0) {
    sheetName = "Dados_" + (nomeArquivo || "Import").replace(/[^a-zA-Z0-9]/g, "_").substring(0, 15);
    if (ss.getSheetByName(sheetName)) sheetName += "_" + Math.floor(Math.random()*1000);
    const sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, matrixDados.length, matrixDados[0].length).setValues(matrixDados);
    headers = matrixDados[0];
  } else { return { erro: "Sem dados." }; }

  // Prompt Expandido: Pede Mapeamento E Sugestões de Gráficos
  const prompt = `
    ATUE COMO: Clinical Data Scientist.
    INPUTS: 
    1. Cabeçalhos do Excel: ${JSON.stringify(headers)}.
    2. Documento Anexo: PDF do Protocolo Clínico.

    TAREFA 1 (Mapeamento):
    Associe as colunas do Excel às variáveis clínicas padrão.
    Chaves possíveis: "colProntFatos", "colMed", "colDtIni", "colDose24h", "colPeso", "colCreat", "colNasc", "colSexo".

    TAREFA 2 (Smart Dashboard):
    Com base no objetivo do estudo (PDF), sugira 4 gráficos essenciais para o painel inicial.
    - Use as chaves do mapeamento (ex: "colCreat") ou colunas reais do Excel para os eixos.
    - Tipos permitidos: "bar", "pie", "scatter", "histogram", "box".
    - Se o estudo focar em evolução temporal, sugira "scatter" com datas.
    - Se focar em distribuição, sugira "histogram" ou "box".

    RETORNE JSON ESTRITO:
    {
      "studyName": "Nome Curto do Estudo",
      "studySummary": "Resumo de 1 linha",
      "mapping": { "colProntFatos": "...", "colMed": "...", ... },
      "smartCharts": [
        { "title": "Distribuição por Sexo", "type": "pie", "x": "colSexo", "y": null },
        { "title": "Correlação Dose vs Creatinina", "type": "scatter", "x": "colDose24h", "y": "colCreat" },
        { "title": "Evolução Temporal", "type": "line", "x": "colDtIni", "y": null },
        { "title": "Ranking de Medicamentos", "type": "bar", "x": "colMed", "y": null }
      ]
    }`;

  const payload = { contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: "application/pdf", data: pdfBase64 } }] }] };
  
  try {
    const txt = callGeminiSmart_(payload);
    const resIA = cleanJson_(txt);
    resIA.createdSheet = sheetName;
    
    // Salva a sugestão da IA na configuração global para o Dashboard usar depois
    const currentConfig = JSON.parse(PropertiesService.getScriptProperties().getProperty('EUM_CONFIG_MASTER') || '{}');
    currentConfig.smartCharts = resIA.smartCharts; // Guarda os gráficos sugeridos
    PropertiesService.getScriptProperties().setProperty('EUM_CONFIG_MASTER', JSON.stringify(currentConfig));

    return resIA;
  } catch (e) { return { erro: e.message }; }
}

// ============================================================================
// ★ FUNÇÃO ORÁCULO (INTELIGÊNCIA DO EXPLORADOR)
// Certifique-se que o nome é 'apiOracleInterpret' para bater com o HTML.
// ============================================================================

function apiOracleInterpret(pergunta, sheetName) {
  try {
    // 1. Pega as colunas reais da planilha
    let cols = apiGetColumns(sheetName);
    
    // 2. ★ INJEÇÃO CRÍTICA: Adiciona variáveis virtuais que o sistema calcula sozinho
    // Se não fizer isso, a IA não sabe que '_IDADE_ANOS' existe e o gráfico falha.
    cols.push("_IDADE_ANOS", "_PID", "DT_REFERENCIA", "DT_NASCIMENTO"); 
    
    const prompt = `
    Atue como Data Scientist especialista em Plotly.js.
    PERGUNTA DO USUÁRIO: "${pergunta}"
    COLUNAS DISPONÍVEIS: ${JSON.stringify(cols)}
    
    REGRAS ESTRITAS DE RESPOSTA (JSON):
    1. Retorne APENAS um objeto JSON válido. Sem markdown, sem texto antes/depois.
    2. Use "aggregation" para definir a operação: "count", "sum", "avg", "max", "min".
    3. Se a pergunta for "Quantos..." ou distribuição, defina "yAxis": null e "aggregation": "count".
    4. Se a pergunta implicar pacientes únicos (ex: "quantos pacientes"), defina "group": true.
    5. Para idade, use SEMPRE a coluna "_IDADE_ANOS".
    
    FORMATO ESPERADO:
    {
      "xAxis": "NomeDaColunaX",
      "yAxis": "NomeDaColunaY", 
      "chartType": "bar" | "pie" | "scatter" | "box",
      "aggregation": "count", 
      "group": true,
      "filter": "row['_IDADE_ANOS'] >= 40", 
      "explanation": "Explicação curta para o usuário"
    }`;
    
    // Chama o Gemini (Smart ou Direct, dependendo do que tiver configurado)
    // Se tiver a função 'callGeminiSmart_', use-a. Se não, use 'callGeminiDirect_'
    const txt = callGeminiSmart_({ contents: [{ parts: [{ text: prompt }] }] });
    
    // Limpeza de JSON Blindada
    const cleanTxt = txt.replace(/```json/gi, "").replace(/```/g, "").trim();
    const jsonStart = cleanTxt.indexOf('{');
    const jsonEnd = cleanTxt.lastIndexOf('}');
    
    if (jsonStart === -1 || jsonEnd === -1) throw new Error("IA não retornou JSON.");
    
    return { sucesso: true, config: JSON.parse(cleanTxt.substring(jsonStart, jsonEnd + 1)) };
    
  } catch(e) { 
    return { sucesso: false, erro: "Oráculo: " + e.message }; 
  }
}

// ============================================================================
// ★ MÓDULO 3: ENGINE DE DADOS 6.0 (COM LIMPEZA)
// ============================================================================
function apiGetRawExplorerData(sheetName, cols, filterScript, virtualFields = []) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { sucesso: false, erro: "Aba não encontrada." };
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toUpperCase().trim());
  const colMap = {}; headers.forEach((h, i) => colMap[h] = i);
  
  const config = JSON.parse(PropertiesService.getScriptProperties().getProperty('EUM_CONFIG_MASTER') || '{}');
  const idxPront = config.core?.colProntFatos ? colMap[config.core.colProntFatos.toUpperCase()] : -1;
  let idxNasc = config.core?.colNasc ? colMap[config.core.colNasc.toUpperCase()] : headers.findIndex(h => h.includes("NASC"));
  let idxDataRef = config.core?.colDtIni ? colMap[config.core.colDtIni.toUpperCase()] : headers.findIndex(h => h.includes("DATA"));

  // Compiladores
  let filterFunc = null;
  const virtualFuncs = [];
  try {
    if (virtualFields && Array.isArray(virtualFields)) virtualFields.forEach(vf => virtualFuncs.push({ name: vf.name, func: new Function('row', 'return ' + vf.logic) }));
    if (filterScript && filterScript !== "true") filterFunc = new Function('row', 'return ' + filterScript);
  } catch (e) { return { sucesso: false, erro: "Erro Lógica: " + e.message }; }

  const result = [];
  
  // Loop Principal
  for (let i = 1; i < data.length; i++) {
    const row = data[i], rowContext = {}, rowObj = {};
    
    // 1. LEITURA COM LIMPEZA (Sua ideia aplicada)
    headers.forEach((h, idx) => rowContext[h] = cleanHospitalValue_(row[idx]));
    
    // 2. Derivados
    const pid = (idxPront > -1) ? String(row[idxPront]).trim() : "ID_"+i;
    rowContext['_PID'] = pid; rowObj['_PID'] = pid; 

    // Datas Inteligentes
    const dtNasc = idxNasc > -1 ? parseSmartDate_(row[idxNasc]) : null;
    const dtRef = idxDataRef > -1 ? parseSmartDate_(row[idxDataRef]) : new Date();
    rowContext['_IDADE_ANOS'] = (dtNasc && dtRef) ? parseFloat(((dtRef-dtNasc)/(365.25*86400000)).toFixed(2)) : 0;

    // 3. Virtuais & Filtros
    try { virtualFuncs.forEach(vf => rowContext[vf.name] = vf.func(rowContext)); } catch(e) { continue; }
    if (filterFunc) { try { if (!filterFunc(rowContext)) continue; } catch(e) { continue; } }

    // 4. Seleção
    let hasData = false;
    cols.forEach(c => {
      if (rowContext[c] !== undefined) { rowObj[c] = rowContext[c]; hasData = true; } 
      else { rowObj[c] = null; }
    });
    rowObj['_IDADE_ANOS'] = rowContext['_IDADE_ANOS'];

    if(hasData) result.push(rowObj);
  }
  return { sucesso: true, dados: result };
}

// ============================================================================
// ★ MÓDULO 4: SEGURANÇA (RAMs)
// ============================================================================
function apiGetSegurancaData(params) {
  const state = apiGetInitialState(); 
  if (!state.status.hasRams) return { sucesso: false, erro: "Sem RAMs." };
  try {
    const c = state.config.core, r = state.config.rams, ss = SpreadsheetApp.getActiveSpreadsheet();
    const fats = ss.getSheetByName(c.abaFatos).getDataRange().getValues(), hF = fats.shift();
    const iPF = hF.indexOf(c.colProntFatos), iMed = hF.indexOf(c.colMed), iDI = hF.indexOf(c.colDtIni);
    const coorte = {}; 
    fats.forEach(row => { 
      if (row[iMed] === params.medicamento) { 
        const p = String(row[iPF]).trim(), d = parseSmartDate_(row[iDI]); 
        if (d && (!coorte[p] || d < coorte[p])) coorte[p] = d; 
      }
    });
    
    const rams = ss.getSheetByName(r.aba).getDataRange().getValues(), hR = rams.shift();
    const iG = hR.indexOf(r.colGrav), iC = hR.indexOf(r.colCaus), iPR = hR.findIndex(k => k.includes("PRONT")), iDtR = hR.findIndex(k => k.includes("DATA"));
    
    const evts = []; 
    rams.forEach(row => {
      const p = String(row[iPR]).trim();
      if (coorte[p]) { 
        const dt = parseSmartDate_(row[iDtR]);
        const dias = dt ? Math.floor((dt - coorte[p]) / 86400000) : null; 
        evts.push({ gravidade: String(row[iG]||"N/D"), causalidade: String(row[iC]||"N/D"), dias: dias, prontuario: p }); 
      }
    });
    return { sucesso: true, dados: evts, totalExpostos: Object.keys(coorte).length };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

// ============================================================================
// HELPERS GERAIS
// ============================================================================
function callGeminiSmart_(payload) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODELO_FIXO}:generateContent?key=${GEMINI_API_KEY}`;
  const resp = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error("Gemini: " + resp.getContentText());
  return JSON.parse(resp.getContentText()).candidates[0].content.parts[0].text;
}
function cleanJson_(text) {
  try {
    // Remove marcadores de código (```json, ```JSON, ```)
    let clean = text.replace(/```json/gi, "").replace(/```/g, "").trim();
    
    // Encontra o bloco JSON real (entre o primeiro { e o último })
    const start = clean.indexOf('{');
    const end = clean.lastIndexOf('}');
    
    if (start === -1 || end === -1) throw new Error("A IA não retornou um objeto JSON válido.");
    
    // Extrai e faz o parse
    clean = clean.substring(start, end + 1);
    return JSON.parse(clean);
  } catch (e) {
    throw new Error("Erro de Sintaxe no JSON da IA: " + e.message);
  }
}
function apiGetColumns(s) { const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(s); return sh?sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>String(h).toUpperCase().trim()):[]; }
function apiGetMedicamentos(s,c) { return apiGetUniqueValues(s,c); }
function apiGetUniqueValues(s,c) { const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(s); if(!sh)return[]; const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v).toUpperCase().trim()); const idx=h.indexOf(String(c).toUpperCase().trim()); if(idx<0)return[]; const v=sh.getRange(2,idx+1,sh.getLastRow()-1,1).getValues().flat(); return [...new Set(v.filter(String))].sort(); }
function parseSmartDate_(d) {
  if (!d) return null; if (d instanceof Date) return d;
  const str = String(d).trim(); if (str === "") return null;
  let p = str.split(/[\/\-\.\sT]/).filter(i => i !== "");
  if (p.length >= 3) {
    const v1=parseInt(p[0]), v2=parseInt(p[1])-1, v3=parseInt(p[2]);
    if (v1 > 31) { const dt=new Date(v1, v2, v3); if(!isNaN(dt.getTime())) return dt; } // ISO
    if (v3 > 31) { const dt=new Date(v3, v2, v1); if(!isNaN(dt.getTime())) return dt; } // BR
  }
  const nDt = new Date(str); return isNaN(nDt.getTime()) ? null : nDt;
}
// Alias para manter compatibilidade
function parseDatePTBR_(d) { return parseSmartDate_(d); }

// ============================================================================
// ★ HELPER: LIMPEZA ADAPTATIVA (SUA IDEIA IMPLEMENTADA)
// ============================================================================
function cleanHospitalValue_(val) {
  // Se já for numérico puro, devolve
  if (typeof val === 'number') return val;
  if (!val) return null;

  let str = String(val).trim();

  // 1. Remove Datas (DD/MM/YY, YYYY-MM-DD, etc)
  // Regex procura sequências de números separados por / ou -
  str = str.replace(/\d{2,4}[\/\-]\d{1,2}[\/\-]\d{2,4}/g, "").trim();

  // 2. Remove Horas (HH:MM:SS)
  str = str.replace(/\d{1,2}:\d{2}(?::\d{2})?/g, "").trim();

  // 3. Normaliza Decimal (7,8 -> 7.8)
  str = str.replace(',', '.');

  // 4. Converte e Valida
  const num = parseFloat(str);
  
  // Se depois de limpar sobrou texto (ex: "Reagente"), devolve o texto original
  // Se sobrou um número, devolve o número limpo.
  return isNaN(num) ? val : num;
}

// Funções de Referência e Importação (Mantidas para compatibilidade)
function apiImportarReferencias() {
  try {
    const targetSS = SpreadsheetApp.getActiveSpreadsheet();
    const masterSS = SpreadsheetApp.openById(MASTER_DB_ID);
    let sourceSheet = masterSS.getSheetByName("Exames_Referencia") || masterSS.getSheets()[0];
    if (!sourceSheet) return { sucesso: false, erro: "Mestre vazio." };
    const data = sourceSheet.getDataRange().getValues();
    let targetSheet = targetSS.getSheetByName("Exames_Referencia");
    if (targetSheet) targetSheet.clear(); else { targetSheet = targetSS.insertSheet("Exames_Referencia"); targetSheet.hideSheet(); }
    targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    return { sucesso: true, total: data.length - 1 };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

// ============================================================================
// ★ MÓDULO 5: FARMACOVIGILÂNCIA (TIMELINE DO PACIENTE)
// ============================================================================
function apiGetPatientTimeline(pid) {
  const state = apiGetInitialState();
  if (!state.status.hasCore) return { sucesso: false, erro: "Sem dados carregados." };

  const config = state.config.core;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(config.abaFatos);
  
  if (!sheet) return { sucesso: false, erro: "Aba de dados não encontrada." };

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toUpperCase().trim());
  
  // Mapeamento de Colunas (Inteligente)
  const colMap = {}; headers.forEach((h, i) => colMap[h] = i);
  
  // Identifica índices vitais
  const idxPront = config.colProntFatos ? colMap[config.colProntFatos.toUpperCase()] : -1;
  const idxData = config.colDtIni ? colMap[config.colDtIni.toUpperCase()] : headers.findIndex(h => h.includes("DATA") || h.includes("DT_"));
  const idxMed = config.colMed ? colMap[config.colMed.toUpperCase()] : headers.findIndex(h => h.includes("MEDICAMENTO"));
  const idxDose = config.colDose24h ? colMap[config.colDose24h.toUpperCase()] : headers.findIndex(h => h.includes("DOSE"));
  
  // Identifica Exames (Procura por nomes comuns no cabeçalho)
  const idxCreat = headers.findIndex(h => h.includes("CREAT") || h.includes("CKD"));
  const idxLeuco = headers.findIndex(h => h.includes("LEUCOCITOS") || h.includes("WBC"));
  const idxPlaq = headers.findIndex(h => h.includes("PLAQUETAS") || h.includes("PLAT"));
  const idxAST = headers.findIndex(h => h.includes("AST") || h.includes("TGO"));

  if (idxPront === -1) return { sucesso: false, erro: "Coluna de Prontuário não mapeada." };

  const timeline = { meds: [], exams: [] };
  const targetPID = String(pid).trim();

  // Varredura
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowPID = String(row[idxPront]).trim();

    if (rowPID === targetPID) {
      const dt = parseSmartDate_(row[idxData]);
      if (!dt) continue;
      
      const dateStr = dt.toISOString().split('T')[0]; // YYYY-MM-DD

      // 1. Medicamento
      if (idxMed > -1) {
        const medName = String(row[idxMed]);
        const doseVal = cleanHospitalValue_(row[idxDose]);
        if (doseVal > 0) {
          timeline.meds.push({ date: dateStr, name: medName, dose: doseVal });
        }
      }

      // 2. Exames (Se existirem na linha)
      const addExam = (idx, type) => {
        if (idx > -1) {
          const val = cleanHospitalValue_(row[idx]);
          if (val && typeof val === 'number') {
            timeline.exams.push({ date: dateStr, type: type, value: val });
          }
        }
      };

      addExam(idxCreat, 'Creatinina');
      addExam(idxLeuco, 'Leucócitos');
      addExam(idxPlaq, 'Plaquetas');
      addExam(idxAST, 'AST/TGO');
    }
  }

  // Ordenação
  timeline.meds.sort((a, b) => a.date.localeCompare(b.date));
  timeline.exams.sort((a, b) => a.date.localeCompare(b.date));

  return { sucesso: true, data: timeline };
}

// ============================================================================
// ★ MOTOR DE DESCOBERTA DE MODELOS (AUTO-UPDATE)
// 1. Consulta a API para ver o que existe.
// 2. Filtra apenas modelos de texto.
// 3. Classifica (Ranking) por versão e velocidade.
// 4. Salva na memória (Persistência).
// ============================================================================

function getAvailableModels_() {
  // Tenta ler do cache (Memória Persistente)
  const cached = PropertiesService.getScriptProperties().getProperty('EUM_MODEL_LIST');
  if (cached) return JSON.parse(cached);

  // Se não tiver, aciona o Scanner
  return forceModelUpdate();
}

function forceModelUpdate() {
  try {
    // 1. Scanner (GET /models)
    const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${GEMINI_API_KEY}`;
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    
    if (resp.getResponseCode() !== 200) throw new Error("Falha ao listar modelos");

    const json = JSON.parse(resp.getContentText());
    
    // 2. Filtro de Inteligência & 3. Ranking Automático
    const models = json.models
      .filter(m => m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent")) // Apenas texto/multimodal
      .map(m => m.name.replace("models/", "")) // Limpa nome
      .sort((a, b) => {
        // Extrai versão (ex: 2.0, 1.5)
        const getVer = (s) => parseFloat(s.match(/(\d+\.\d+)/)?.[0] || 0);
        const vA = getVer(a), vB = getVer(b);
        
        // Regras de Ranking:
        if (vA !== vB) return vB - vA; // Versão maior primeiro (2.5 > 2.0)
        if (a.includes("flash") && !b.includes("flash")) return -1; // Flash ganha de Pro (Velocidade)
        if (a.includes("live") && !b.includes("live")) return -1; // Live ganha (Mais recente)
        return 0;
      });

    // 4. Memória Persistente
    PropertiesService.getScriptProperties().setProperty('EUM_MODEL_LIST', JSON.stringify(models));
    console.log("Modelos Atualizados e Rankeados:", models);
    return models;

  } catch (e) {
    console.warn("Erro no Auto-Update, usando fallback:", e);
    return FALLBACK_MODELS;
  }
}

// ============================================================================
// ★ SMART CALLER (Usa a lista gerada acima)
// ============================================================================
function callGeminiSmart_(payload) {
  const modelList = getAvailableModels_(); // Pega a lista "de ouro"
  let lastError = null;

  // Tenta em cascata: O Top 1 primeiro, se falhar, o próximo
  for (const model of modelList) {
    // Ignora modelos muito antigos para não perder tempo
    if (model.includes("1.0")) continue; 

    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${GEMINI_API_KEY}`;
    
    try {
      console.log(`Tentando IA: ${model}...`);
      const resp = UrlFetchApp.fetch(url, { 
        method: 'post', 
        contentType: 'application/json', 
        payload: JSON.stringify(payload), 
        muteHttpExceptions: true 
      });
      
      const code = resp.getResponseCode();
      
      if (code === 200) {
        return JSON.parse(resp.getContentText()).candidates[0].content.parts[0].text;
      }
      
      // Se for erro de Cota (429) ou Server (503), tenta o próximo da lista
      if (code === 429 || code === 503) {
        console.warn(`⚠️ ${model} indisponível (${code}). Roteando para o próximo...`);
        continue; 
      }
      
      throw new Error(`Erro API (${code}): ${resp.getContentText()}`);

    } catch (e) {
      lastError = e.message;
      if (e.message.includes("Fatal")) throw e;
    }
  }
  throw new Error(`Todas as IAs falharam. Verifique sua chave API. Erro: ${lastError}`);
}

