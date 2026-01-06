/**
 * ============================================================================
 * EUM MANAGER 6.9 - SELF-HEALING SYNTAX EDITION
 * * Histórico de Mudanças:
 * - v6.9: Auto-correção de sintaxe IA (and->&&, or->||) e proteção contra erros lógicos.
 * - v6.8: Suporte a Eixos Calculados e Matemática Virtual.
 * - v6.7: Proteção de Datas.
 * ============================================================================
 */

// SEGURANÇA: Chave movida para Script Properties
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'); 

const MASTER_DB_ID = "1zKqYVR9seTPy3eyX5CR2xuXZtqSwwvERQn_KX3GK5JM"; // Base HCPA (Exemplo)
// Lista de segurança caso a API de descoberta falhe
const FALLBACK_MODELS = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"];

// --- UI & NAVEGAÇÃO ---
function onOpen() { SpreadsheetApp.getUi().createMenu('🚀 EUM App').addItem('Abrir Painel', 'abrirDashboard').addToUi(); }

function abrirDashboard() { 
  var html = HtmlService.createTemplateFromFile('LabHome')
      .evaluate()
      .setTitle('EUM Manager 6.9')
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
  } catch (e) { return { sucesso: false, erro: e.message };
  }
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
    - Tipos permitidos: "bar", "pie", "scatter", "histogram", "line", "box".
    
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
  
  try {
    const txt = callGeminiSmart_({ contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: "application/pdf", data: pdfBase64 } }] }] });
    const resIA = cleanJson_(txt);
    resIA.createdSheet = sheetName;
    
    // Salva a sugestão da IA na configuração global
    const currentConfig = JSON.parse(PropertiesService.getScriptProperties().getProperty('EUM_CONFIG_MASTER') || '{}');
    currentConfig.smartCharts = resIA.smartCharts; 
    PropertiesService.getScriptProperties().setProperty('EUM_CONFIG_MASTER', JSON.stringify(currentConfig));

    return resIA;
  } catch (e) { return { erro: e.message }; }
}

// ============================================================================
// ★ FUNÇÃO ORÁCULO (INTELIGÊNCIA DO EXPLORADOR - V3 COM CORREÇÃO)
// ============================================================================

function apiOracleInterpret(pergunta, sheetName) {
  try {
    let cols = apiGetColumns(sheetName);
    cols.push("_IDADE_ANOS", "_PID", "DT_REFERENCIA", "DT_NASCIMENTO"); 
    
    const dateCols = cols.filter(c => c.includes('DT_') || c.includes('DATA') || c.includes('NASC'));

    const prompt = `
    Atue como Data Scientist especialista em Plotly.js e JavaScript.
    PERGUNTA DO USUÁRIO: "${pergunta}"
    COLUNAS DISPONÍVEIS: ${JSON.stringify(cols)}
    COLUNAS DE DATA: ${JSON.stringify(dateCols)}
    
    REGRAS ESTRITAS DE CÓDIGO:
    1. USE SINTAXE JAVASCRIPT! Use '&&' para E, '||' para OU. NÃO use 'and', 'or'.
    2. Strings devem estar entre aspas simples ou duplas.
    
    REGRAS DE RESPOSTA (JSON):
    1. Se a pergunta for sobre "evolução", "temporal", USE uma coluna de data no "xAxis" e "chartType": "line".
    2. Se a pergunta exigir CÁLCULO (ex: "Dose por Peso"), CRIE um campo virtual.
       - Exemplo: { "name": "_DOSE_KG", "logic": "row['DOSE_24HS'] / row['PESO']" }
       - Use o nome criado (ex: "_DOSE_KG") no "yAxis" e no filtro se necessário.
    3. Se a pergunta for "Quantos pacientes...", DEFINA "group": true e "yAxis": null.
    4. Para idade, use SEMPRE "_IDADE_ANOS".
    
    FORMATO ESPERADO:
    {
      "xAxis": "NomeDaColunaX",
      "yAxis": "NomeDaColunaY",
      "chartType": "bar" | "pie" | "scatter" | "line" | "box",
      "aggregation": "count" | "avg" | "sum", 
      "group": true,
      "virtualFields": [ { "name": "_DOSE_KG", "logic": "row['DOSE_24HS'] / row['PESO']" } ], 
      "filter": "row['_DOSE_KG'] > 20 && row['_IDADE_ANOS'] < 12", 
      "explanation": "Filtrado por Dose/Peso > 20 e Idade < 12."
    }`;
    
    const txt = callGeminiSmart_({ contents: [{ parts: [{ text: prompt }] }] });
    return { sucesso: true, config: cleanJson_(txt) };
  } catch(e) { 
    return { sucesso: false, erro: "Oráculo: " + e.message };
  }
}

// ============================================================================
// ★ MÓDULO 3: ENGINE DE DADOS 6.9 (SELF-HEALING SYNTAX)
// ============================================================================
function apiGetRawExplorerData(sheetName, cols, filterScript, virtualFields = []) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { sucesso: false, erro: "Aba não encontrada." };
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toUpperCase().trim());
  const dateColIndexes = headers.map((h, i) => (h.includes('DATA') || h.includes('DT_') || h.includes('NASC')) ? i : -1).filter(i => i !== -1);

  const config = JSON.parse(PropertiesService.getScriptProperties().getProperty('EUM_CONFIG_MASTER') || '{}');
  const idxPront = config.core?.colProntFatos ? headers.indexOf(config.core.colProntFatos.toUpperCase()) : -1;
  let idxNasc = config.core?.colNasc ? headers.indexOf(config.core.colNasc.toUpperCase()) : headers.findIndex(h => h.includes("NASC"));
  let idxDataRef = config.core?.colDtIni ? headers.indexOf(config.core.colDtIni.toUpperCase()) : headers.findIndex(h => h.includes("DATA"));
  
  // --- COMPILADORES COM AUTO-CORREÇÃO (A MÁGICA ESTÁ AQUI) ---
  let filterFunc = null;
  const virtualFuncs = [];
  
  try {
    // 1. Compila Campos Virtuais
    if (virtualFields && Array.isArray(virtualFields)) {
       virtualFields.forEach(vf => {
         // Sanitiza lógica do campo virtual também
         let logic = sanitizeJs_(vf.logic);
         virtualFuncs.push({ name: vf.name, func: new Function('row', 'return ' + logic) });
       });
    }

    // 2. Compila Filtro com Sanitização
    if (filterScript && filterScript !== "true") {
       let safeFilter = sanitizeJs_(filterScript);
       console.log("Filtro Original:", filterScript, " | Filtro Sanitizado:", safeFilter); // Log para debug
       filterFunc = new Function('row', 'return ' + safeFilter);
    }
  } catch (e) { 
    return { sucesso: false, erro: "Erro de Sintaxe (Auto-Correção falhou): " + e.message };
  }

  const result = [];
  
  // Loop Principal
  for (let i = 1; i < data.length; i++) {
    const row = data[i], rowContext = {}, rowObj = {};
    
    // Leitura
    headers.forEach((h, idx) => {
      if (dateColIndexes.includes(idx)) rowContext[h] = row[idx];
      else rowContext[h] = cleanHospitalValue_(row[idx]);
    });
    
    // Derivados
    const pid = (idxPront > -1) ? String(row[idxPront]).trim() : "ID_"+i;
    rowContext['_PID'] = pid; rowObj['_PID'] = pid; 

    const dtNasc = idxNasc > -1 ? parseSmartDate_(row[idxNasc]) : null;
    const dtRef = idxDataRef > -1 ? parseSmartDate_(row[idxDataRef]) : new Date();
    rowContext['_IDADE_ANOS'] = (dtNasc && dtRef) ? parseFloat(((dtRef-dtNasc)/(365.25*86400000)).toFixed(2)) : 0;
    rowContext['DT_REFERENCIA'] = dtRef;

    // Execução Virtual (Matemática)
    try { 
        virtualFuncs.forEach(vf => rowContext[vf.name] = vf.func(rowContext)); 
    } catch(e) { continue; } // Pula linha se cálculo falhar (ex: div por zero)

    // Execução do Filtro
    if (filterFunc) { 
        try { if (!filterFunc(rowContext)) continue; } catch(e) { continue; } 
    }

    // Seleção Final
    let hasData = false;
    cols.forEach(c => {
      if (rowContext[c] !== undefined) { rowObj[c] = rowContext[c]; hasData = true; } 
      else { rowObj[c] = null; }
    });
    rowObj['_IDADE_ANOS'] = rowContext['_IDADE_ANOS'];

    if(cols.includes('DT_REFERENCIA') && rowContext['DT_REFERENCIA'] instanceof Date) {
       rowObj['DT_REFERENCIA'] = rowContext['DT_REFERENCIA'].toISOString().split('T')[0];
    }

    if(hasData) result.push(rowObj);
  }
  return { sucesso: true, dados: result };
}

// ============================================================================
// ★ HELPER: SANITIZADOR DE JS (CORRIGE 'AND', 'OR', ETC)
// ============================================================================
function sanitizeJs_(script) {
  if (!script) return "true";
  let s = script;
  // Corrige ' and ' para ' && ' (case insensitive)
  s = s.replace(/\s+and\s+/gi, " && ");
  // Corrige ' or ' para ' || '
  s = s.replace(/\s+or\s+/gi, " || ");
  // Corrige '=' único em comparações (mas protege >=, <=, ==, !=)
  // Estratégia: Substitui ' = ' solto por ' == '
  s = s.replace(/([^!<>=])\s*=\s*([^=])/g, "$1 == $2");
  
  return s;
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
  } catch(e) { return { sucesso: false, erro: e.message };
  }
}

// ============================================================================
// HELPERS GERAIS
// ============================================================================

function cleanJson_(text) {
  try {
    let clean = text.replace(/```json/gi, "").replace(/```/g, "").trim();
    const start = clean.indexOf('{');
    const end = clean.lastIndexOf('}');
    if (start === -1 || end === -1) throw new Error("A IA não retornou um objeto JSON válido.");
    clean = clean.substring(start, end + 1);
    return JSON.parse(clean);
  } catch (e) {
    throw new Error("Erro de Sintaxe no JSON da IA: " + e.message);
  }
}
function apiGetColumns(s) { const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(s); return sh?sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>String(h).toUpperCase().trim()):[]; }
function apiGetMedicamentos(s,c) { return apiGetUniqueValues(s,c); }
function apiGetUniqueValues(s,c) { const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(s); if(!sh)return[]; const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v).toUpperCase().trim());
  const idx=h.indexOf(String(c).toUpperCase().trim()); if(idx<0)return[]; const v=sh.getRange(2,idx+1,sh.getLastRow()-1,1).getValues().flat(); return [...new Set(v.filter(String))].sort(); }

function parseSmartDate_(d) {
  if (!d) return null;
  if (d instanceof Date) return d;
  const str = String(d).trim(); if (str === "") return null;
  let p = str.split(/[\/\-\.\sT]/).filter(i => i !== "");
  if (p.length >= 3) {
    const v1=parseInt(p[0]), v2=parseInt(p[1])-1, v3=parseInt(p[2]);
    if (v1 > 31) { const dt=new Date(v1, v2, v3); if(!isNaN(dt.getTime())) return dt; } // ISO
    if (v3 > 31) { const dt=new Date(v3, v2, v1); if(!isNaN(dt.getTime())) return dt; } // BR
  }
  const nDt = new Date(str); return isNaN(nDt.getTime()) ? null : nDt;
}
function parseDatePTBR_(d) { return parseSmartDate_(d); }

function cleanHospitalValue_(val) {
  if (typeof val === 'number') return val;
  if (!val) return null;
  if (val instanceof Date) return val; 
  let str = String(val).trim();
  str = str.replace(/\b\d{2}\/\d{2}\/\d{4}\b/g, "").trim();
  str = str.replace(/\b\d{4}-\d{2}-\d{2}\b/g, "").trim();
  str = str.replace(/\d{1,2}:\d{2}(?::\d{2})?/g, "").trim();
  str = str.replace(',', '.');
  const num = parseFloat(str);
  return isNaN(num) ? val : num;
}

function apiImportarReferencias() {
  try {
    const targetSS = SpreadsheetApp.getActiveSpreadsheet();
    const masterSS = SpreadsheetApp.openById(MASTER_DB_ID);
    let sourceSheet = masterSS.getSheetByName("Exames_Referencia") || masterSS.getSheets()[0];
    if (!sourceSheet) return { sucesso: false, erro: "Mestre vazio." };
    const data = sourceSheet.getDataRange().getValues();
    let targetSheet = targetSS.getSheetByName("Exames_Referencia");
    if (targetSheet) targetSheet.clear(); else { targetSheet = targetSS.insertSheet("Exames_Referencia"); targetSheet.hideSheet();
    }
    targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    return { sucesso: true, total: data.length - 1 };
  } catch (e) { return { sucesso: false, erro: e.message };
  }
}

// ============================================================================
// ★ MÓDULO 5: FARMACOVIGILÂNCIA
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
  const colMap = {}; headers.forEach((h, i) => colMap[h] = i);
  const idxPront = config.colProntFatos ? headers.indexOf(config.colProntFatos.toUpperCase()) : -1;
  const idxData = config.colDtIni ? headers.indexOf(config.colDtIni.toUpperCase()) : headers.findIndex(h => h.includes("DATA") || h.includes("DT_"));
  const idxMed = config.colMed ? headers.indexOf(config.colMed.toUpperCase()) : headers.findIndex(h => h.includes("MEDICAMENTO"));
  const idxDose = config.colDose24h ? headers.indexOf(config.colDose24h.toUpperCase()) : headers.findIndex(h => h.includes("DOSE"));
  const idxCreat = headers.findIndex(h => h.includes("CREAT") || h.includes("CKD"));
  const idxLeuco = headers.findIndex(h => h.includes("LEUCOCITOS") || h.includes("WBC"));
  const idxPlaq = headers.findIndex(h => h.includes("PLAQUETAS") || h.includes("PLAT"));
  const idxAST = headers.findIndex(h => h.includes("AST") || h.includes("TGO"));

  if (idxPront === -1) return { sucesso: false, erro: "Coluna de Prontuário não mapeada." };

  const timeline = { meds: [], exams: [] };
  const targetPID = String(pid).trim();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowPID = String(row[idxPront]).trim();
    if (rowPID === targetPID) {
      const dt = parseSmartDate_(row[idxData]);
      if (!dt) continue;
      const dateStr = dt.toISOString().split('T')[0]; 
      if (idxMed > -1) {
        const medName = String(row[idxMed]);
        const doseVal = cleanHospitalValue_(row[idxDose]);
        if (doseVal > 0) timeline.meds.push({ date: dateStr, name: medName, dose: doseVal });
      }
      const addExam = (idx, type) => {
        if (idx > -1) {
          const val = cleanHospitalValue_(row[idx]);
          if (val && typeof val === 'number') timeline.exams.push({ date: dateStr, type: type, value: val });
        }
      };
      addExam(idxCreat, 'Creatinina');
      addExam(idxLeuco, 'Leucócitos');
      addExam(idxPlaq, 'Plaquetas');
      addExam(idxAST, 'AST/TGO');
    }
  }
  timeline.meds.sort((a, b) => a.date.localeCompare(b.date));
  timeline.exams.sort((a, b) => a.date.localeCompare(b.date));
  return { sucesso: true, data: timeline };
}

// ============================================================================
// ★ MOTOR DE DESCOBERTA DE MODELOS (AUTO-UPDATE)
// ============================================================================

function getAvailableModels_() {
  const cached = PropertiesService.getScriptProperties().getProperty('EUM_MODEL_LIST');
  if (cached) return JSON.parse(cached);
  return forceModelUpdate();
}

function forceModelUpdate() {
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${GEMINI_API_KEY}`;
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) throw new Error("Falha ao listar modelos");
    const json = JSON.parse(resp.getContentText());
    const models = json.models
      .filter(m => m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent"))
      .map(m => m.name.replace("models/", ""))
      .sort((a, b) => {
        const getVer = (s) => parseFloat(s.match(/(\d+\.\d+)/)?.[0] || 0);
        const vA = getVer(a), vB = getVer(b);
        if (vA !== vB) return vB - vA;
        if (a.includes("flash") && !b.includes("flash")) return -1;
        if (a.includes("live") && !b.includes("live")) return -1;
        return 0;
      });
    PropertiesService.getScriptProperties().setProperty('EUM_MODEL_LIST', JSON.stringify(models));
    return models;
  } catch (e) {
    console.warn("Erro no Auto-Update, usando fallback:", e);
    return FALLBACK_MODELS;
  }
}

function callGeminiSmart_(payload) {
  const modelList = getAvailableModels_();
  let lastError = null;
  for (const model of modelList) {
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
      if (code === 200) return JSON.parse(resp.getContentText()).candidates[0].content.parts[0].text;
      if (code === 429 || code === 503) { console.warn(`⚠️ ${model} indisponível. Roteando...`); continue; }
      throw new Error(`Erro API (${code}): ${resp.getContentText()}`);
    } catch (e) {
      lastError = e.message;
      if (e.message.includes("Fatal")) throw e;
    }
  }
  throw new Error(`Todas as IAs falharam. Verifique sua chave API. Erro: ${lastError}`);
}
