/**
 * ============================================================================
 * EUM MANAGER 7.2 - FINAL STABLE & RESTORED AI CALLER
 * Inclui:
 * 1. Detecção Automática de Coluna de Nascimento e SEXO (Zero-Config).
 * 2. RESTAURADO: Motor de descoberta de modelos de IA para chamadas robustas.
 * 3. Oráculo e Motor de Dados mantidos da v7.1.
 * ============================================================================
 */

const GEMINI_API_KEY = "AIzaSyAvJSN4OUt0vcPr5u5ozmlLqzpZmzeQSF8"; 
const MASTER_DB_ID = "1zKqYVR9seTPy3eyX5CR2xuXZtqSwwvERQn_KX3GK5JM"; // Base HCPA
const FALLBACK_MODELS = ["gemini-1.5-flash", "gemini-1.5-pro"];

// --- UI & NAVEGAÇÃO ---
function onOpen() { SpreadsheetApp.getUi().createMenu('🚀 EUM App').addItem('Abrir Painel', 'abrirDashboard').addToUi(); }

function abrirDashboard() { 
  var html = HtmlService.createTemplateFromFile('LabHome')
      .evaluate()
      .setTitle('EUM Manager 7.2')
      .setWidth(1200).setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'EUM Manager');
}

function include(filename) { return HtmlService.createTemplateFromFile(filename).evaluate().getContent(); }

// --- CONFIGURAÇÃO & ESTADO ---
function apiGetInitialState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(function(s) { return s.getName(); }).filter(function(n) { return !['Exames_Referencia','Doses_Referencia','Config','Dashboard'].includes(n); });
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
// ★ ORÁCULO COM MEMÓRIA E INTELIGÊNCIA DE GRÁFICO (v7.2)
// ============================================================================

const HISTORICO_ORACULO_KEY = 'EUM_CONVERSATION_HISTORY';
function getConversationHistory_() { const cache = CacheService.getUserCache(); const historyJson = cache.get(HISTORICO_ORACULO_KEY); return historyJson ? JSON.parse(historyJson) : []; }
function saveConversationHistory_(history) { const cache = CacheService.getUserCache(); cache.put(HISTORICO_ORACULO_KEY, JSON.stringify(history), 3600); }
function clearOracleHistory() { const cache = CacheService.getUserCache(); cache.remove(HISTORICO_ORACULO_KEY); return "Histórico do Oráculo foi limpo."; }

function apiOracleInterpret(pergunta, sheetName) {
  try {
    let cols = apiGetColumns(sheetName);
    cols.push("_IDADE_ANOS", "_PID", "_SEXO"); // Adiciona as colunas virtuais para a IA saber que existem
    let history = getConversationHistory_();
    
    const prompt = [
      'You are a Natural Language to Plotly.js JSON Parser.',
      'Your task is to convert the user\\\'s request into a STRICT JSON object. NO MARKDOWN, NO EXTRA TEXT.',
      '---',
      'AVAILABLE DATA COLUMNS: ' + JSON.stringify(cols),
      'CONVERSATION HISTORY: ' + JSON.stringify(history.slice(-4)),
      '---',
      'RULE 1: If the user asks for a chart, respond with the "Chart JSON" format. Otherwise, use the "Response JSON" format.',
      '',
      'Chart JSON Format:',
      '{',
      '  "config": {',
      '    "xAxis": "COLUMN_NAME",',
      '    "yAxis": "COLUMN_NAME",',
      '    "colorBy": "COLUMN_NAME",',
      '    "chartType": "bar | pie | scatter | box | histogram",',
      '    "aggregation": "count | sum | avg",',
      '    "group": true | false,',
      '    "filter": "row[\\\'COLUMN\\\'] > 10"',
      '  },',
      '  "explanation": "A friendly explanation of the chart you created."',
      '}',
      '',
      'Response JSON Format:',
      '{',
      '  "response": "Your friendly text response here."',
      '}',
      '---',
      'CRITICAL CHARTING RULES:',
      'A. **DUAL AXIS REQUESTS**: If the user asks for "X por Y" (e.g., "idade por sexo"), this is a DUAL AXIS request. It is MANDATORY to use both axes.',
      '   - The axis with many distinct values (like the virtual \`_IDADE_ANOS\` column) MUST be the \`xAxis\`.',
      '   - The axis with FEW distinct values (like the standardized \`_SEXO\` column) MUST be the \`colorBy\`.',
      '   - **FAILURE CONDITION**: If the request is "idade por sexo", you MUST use \`_SEXO\` for the \`colorBy\` field. Returning \`null\` is a failure.',
      '',
      'B. **HISTOGRAMS**: For analyzing the distribution of a continuous numerical variable like \`_IDADE_ANOS\`, ALWAYS set \`chartType\` to \`"histogram"\`.',
      '',
      'C. **VIRTUAL COLUMNS**: For age and sex, ALWAYS prefer the virtual, standardized columns \`_IDADE_ANOS\` and \`_SEXO\`.',
      '---',
      'USER REQUEST: "' + pergunta + '"',
      '---',
      'NOW, GENERATE THE JSON RESPONSE. DOUBLE-CHECK YOUR WORK AGAINST THE CRITICAL RULES BEFORE RESPONDING.'
    ].join('\\n');
        
    const payload = {
      contents: [{ parts: [{ text: prompt }] }],
      "safetySettings": [
          { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE" },
          { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE" },
          { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE" },
          { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE" }
      ]
    };
    
    const aiResponseText = callGeminiSmart_(payload);
    const aiJson = cleanJson_(aiResponseText);

    history.push({ role: "user", parts: [{ text: pergunta }] });
    history.push({ role: "model", parts: [{ text: aiResponseText }] });
    saveConversationHistory_(history);

    if (aiJson.config) {
      return { sucesso: true, config: aiJson.config, explanation: aiJson.explanation };
    } else if (aiJson.response) {
      return { sucesso: true, response: aiJson.response };
    } else {
      return { sucesso: true, response: aiResponseText };
    }

  } catch (e) {
    return { sucesso: false, erro: "Oráculo: " + e.message };
  }
}

// ============================================================================
// ★ MÓDULO 3: ENGINE DE DADOS 7.2 (FINAL ZERO-CONFIG)
// ============================================================================
function apiGetRawExplorerData(sheetName, cols, filterScript, virtualFields) {
  virtualFields = virtualFields || [];
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { sucesso: false, erro: "Aba não encontrada." };
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(function(h) { return String(h).toUpperCase().trim(); });
  const colMap = {}; headers.forEach(function(h, i) { colMap[h] = i; });
  
  const config = JSON.parse(PropertiesService.getScriptProperties().getProperty('EUM_CONFIG_MASTER') || '{}');
  const idxPront = config.core?.colProntFatos ? colMap[config.core.colProntFatos.toUpperCase()] : -1;

  // --- ZERO-CONFIG V3: Auto-detect birth date and sex columns ---
  let idxNasc = -1, idxSexo = -1;
  const possibleNascCols = ['DT_NASCIMENTO', 'NASCIMENTO', 'DATA DE NASCIMENTO', 'DT NASC', 'DOB', 'DT_NASC', 'NASC'];
  const possibleSexoCols = ['SEXO', 'SEX', 'GENERO', 'GENDER'];

  for (var i = 0; i < headers.length; i++) {
    if (idxNasc === -1 && possibleNascCols.indexOf(headers[i]) !== -1) idxNasc = i;
    if (idxSexo === -1 && possibleSexoCols.indexOf(headers[i]) !== -1) idxSexo = i;
    if (idxNasc !== -1 && idxSexo !== -1) break; // Optimization
  }

  let idxDataRef = config.core?.colDtIni ? colMap[config.core.colDtIni.toUpperCase()] : headers.indexOf("DATA");

  let filterFunc = null;
  const virtualFuncs = [];
  try {
    if (virtualFields && Array.isArray(virtualFields)) {
      virtualFields.forEach(function(vf) { virtualFuncs.push({ name: vf.name, func: new Function('row', 'return ' + vf.logic) }); });
    }
    if (filterScript && filterScript !== "true") filterFunc = new Function('row', 'return ' + filterScript);
  } catch (e) { return { sucesso: false, erro: "Erro Lógica: " + e.message }; }

  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i], rowContext = {}, rowObj = {};
    
    headers.forEach(function(h, idx) { rowContext[h] = cleanHospitalValue_(row[idx]); });
    
    const pid = (idxPront > -1) ? String(row[idxPront]).trim() : "ID_"+i;
    rowContext['_PID'] = pid;

    const dtNasc = idxNasc > -1 ? parseSmartDate_(row[idxNasc]) : null;
    const dtRef = idxDataRef > -1 ? parseSmartDate_(row[idxDataRef]) : new Date();
    rowContext['_IDADE_ANOS'] = (dtNasc && dtRef) ? parseFloat(((dtRef-dtNasc)/(365.25*86400000)).toFixed(2)) : 0;
    rowContext['_SEXO'] = idxSexo > -1 ? String(row[idxSexo] || '-').toUpperCase().charAt(0) : '-';

    try { virtualFuncs.forEach(function(vf) { rowContext[vf.name] = vf.func(rowContext); }); } catch(e) { continue; }
    if (filterFunc) { try { if (!filterFunc(rowContext)) continue; } catch(e) { continue; } }

    cols.forEach(function(c) { rowObj[c] = rowContext[c] !== undefined ? rowContext[c] : null; });
    
    // Always include virtual columns in the final object for drill-down
    rowObj['_PID'] = rowContext['_PID'];
    rowObj['_IDADE_ANOS'] = rowContext['_IDADE_ANOS'];
    rowObj['_SEXO'] = rowContext['_SEXO'];

    result.push(rowObj);
  }

  // --- DIAGNÓSTICO AUTOMÁTICO ---
  if (idxNasc === -1 && result.length > 5) {
    return { 
      sucesso: false, 
      erro: 'DIAGNÓSTICO: Falha ao calcular a idade.\\n\\nCAUSA: Não foi possível encontrar a coluna de Data de Nascimento.\\n\\nSOLUÇÃO: Renomeie a coluna para \\\'NASCIMENTO\\\' ou \\\'DT_NASCIMENTO\\\'.'
    };
  }

  return { sucesso: true, dados: result };
}

// ============================================================================
// ★★★ MÓDULO DE AUDITORIA DE DOSE (v7.2 - Simplificado) ★★★
// ============================================================================
function apiRunDoseAudit(sheetName, auditConfig) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(sheetName);
    let doseRefSheet = ss.getSheetByName("Doses_Referencia");

    if (!dataSheet) return { sucesso: false, erro: "Aba de dados não encontrada." };
    
    if (!doseRefSheet) {
      doseRefSheet = ss.insertSheet("Doses_Referencia");
      doseRefSheet.appendRow(["MEDICAMENTO", "DOSE_MAXIMA_DIA"]);
      return { 
        sucesso: false, 
        erro: "Aba de referência 'Doses_Referencia' não foi encontrada e acabou de ser criada. Por favor, preencha os limites de dose e tente novamente." 
      };
    }

    const data = dataSheet.getDataRange().getValues();
    const dataHeaders = data[0].map(h => String(h).toUpperCase().trim());
    const doseRefs = doseRefSheet.getDataRange().getValues();
    
    if (doseRefs.length < 2) {
      return { sucesso: false, erro: "Aba 'Doses_Referencia' está vazia. Por favor, preencha os limites de dose." };
    }
    
    const doseRefHeaders = doseRefs[0].map(h => String(h).toUpperCase().trim());

    const idxPront = dataHeaders.indexOf("PRONTUARIO");
    const idxMed = dataHeaders.indexOf("MEDICAMENTO");
    const idxDose = dataHeaders.indexOf("DOSE_24HS");

    if (idxPront === -1 || idxMed === -1 || idxDose === -1) {
      return { sucesso: false, erro: "Colunas necessárias (PRONTUARIO, MEDICAMENTO, DOSE_24HS) não encontradas na aba de dados." };
    }

    const idxRefMed = doseRefHeaders.indexOf("MEDICAMENTO");
    const idxRefDoseMax = doseRefHeaders.indexOf("DOSE_MAXIMA_DIA");

    if (idxRefMed === -1 || idxRefDoseMax === -1) {
      return { sucesso: false, erro: "Colunas necessárias (MEDICAMENTO, DOSE_MAXIMA_DIA) não encontradas na aba 'Doses_Referencia'." };
    }

    const doseLimits = {};
    for (let i = 1; i < doseRefs.length; i++) {
      const medName = String(doseRefs[i][idxRefMed]).trim().toUpperCase();
      const maxDose = parseFloat(doseRefs[i][idxRefDoseMax]);
      if (medName && !isNaN(maxDose)) {
        doseLimits[medName] = maxDose;
      }
    }

    const problemas = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const medName = String(row[idxMed]).trim().toUpperCase();
      const dose = parseFloat(row[idxDose]);
      const paciente = row[idxPront];

      if (doseLimits[medName] && !isNaN(dose)) {
        if (dose > doseLimits[medName]) {
          problemas.push({
            paciente: paciente,
            medicamento: row[idxMed],
            dose: dose,
            motivo: `Dose ${dose} excede o máximo de ${doseLimits[medName]}`
          });
        }
      }
    }

    return { sucesso: true, problemas: problemas };

  } catch (e) {
    return { sucesso: false, erro: "Erro inesperado na auditoria de dose: " + e.message };
  }
}


// ============================================================================
// ★ HELPERS GERAIS
// ============================================================================
function cleanJson_(text) {
  try {
    let clean = text.replace(/\\\`\\\`\\\`json/gi, "").replace(/\\\`\\\`\\\`/g, "").trim();
    const start = clean.indexOf('{');
    const end = clean.lastIndexOf('}');
    if (start === -1 || end === -1) throw new Error("A IA não retornou um objeto JSON válido.");
    clean = clean.substring(start, end + 1);
    return JSON.parse(clean);
  } catch (e) {
    throw new Error("Erro de Sintaxe no JSON da IA: " + e.message);
  }
}
function apiGetColumns(s) { var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(s); return sh?sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(h){return String(h).toUpperCase().trim();}):[]; }
function parseSmartDate_(d) {
  if (!d) return null; if (d instanceof Date) return d;
  const str = String(d).trim(); if (str === "") return null;
  let p = str.split(/[\\/\\-\\.\\sT]/).filter(function(i){return i !== "";});
  if (p.length >= 3) {
    const v1=parseInt(p[0]), v2=parseInt(p[1])-1, v3=parseInt(p[2]);
    if (v1 > 31) { var dt=new Date(v1, v2, v3); if(!isNaN(dt.getTime())) return dt; } // ISO
    if (v3 > 31) { var dt=new Date(v3, v2, v1); if(!isNaN(dt.getTime())) return dt; } // BR
  }
  const nDt = new Date(str); return isNaN(nDt.getTime()) ? null : nDt;
}
function cleanHospitalValue_(val) {
  if (typeof val === 'number') return val;
  if (!val) return null;
  let str = String(val).trim();
  str = str.replace(/\\d{2,4}[\\/-]\\d{1,2}[\\/-]\\d{2,4}/g, "").trim();
  str = str.replace(/\\d{1,2}:\\d{2}(?::\\d{2})?/g, "").trim();
  str = str.replace(',', '.');
  const num = parseFloat(str);
  return isNaN(num) ? val : num;
}

// ============================================================================
// ★ MODELOS DE GRÁFICO
// ============================================================================
function salvarModeloGrafico(nomeModelo, configuracao) {
  try {
    const db = SpreadsheetApp.openById(MASTER_DB_ID);
    let sheet = db.getSheetByName('GraficosModelos');
    if (!sheet) {
      sheet = db.insertSheet('GraficosModelos');
      sheet.appendRow(['NomeModelo', 'TipoGrafico', 'Configuracao', 'CriadoEm']);
    }
    const configObj = JSON.parse(configuracao);
    sheet.appendRow([nomeModelo, configObj.type, configuracao, new Date()]);
    return { sucesso: true };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

function getChartModels() {
  try {
    const db = SpreadsheetApp.openById(MASTER_DB_ID);
    const sheet = db.getSheetByName('GraficosModelos');
    if (!sheet) {
      return []; // Retorna um array vazio se a aba não existe
    }
    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Remove o cabeçalho
    const nomeModeloIndex = headers.indexOf('NomeModelo');
    const configuracaoIndex = headers.indexOf('Configuracao');

    return data.map(row => ({
      nomeModelo: row[nomeModeloIndex],
      configuracao: row[configuracaoIndex]
    }));
  } catch (e) {
    // Log the error for debugging, but return an empty array to the client
    console.error("Erro ao carregar modelos de gráfico: " + e.message);
    return [];
  }
}

// ============================================================================
// ★ MOTOR DE DESCOBERTA E CHAMADA DE IA (RESTAURADO)
// ============================================================================
function getAvailableModels_() {
  const cached = PropertiesService.getScriptProperties().getProperty('EUM_MODEL_LIST');
  if (cached) return JSON.parse(cached);
  return forceModelUpdate();
}

function forceModelUpdate() {
  try {
    const url = "https://generativelanguage.googleapis.com/v1beta/models?key=" + GEMINI_API_KEY;
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    
    if (resp.getResponseCode() !== 200) throw new Error("Falha ao listar modelos");

    const json = JSON.parse(resp.getContentText());
    
    const models = json.models
      .filter(function(m) { return m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent"); })
      .map(function(m) { return m.name.replace("models/", ""); })
      .sort(function(a, b) {
        const getVer = function(s) { return parseFloat(s.match(/(\\d+\\.\\d+)/)?.[0] || 0); };
        const vA = getVer(a), vB = getVer(b);
        if (vA !== vB) return vB - vA;
        if (a.includes("flash") && !b.includes("flash")) return -1;
        if (a.includes("live") && !b.includes("live")) return -1;
        return 0;
      });

    PropertiesService.getScriptProperties().setProperty('EUM_MODEL_LIST', JSON.stringify(models));
    console.log("Modelos Atualizados e Rankeados:", models);
    return models;

  } catch (e) {
    console.warn("Erro no Auto-Update, usando fallback:", e);
    return FALLBACK_MODELS;
  }
}

function callGeminiSmart_(payload) {
  const modelList = getAvailableModels_();
  let lastError = "Nenhum modelo de IA disponível ou todos falharam sem um erro específico.";

  for (var i = 0; i < modelList.length; i++) {
    var model = modelList[i];
    if (model.includes("1.0")) continue;

    const url = "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent?key=" + GEMINI_API_KEY;
    
    try {
      console.log("Tentando IA: " + model + "...");
      const resp = UrlFetchApp.fetch(url, { 
        method: 'post', 
        contentType: 'application/json', 
        payload: JSON.stringify(payload), 
        muteHttpExceptions: true 
      });
      
      const code = resp.getResponseCode();
      const responseText = resp.getContentText();

      if (code === 200) {
        const responseJson = JSON.parse(responseText);
        
        if (responseJson.candidates && responseJson.candidates[0].content && responseJson.candidates[0].content.parts &&
            responseJson.candidates[0].content.parts.length > 0 &&
            responseJson.candidates[0].content.parts[0].text) {
          console.log("Sucesso com o modelo: " + model);
          return responseJson.candidates[0].content.parts[0].text;
        } 
        
        lastError = "Resposta da IA bloqueada ou malformada para o modelo '" + model + "'. Detalhes: " + responseText;
        console.warn("⚠️ " + lastError);
        continue;
      }
      
      if (code === 429 || code === 503) {
        lastError = "Modelo '" + model + "' indisponível (código " + code + ").";
        console.warn("⚠️ " + lastError + " Roteando para o próximo...");
        continue;
      }
      
      lastError = "Erro na API com o modelo '" + model + "' (código " + code + "): " + responseText;
      console.warn("⚠️ " + lastError);
      continue;

    } catch (e) {
      lastError = "Exceção ao chamar o modelo '" + model + "': " + e.message;
      console.error("⛔ " + lastError);
      if (e.message.includes("Fatal")) throw e;
      continue;
    }
  }
  
  throw new Error("Falha em todas as IAs. Verifique sua chave API e as cotas. Último erro registrado: " + lastError);
}

// ============================================================================
// ★★★ MÓDULO DE FARMACOVIGILÂNCIA (Modernizado na v8.0) ★★★
// ============================================================================

/**
 * Busca todos os dados relevantes de um único paciente para construir uma linha do tempo.
 * @param {string} pid O Prontuário (ID) do paciente a ser buscado.
 * @returns {object} Um objeto com { sucesso: boolean, data: object, patientInfo: object, erro: string }.
 */
function apiGetPatientTimeline(pid) {
  try {
    const state = apiGetInitialState();
    if (!state.status.hasCore) {
      return { sucesso: false, erro: "Aba de dados principal (Fatos) não configurada." };
    }

    const config = state.config.core;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(config.abaFatos);
    
    if (!sheet) {
      return { sucesso: false, erro: `Aba de dados '${config.abaFatos}' não foi encontrada.` };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).toUpperCase().trim());
    const colMap = {};
    headers.forEach((h, i) => { colMap[h] = i; });

    // --- MAPEAMENTO DE COLUNAS (Essencial) ---
    const idxPront = colMap['PRONTUARIO'] !== undefined ? colMap['PRONTUARIO'] : colMap['PID'];
    const idxData = colMap['DT_REFERENCIA'] !== undefined ? colMap['DT_REFERENCIA'] : colMap['DATA'];
    const idxMed = colMap['MEDICAMENTO'];
    const idxDose = colMap['DOSE_24HS'] !== undefined ? colMap['DOSE_24HS'] : colMap['DOSE'];
    
    // --- MAPEAMENTO DE COLUNAS (Demográfico - Zero Config) ---
    let idxNasc = -1, idxSexo = -1;
    const possibleNascCols = ['DT_NASCIMENTO', 'NASCIMENTO', 'DATA DE NASCIMENTO', 'DT NASC', 'DOB', 'DT_NASC', 'NASC'];
    const possibleSexoCols = ['SEXO', 'SEX', 'GENERO', 'GENDER'];
    for (let i = 0; i < headers.length; i++) {
        if (idxNasc === -1 && possibleNascCols.includes(headers[i])) idxNasc = i;
        if (idxSexo === -1 && possibleSexoCols.includes(headers[i])) idxSexo = i;
        if (idxNasc !== -1 && idxSexo !== -1) break;
    }

    // Mapeamento de colunas de exames
    const examCols = {
      'Creatinina': colMap['CREAT_EQUACAO_CKD_EPI'] !== undefined ? colMap['CREAT_EQUACAO_CKD_EPI'] : colMap['CREATININA'],
      'Leucócitos': colMap['LEUCOCITOS'],
      'Plaquetas': colMap['PLAQUETAS'],
      'AST/TGO': colMap['AST']
    };

    if (idxPront === undefined) return { sucesso: false, erro: "Coluna de Prontuário ('PRONTUARIO' ou 'PID') não encontrada." };
    if (idxData === undefined) return { sucesso: false, erro: "Coluna de Data ('DT_REFERENCIA' ou 'DATA') não encontrada." };

    const timeline = { meds: [], exams: [] };
    let patientInfo = null; // Objeto para o "Dossiê"
    const targetPID = String(pid).trim();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowPID = String(row[idxPront]).trim();

      if (rowPID === targetPID) {
        const dt = parseSmartDate_(row[idxData]);
        if (!dt) continue;
        
        const dateStr = dt.toISOString().split('T')[0]; // Formato YYYY-MM-DD

        // --- NOVO: Captura dados demográficos na primeira ocorrência ---
        if (!patientInfo) {
            const dtNasc = idxNasc > -1 ? parseSmartDate_(row[idxNasc]) : null;
            const idade = (dtNasc && dt) ? parseFloat(((dt - dtNasc) / (365.25 * 86400000)).toFixed(1)) : null;
            const sexo = idxSexo > -1 ? String(row[idxSexo] || '-').toUpperCase().charAt(0) : '-';
            patientInfo = {
                id: targetPID,
                age: idade,
                sex: sexo,
                firstEventDate: dateStr,
                lastEventDate: dateStr
            };
        } else {
            // Atualiza a última data do evento
            if (dateStr > patientInfo.lastEventDate) {
                patientInfo.lastEventDate = dateStr;
            }
        }
        
        // Adiciona medicamento se existir dose válida
        if (idxMed !== undefined && idxDose !== undefined) {
          const medName = String(row[idxMed]);
          const doseVal = cleanHospitalValue_(row[idxDose]);
          if (doseVal && parseFloat(doseVal) > 0) {
            timeline.meds.push({ date: dateStr, name: medName, dose: doseVal });
          }
        }

        // Adiciona exames se existirem valores válidos
        Object.keys(examCols).forEach(examName => {
          const examIdx = examCols[examName];
          if (examIdx !== undefined) {
            const val = cleanHospitalValue_(row[examIdx]);
            if (val !== null && typeof val === 'number') {
              timeline.exams.push({ date: dateStr, type: examName, value: val });
            }
          }
        });
      }
    }

    if (!patientInfo) {
      return { sucesso: false, erro: `Paciente com Prontuário '${pid}' não encontrado na base de dados.` };
    }

    // Ordena os resultados por data
    timeline.meds.sort((a, b) => a.date.localeCompare(b.date));
    timeline.exams.sort((a, b) => a.date.localeCompare(b.date));

    return { sucesso: true, data: timeline, patientInfo: patientInfo };
  } catch (e) {
    return { sucesso: false, erro: "Erro inesperado no servidor: " + e.message };
  }
}

// ============================================================================
// ★★★ MÓDULO DE EXPORTAÇÃO PARA GOOGLE SLIDES (v1.0)
// Adicionar ao final do arquivo code.gs
// ============================================================================

/**
 * Exporta um item do relatório para uma apresentação Google Slides.
 * Mantém uma apresentação master persistente por instância (armazenada em ScriptProperties).
 *
 * @param {object} itemData - Objeto com as propriedades do item a exportar:
 *   - title {string}        : Título do card
 *   - type  {string}        : 'chart' | 'table'
 *   - timestamp {string}    : Data/hora de criação
 *   - userComment {string}  : Parecer técnico do utilizador (opcional)
 *   - imageBase64 {string}  : PNG em Base64 (obrigatório se type === 'chart')
 *   - tableData {Array}     : Array de objetos (obrigatório se type === 'table')
 *   - source {string}       : Módulo de origem (ex: "Explorador")
 * @returns {{ sucesso: boolean, url: string, slideNumber: number, erro: string }}
 */
function apiExportItemToSlide(itemData) {
  try {
    const props = PropertiesService.getScriptProperties();
    let presentation;
    const masterPresId = props.getProperty('EUM_SLIDES_MASTER_ID');

    if (masterPresId) {
      try {
        presentation = SlidesApp.openById(masterPresId);
      } catch (e) {
        // Apresentação foi apagada ou acesso revogado → criar nova
        presentation = createEumPresentation_();
        props.setProperty('EUM_SLIDES_MASTER_ID', presentation.getId());
      }
    } else {
      presentation = createEumPresentation_();
      props.setProperty('EUM_SLIDES_MASTER_ID', presentation.getId());
    }

    // Adiciona novo slide em branco
    const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    const slideNumber = presentation.getSlides().length;

    // Monta o conteúdo profissional do slide
    buildProfessionalSlide_(slide, itemData, slideNumber);

    // URL da apresentação (navegador abrirá no slide correto via hash)
    const presUrl = 'https://docs.google.com/presentation/d/'
      + presentation.getId()
      + '/edit#slide=p' + slideNumber;

    return {
      sucesso: true,
      url: presUrl,
      presentationId: presentation.getId(),
      slideNumber: slideNumber
    };

  } catch (e) {
    return { sucesso: false, erro: 'Erro ao exportar para Slide: ' + e.message };
  }
}

/**
 * Cria uma nova apresentação EUM com slide de capa.
 * @returns {GoogleAppsScript.Slides.Presentation}
 */
function createEumPresentation_() {
  const title = 'EUM Manager — Relatório de Evidências ('
    + new Date().toLocaleDateString('pt-BR') + ')';
  const presentation = SlidesApp.create(title);

  // Usa o primeiro slide padrão como capa
  const coverSlide = presentation.getSlides()[0];
  buildCoverSlide_(coverSlide);

  return presentation;
}

/**
 * Constrói o slide de capa da apresentação.
 * @param {GoogleAppsScript.Slides.Slide} slide
 */
function buildCoverSlide_(slide) {
  const W = 9144000; // 10 polegadas em EMU
  const H = 6858000; // 7.5 polegadas em EMU

  // Fundo escuro
  const bg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, W, H);
  bg.getFill().setSolidFill('#0f172a');
  bg.getBorder().setTransparent();

  // Faixa de acento
  const accent = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, H * 0.55, W * 0.15, H * 0.3);
  accent.getFill().setSolidFill('#2563eb');
  accent.getBorder().setTransparent();

  // Título principal
  const titleBox = slide.insertTextBox('EUM Manager', Math.round(W * 0.1), Math.round(H * 0.28), Math.round(W * 0.8), Math.round(H * 0.2));
  titleBox.getText().getTextStyle()
    .setFontSize(52).setForegroundColor('#FFFFFF').setBold(true).setFontFamily('Arial');
  titleBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  titleBox.getFill().setTransparent();
  titleBox.getBorder().setTransparent();

  // Subtítulo
  const subBox = slide.insertTextBox('Relatório de Evidências Clínicas', Math.round(W * 0.1), Math.round(H * 0.5), Math.round(W * 0.8), Math.round(H * 0.1));
  subBox.getText().getTextStyle().setFontSize(22).setForegroundColor('#94a3b8').setFontFamily('Arial');
  subBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  subBox.getFill().setTransparent();
  subBox.getBorder().setTransparent();

  // Data de geração
  const dateStr = new Date().toLocaleDateString('pt-BR', { year: 'numeric', month: 'long', day: 'numeric' });
  const dateBox = slide.insertTextBox(dateStr, Math.round(W * 0.1), Math.round(H * 0.65), Math.round(W * 0.8), Math.round(H * 0.08));
  dateBox.getText().getTextStyle().setFontSize(14).setForegroundColor('#475569').setItalic(true).setFontFamily('Arial');
  dateBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  dateBox.getFill().setTransparent();
  dateBox.getBorder().setTransparent();
}

/**
 * Constrói o conteúdo completo de um slide de evidência.
 * @param {GoogleAppsScript.Slides.Slide} slide
 * @param {object} item - Dados do item
 * @param {number} slideNum - Número sequencial do slide
 */
function buildProfessionalSlide_(slide, item, slideNum) {
  // === DIMENSÕES (EMU) ===
  const W        = 9144000;
  const H        = 6858000;
  const MARGIN   = 457200;   // 0.5"
  const HEADER_H = 1188720;  // ~1.3"
  const FOOTER_H = 480060;   // ~0.525"

  const CONTENT_TOP = HEADER_H + Math.round(MARGIN * 0.4);
  const CONTENT_H   = H - HEADER_H - FOOTER_H - Math.round(MARGIN * 0.8);
  const CONTENT_W   = W - MARGIN * 2;

  // === 1. CABEÇALHO ===
  const headerBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, W, HEADER_H);
  headerBg.getFill().setSolidFill('#1e293b');
  headerBg.getBorder().setTransparent();

  // Faixa azul de acento na base do cabeçalho
  const accentLine = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, HEADER_H - 91440, W, 91440);
  accentLine.getFill().setSolidFill('#2563eb');
  accentLine.getBorder().setTransparent();

  // Título
  const titleBox = slide.insertTextBox(
    item.title || 'Evidência Clínica',
    MARGIN, Math.round(MARGIN * 0.5), W - MARGIN * 2, HEADER_H - MARGIN
  );
  titleBox.getText().getTextStyle()
    .setFontSize(24).setForegroundColor('#FFFFFF').setBold(true).setFontFamily('Arial');
  titleBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  titleBox.getFill().setTransparent();
  titleBox.getBorder().setTransparent();

  // Badge de fonte/tipo (canto superior direito)
  const sourceBadge = slide.insertTextBox(
    (item.source || 'EUM') + ' · #' + slideNum,
    W - MARGIN - 1097280, Math.round(MARGIN * 0.5), 1097280, Math.round(MARGIN * 0.8)
  );
  sourceBadge.getText().getTextStyle().setFontSize(10).setForegroundColor('#94a3b8').setFontFamily('Arial');
  const sPara = sourceBadge.getText().getParagraphs()[0].getRange().getParagraphStyle();
  sPara.setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
  sourceBadge.getFill().setTransparent();
  sourceBadge.getBorder().setTransparent();

  // === 2. ÁREA DE CONTEÚDO ===
  if (item.type === 'chart' && item.imageBase64) {
    insertChartImageInSlide_(slide, item.imageBase64, MARGIN, CONTENT_TOP, CONTENT_W, CONTENT_H);
  } else if (item.type === 'table' && item.tableData && item.tableData.length > 0) {
    insertTableInSlide_(slide, item.tableData, MARGIN, CONTENT_TOP, CONTENT_W, CONTENT_H);
  } else {
    // Fallback textual
    const fallbackBox = slide.insertTextBox(
      '⚠️ Conteúdo não disponível para este tipo de item.',
      MARGIN, CONTENT_TOP + Math.round(CONTENT_H * 0.3), CONTENT_W, Math.round(CONTENT_H * 0.4)
    );
    fallbackBox.getText().getTextStyle().setFontSize(16).setForegroundColor('#94a3b8').setFontFamily('Arial');
    fallbackBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    fallbackBox.getFill().setSolidFill('#f8fafc');
  }

  // === 3. RODAPÉ ===
  const footerBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, H - FOOTER_H, W, FOOTER_H);
  footerBg.getFill().setSolidFill('#f1f5f9');
  footerBg.getBorder().setTransparent();

  const footerLeft = slide.insertTextBox(
    '🔬 EUM Manager — Plataforma de Vigilância Clínica',
    MARGIN, H - FOOTER_H, Math.round(W * 0.6), FOOTER_H
  );
  footerLeft.getText().getTextStyle()
    .setFontSize(9).setForegroundColor('#64748b').setItalic(true).setFontFamily('Arial');
  footerLeft.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  footerLeft.getFill().setTransparent();
  footerLeft.getBorder().setTransparent();

  const footerRight = slide.insertTextBox(
    item.timestamp || new Date().toLocaleString('pt-BR'),
    Math.round(W * 0.6), H - FOOTER_H, Math.round(W * 0.4) - MARGIN, FOOTER_H
  );
  footerRight.getText().getTextStyle()
    .setFontSize(9).setForegroundColor('#64748b').setFontFamily('Arial');
  footerRight.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  const rPara = footerRight.getText().getParagraphs()[0].getRange().getParagraphStyle();
  rPara.setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
  footerRight.getFill().setTransparent();
  footerRight.getBorder().setTransparent();

  // === 4. NOTAS DO APRESENTADOR (Parecer Técnico) ===
  if (item.userComment && item.userComment.trim()) {
    const notes = slide.getNotesPage().getSpeakerNotesShape();
    notes.getText().setText('📝 PARECER TÉCNICO DO FARMACÊUTICO/MÉDICO:\n\n' + item.userComment.trim());
    notes.getText().getTextStyle().setFontSize(12).setFontFamily('Arial');
  }
}

/**
 * Insere imagem de gráfico Base64 no slide, centralizada e com proporção 16:9.
 */
function insertChartImageInSlide_(slide, imageBase64, left, top, width, height) {
  try {
    const b64 = imageBase64.replace(/^data:image\/[a-z]+;base64,/, '');
    const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/png', 'eum_chart.png');
    const img = slide.insertImage(blob);

    // Calcula dimensões mantendo proporção 16:9
    const targetW = width;
    const targetH = Math.round(width * 9 / 16);
    const finalH = Math.min(targetH, height);
    const finalW = Math.round(finalH * 16 / 9);

    // Centraliza dentro da área de conteúdo
    const imgLeft = left + Math.round((width - finalW) / 2);
    const imgTop  = top  + Math.round((height - finalH) / 2);

    img.setLeft(imgLeft).setTop(imgTop).setWidth(finalW).setHeight(finalH);

  } catch (e) {
    // Se inserção de imagem falhar, adiciona mensagem de erro
    const errBox = slide.insertTextBox(
      '⚠️ Falha ao renderizar imagem: ' + e.message, left, top, width, Math.round(height * 0.3)
    );
    errBox.getText().getTextStyle().setFontSize(12).setForegroundColor('#ef4444').setFontFamily('Arial');
    errBox.getFill().setSolidFill('#fef2f2');
  }
}

/**
 * Insere dados tabulares como uma tabela formatada no slide.
 * Limita a 15 linhas de dados e 8 colunas para legibilidade.
 */
function insertTableInSlide_(slide, tableData, left, top, width, height) {
  const headers    = Object.keys(tableData[0]).slice(0, 8);       // Máx. 8 colunas
  const numDataRows = Math.min(tableData.length, 15);             // Máx. 15 linhas
  const numRows    = numDataRows + 1;                             // +1 para cabeçalho
  const numCols    = headers.length;

  // Calcula altura da tabela para caber na área de conteúdo
  const maxRowH = 457200; // 0.5" por linha máx.
  const rowH    = Math.min(Math.floor(height / numRows), maxRowH);
  const tableH  = rowH * numRows;
  const tableTop = top + Math.floor((height - tableH) / 2);

  const table = slide.insertTable(numRows, numCols, left, tableTop, width, tableH);

  // === Linha de Cabeçalho ===
  for (let c = 0; c < numCols; c++) {
    const cell = table.getCell(0, c);
    cell.getText().setText(String(headers[c]).substring(0, 22));
    cell.getText().getTextStyle()
      .setBold(true).setFontSize(10).setForegroundColor('#FFFFFF').setFontFamily('Arial');
    cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    cell.getFill().setSolidFill('#1e293b');
  }

  // === Linhas de Dados ===
  for (let r = 0; r < numDataRows; r++) {
    const rowData = tableData[r];
    const isEven  = r % 2 === 0;
    for (let c = 0; c < numCols; c++) {
      const cell = table.getCell(r + 1, c);
      const val  = (rowData[headers[c]] !== undefined && rowData[headers[c]] !== null)
        ? String(rowData[headers[c]]).substring(0, 30)
        : '—';
      cell.getText().setText(val);
      cell.getText().getTextStyle()
        .setFontSize(9).setForegroundColor('#334155').setFontFamily('Arial');
      cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      cell.getFill().setSolidFill(isEven ? '#f8fafc' : '#FFFFFF');
    }
  }

  // Nota de truncamento se necessário
  if (tableData.length > 15 || Object.keys(tableData[0]).length > 8) {
    const noteText = [];
    if (tableData.length > 15) noteText.push('Exibindo 15 de ' + tableData.length + ' registros');
    if (Object.keys(tableData[0]).length > 8) noteText.push('colunas reduzidas para 8');
    const noteBox = slide.insertTextBox(
      'ℹ️ ' + noteText.join(' | ') + '. Consulte o relatório HTML para dados completos.',
      left, tableTop + tableH + 91440, width, 274638
    );
    noteBox.getText().getTextStyle()
      .setFontSize(9).setForegroundColor('#94a3b8').setItalic(true).setFontFamily('Arial');
    noteBox.getFill().setTransparent();
    noteBox.getBorder().setTransparent();
  }
}

/**
 * Limpa o ID da apresentação master armazenado, forçando criação de nova no próximo export.
 * Útil para iniciar um novo relatório de apresentação.
 */
function apiResetSlidesPresentation() {
  PropertiesService.getScriptProperties().deleteProperty('EUM_SLIDES_MASTER_ID');
  return { sucesso: true, msg: 'Apresentação master resetada. Próximo export criará uma nova.' };
}

/**
 * Retorna o URL da apresentação master atual (se existir).
 */
function apiGetSlidesUrl() {
  const presId = PropertiesService.getScriptProperties().getProperty('EUM_SLIDES_MASTER_ID');
  if (!presId) return { sucesso: false, erro: 'Nenhuma apresentação exportada ainda.' };
  return {
    sucesso: true,
    url: 'https://docs.google.com/presentation/d/' + presId + '/edit',
    presentationId: presId
  };
}

// ============================================================
// ADICIONAR AO code.gs — função utilitária apiGetUniqueValues
// Usada por LabEficacia, LabRegrasDose (fallback) e outros módulos.
// Retorna array de strings únicas e não-vazias de uma coluna.
// ============================================================

/**
 * Retorna os valores únicos de uma coluna específica de uma aba.
 * Ignora a linha de cabeçalho, células vazias e valores idênticos ao header.
 *
 * @param {string} sheetName  - Nome da aba na planilha ativa.
 * @param {string} colName    - Nome da coluna (case-insensitive).
 * @returns {string[]}        - Array ordenado de valores únicos.
 */
function apiGetUniqueValues(sheetName, colName) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      console.warn('apiGetUniqueValues: aba não encontrada — ' + sheetName);
      return [];
    }

    const data    = sheet.getDataRange().getValues();
    if (data.length < 2) return []; // só cabeçalho ou vazio

    // Encontra índice da coluna (case-insensitive)
    const headers = data[0].map(function(h) { return String(h).toUpperCase().trim(); });
    const colIdx  = headers.indexOf(colName.toUpperCase().trim());

    if (colIdx === -1) {
      console.warn('apiGetUniqueValues: coluna não encontrada — ' + colName + ' em ' + sheetName);
      return [];
    }

    const seen = {};
    const result = [];

    for (var i = 1; i < data.length; i++) {
      var raw = data[i][colIdx];
      if (raw === null || raw === undefined) continue;

      var val = String(raw).trim();

      // Descarta vazios e valores iguais ao nome da coluna (header acidental)
      if (val === '' || val.toUpperCase() === colName.toUpperCase().trim()) continue;

      if (!seen[val]) {
        seen[val] = true;
        result.push(val);
      }
    }

    return result.sort(function(a, b) {
      return a.localeCompare(b, 'pt-BR', { sensitivity: 'base' });
    });

  } catch (e) {
    console.error('apiGetUniqueValues erro: ' + e.message);
    return [];
  }
}

// ============================================================
// ADICIONAR AO code.gs — função dedicada para auditoria de dose
// Não possui o diagnóstico de idxNasc que bloqueia apiGetRawExplorerData
// ============================================================

/**
 * Busca os dados brutos de uma aba para auditoria de dose.
 * Diferente de apiGetRawExplorerData, NÃO bloqueia execução
 * se a coluna de nascimento não for encontrada pelo auto-detect.
 * 
 * @param {object} cfg
 *   cfg.sheet    {string} - Nome da aba
 *   cfg.cols     {Array}  - Colunas a retornar
 *   cfg.colNasc  {string} - Coluna de nascimento (nome exato, opcional)
 *   cfg.colMed   {string} - Coluna de medicamento
 *   cfg.medSel   {string} - Valor do medicamento a filtrar (opcional, '' = todos)
 */
function apiGetAuditData(cfg) {
  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var sheet  = ss.getSheetByName(cfg.sheet);
    if (!sheet) return { sucesso: false, erro: 'Aba "' + cfg.sheet + '" não encontrada.' };

    var data    = sheet.getDataRange().getValues();
    if (data.length < 2) return { sucesso: true, dados: [] };

    var headers = data[0].map(function(h) { return String(h).toUpperCase().trim(); });
    var colMap  = {};
    headers.forEach(function(h, i) { colMap[h] = i; });

    // Colunas solicitadas
    var cols = (cfg.cols || []).map(function(c) { return String(c).toUpperCase().trim(); });

    // Índice da coluna de nascimento (usa o que o usuário escolheu, não auto-detect)
    var idxNasc = -1;
    if (cfg.colNasc && String(cfg.colNasc).trim() !== '') {
      idxNasc = colMap[String(cfg.colNasc).toUpperCase().trim()];
      if (idxNasc === undefined) idxNasc = -1;
    }
    
    // Fallback: tenta auto-detect se o usuário não especificou
    if (idxNasc === -1) {
      var nascPatterns = ['DT_NASCIMENTO','NASCIMENTO','DATA DE NASCIMENTO','DT NASC','DOB','DT_NASC','NASC','DATA_NASCIMENTO'];
      for (var pi = 0; pi < headers.length; pi++) {
        if (nascPatterns.indexOf(headers[pi]) !== -1) { idxNasc = pi; break; }
      }
    }

    // Filtro de medicamento (opcional)
    var medNorm = cfg.medSel ? String(cfg.medSel).trim().toUpperCase() : '';
    var idxMed  = cfg.colMed ? (colMap[String(cfg.colMed).toUpperCase().trim()] !== undefined
                                ? colMap[String(cfg.colMed).toUpperCase().trim()] : -1) : -1;

    var result = [];
    var today  = new Date();

    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      // Filtro de medicamento no servidor (mais eficiente para planilhas grandes)
      if (medNorm && idxMed > -1) {
        var rowMed = String(row[idxMed] || '').trim().toUpperCase();
        if (rowMed !== medNorm) continue;
      }

      var rowObj = {};

      // Preenche colunas solicitadas
      cols.forEach(function(c) {
        var idx = colMap[c];
        if (idx !== undefined && idx > -1) {
          var raw = row[idx];
          // Converte datas do Sheets para string ISO
          if (raw instanceof Date) {
            rowObj[c] = raw.toISOString().split('T')[0];
          } else if (typeof raw === 'number') {
            rowObj[c] = raw;
          } else {
            var str = String(raw || '').trim();
            var num = parseFloat(str.replace(',', '.'));
            rowObj[c] = (str !== '' && !isNaN(num)) ? num : (str || null);
          }
        } else {
          rowObj[c] = null;
        }
      });

      // Calcula _IDADE_ANOS se coluna de nascimento disponível
      if (idxNasc > -1 && row[idxNasc]) {
        var dobRaw = row[idxNasc];
        var dob    = (dobRaw instanceof Date) ? dobRaw : new Date(String(dobRaw));
        if (!isNaN(dob.getTime())) {
          rowObj['_IDADE_ANOS'] = parseFloat(((today - dob) / (365.25 * 86400000)).toFixed(2));
        } else {
          rowObj['_IDADE_ANOS'] = null;
        }
      } else {
        rowObj['_IDADE_ANOS'] = null;
      }

      result.push(rowObj);
    }

    return { sucesso: true, dados: result, total: result.length };

  } catch (e) {
    return { sucesso: false, erro: 'apiGetAuditData: ' + e.message };
  }
}

// ================================================================
// MÓDULO: AUDITORIA RENAL — code.gs
// Adicionar estas funções ao seu arquivo code.gs existente.
// ================================================================
 
/**
 * API principal do módulo de Auditoria Renal.
 * Cruza os valores de TFG com a dose diária administrada,
 * classificando cada registo por zona de segurança.
 *
 * @param {Object} params
 * @param {string} params.sheet    - Nome da aba do Google Sheets
 * @param {string} params.colTfg   - Nome da coluna de TFG / Creatinina
 * @param {string} params.colMed   - Nome da coluna de Medicamento
 * @param {string} params.colDose  - Nome da coluna de Dose Diária (mg)
 * @param {string} params.colPid   - Nome da coluna de Prontuário / ID Paciente
 * @param {string} params.colUnit  - Nome da coluna de Especialidade (opcional)
 * @param {Object} params.zones    - Definições das zonas de risco
 * @returns {{ sucesso: boolean, dados: Array, meta: Object }}
 */
function apiGetRenalAuditData(params) {
  try {
    // --- 1. VALIDAÇÃO DOS PARÂMETROS ---
    if (!params || !params.sheet) {
      return { sucesso: false, erro: 'Parâmetro "sheet" não informado.' };
    }
    if (!params.colTfg || !params.colMed || !params.colDose || !params.colPid) {
      return { sucesso: false, erro: 'Colunas obrigatórias (TFG, Medicamento, Dose, Prontuário) não mapeadas.' };
    }
 
    // --- 2. LEITURA DA PLANILHA ---
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(params.sheet);
    if (!sheet) {
      return { sucesso: false, erro: `Aba "${params.sheet}" não encontrada na planilha.` };
    }
 
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) {
      return { sucesso: false, erro: 'A aba está vazia ou não possui dados suficientes.' };
    }
 
    // Lê tudo de uma vez (mais performático que leitura por coluna)
    const rawData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = rawData[0].map(h => String(h).trim());
 
    // --- 3. MAPEAMENTO DE ÍNDICES DE COLUNA ---
    const findIdx = (colName) => {
      if (!colName) return -1;
      const idx = headers.findIndex(h => h.toUpperCase() === colName.trim().toUpperCase());
      return idx; // retorna -1 se não encontrar
    };
 
    const idxTfg  = findIdx(params.colTfg);
    const idxMed  = findIdx(params.colMed);
    const idxDose = findIdx(params.colDose);
    const idxPid  = findIdx(params.colPid);
    const idxUnit = params.colUnit ? findIdx(params.colUnit) : -1;
 
    // Tenta encontrar colunas de sexo automaticamente
    const idxSexo = headers.findIndex(h => /^(sexo|sex|gender)$/i.test(h.trim()));
 
    // Valida índices obrigatórios
    const missingCols = [];
    if (idxTfg  < 0) missingCols.push(params.colTfg);
    if (idxMed  < 0) missingCols.push(params.colMed);
    if (idxDose < 0) missingCols.push(params.colDose);
    if (idxPid  < 0) missingCols.push(params.colPid);
 
    if (missingCols.length > 0) {
      return {
        sucesso: false,
        erro: `Colunas não encontradas na aba: ${missingCols.join(', ')}.\nVerifique o mapeamento.`
      };
    }
 
    // --- 4. ZONAS DE RISCO (com fallback para padrão KDIGO) ---
    const zones = params.zones || {
      safe:     { tfgMin: 60,  tfgMax: 9999, doseMax: null, name: 'Seguro'  },
      alert:    { tfgMin: 30,  tfgMax: 59,   doseMax: null, name: 'Alerta'  },
      critical: { tfgMin: 0,   tfgMax: 29,   doseMax: null, name: 'Critico' }
    };
 
    // --- 5. PROCESSAMENTO LINHA A LINHA ---
    const dados    = [];
    const rowCount = rawData.length;
 
    for (let i = 1; i < rowCount; i++) {
      const row = rawData[i];
 
      // Ignora linhas completamente vazias
      if (row.every(cell => cell === '' || cell === null || cell === undefined)) continue;
 
      // Extração e sanitização dos valores principais
      const tfgRaw  = row[idxTfg];
      const doseRaw = row[idxDose];
      const med     = String(row[idxMed]  || '').trim();
      const pid     = String(row[idxPid]  || '').trim();
      const unit    = idxUnit >= 0 ? String(row[idxUnit] || '').trim() : '';
      const sexoRaw = idxSexo >= 0 ? String(row[idxSexo] || '').trim().toUpperCase() : '';
 
      // Salta registros sem prontuário ou medicamento
      if (!pid || !med) continue;
 
      // Conversão robusta de TFG e Dose para número
      const tfg  = parseLabValue(tfgRaw);
      const dose = parseLabValue(doseRaw);
 
      // TFG inválida (nula ou negativa) não entra na análise
      if (tfg === null || tfg < 0) continue;
 
      // Normaliza sexo
      const sexo = normalizeSexo(sexoRaw);
 
      // Classifica na zona correta
      const { status, zoneKey, exceded } = classifyByZone(tfg, dose, zones);
 
      dados.push({
        _pid:     pid,
        _tfg:     tfg,
        _dose:    dose !== null ? dose : 0,
        _med:     med,
        _unit:    unit || null,
        _sexo:    sexo,
        _status:  status,
        _zoneKey: zoneKey,
        _exceded: exceded
      });
    }
 
    // --- 6. META-DADOS PARA KPIs ---
    const meta = {
      total:    dados.length,
      seguros:  dados.filter(d => d._status === 'Seguro').length,
      alertas:  dados.filter(d => d._status === 'Alerta').length,
      criticos: dados.filter(d => d._status === 'Critico').length,
      excedidos: dados.filter(d => d._exceded).length,
      medicamentos: [...new Set(dados.map(d => d._med))].length
    };
 
    return { sucesso: true, dados: dados, meta: meta };
 
  } catch (e) {
    Logger.log('[apiGetRenalAuditData] ERRO: ' + e.message + '\nStack: ' + e.stack);
    return {
      sucesso: false,
      erro: 'Erro interno ao processar auditoria renal: ' + e.message
    };
  }
}
 
// ================================================================
// FUNÇÕES AUXILIARES PRIVADAS
// ================================================================
 
/**
 * Converte um valor bruto de laboratório em número decimal.
 * Trata strings com vírgula, unidades coladas (ex: "45,3 mL/min"),
 * valores do tipo Date (do Sheets) e valores nulos.
 *
 * @param {*} raw - Valor bruto da célula do Sheets
 * @returns {number|null} - Número decimal ou null se inválido
 */
function parseLabValue(raw) {
  if (raw === null || raw === undefined || raw === '') return null;
 
  // Objetos Date do Sheets (não esperado para lab, mas tratado)
  if (raw instanceof Date) return null;
 
  const str = String(raw)
    .trim()
    // Remove unidades comuns coladas ao número
    .replace(/\s*(ml\/min|mg|mg\/dl|umol\/l|g\/dl|%|mEq\/L|mmol\/L|UI\/L|U\/L|mg\/24h)/gi, '')
    // Trata notação europeia (vírgula como separador decimal)
    .replace(/\.(?=\d{3})/g, '')  // remove pontos de milhar (ex: 1.500 → 1500)
    .replace(',', '.')             // substitui vírgula decimal por ponto
    .trim();
 
  // Remove caracteres não numéricos restantes exceto ponto e sinal de menos
  const cleaned = str.replace(/[^0-9.\-]/g, '');
  if (cleaned === '' || cleaned === '-') return null;
 
  const num = parseFloat(cleaned);
  return isNaN(num) ? null : num;
}
 
/**
 * Normaliza o campo de sexo para 'M' ou 'F'.
 *
 * @param {string} raw - Valor bruto do campo sexo
 * @returns {string} - 'M', 'F' ou string original se não identificado
 */
function normalizeSexo(raw) {
  if (!raw) return '';
  const upper = raw.toUpperCase().trim();
 
  // Masculino
  if (['M', 'MASC', 'MASCULINO', 'MALE', 'H', 'HOMBRE', 'HOMME'].includes(upper)) return 'M';
  if (upper.startsWith('M') && upper.length <= 4) return 'M';
 
  // Feminino
  if (['F', 'FEM', 'FEMININO', 'FEMALE', 'MUJER', 'FEMME'].includes(upper)) return 'F';
  if (upper.startsWith('F') && upper.length <= 4) return 'F';
 
  return upper; // retorna como veio se não reconhecido
}
 
/**
 * Classifica um registo (TFG + dose) nas zonas de risco definidas.
 *
 * @param {number} tfg   - Valor de TFG em mL/min
 * @param {number} dose  - Dose diária em mg
 * @param {Object} zones - Objeto com as zonas {safe, alert, critical}
 * @returns {{ status: string, zoneKey: string, exceded: boolean }}
 */
function classifyByZone(tfg, dose, zones) {
  // Ordem de prioridade de avaliação: critical → alert → safe
  const priority = ['critical', 'alert', 'safe'];
 
  for (const key of priority) {
    const z = zones[key];
    if (!z) continue;
 
    const tfgMin = parseFloat(z.tfgMin) || 0;
    const tfgMax = parseFloat(z.tfgMax) || 9999;
 
    if (tfg >= tfgMin && tfg <= tfgMax) {
      const doseMax = z.doseMax ? parseFloat(z.doseMax) : null;
      const exceded = (doseMax !== null && dose !== null && dose > doseMax);
      return {
        status:  exceded ? escalateStatus(key) : (z.name || toTitleCase(key)),
        zoneKey: key,
        exceded: exceded
      };
    }
  }
 
  // Fora de qualquer zona mapeada → trata como "Seguro" sem excesso
  return { status: 'Seguro', zoneKey: 'safe', exceded: false };
}
 
/**
 * Escalona o status quando a dose é excedida dentro de uma zona.
 * Ex: um paciente em zona "Alerta" com dose excedida vira "Critico".
 */
function escalateStatus(zoneKey) {
  const escalation = { safe: 'Alerta', alert: 'Critico', critical: 'Critico' };
  return escalation[zoneKey] || 'Critico';
}
 
/** Capitaliza a primeira letra. */
function toTitleCase(str) {
  return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
}
