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
