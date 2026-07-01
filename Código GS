/**
 * ==========================================
 * 1. ROTEAMENTO E RENDERIZAÇÃO DA INTERFACE
 * ==========================================
 */

function doGet(e) {
  const scriptUrl = ScriptApp.getService().getUrl();
  var tmpDash = HtmlService.createTemplateFromFile('paginainicial');
  tmpDash.urlBase = scriptUrl;
  
  return tmpDash.evaluate()
      .setTitle('GeoFOGO - CBMMG')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Função auxiliar para incluir arquivos HTML de forma modular (Particionamento)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ==========================================
 * 2. CONFIGURAÇÃO DO BANCO DE DADOS (PLANILHA)
 * ==========================================
 */
function configurarSistema() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const abas = {
      "Cadastro_geral": ["ID", "NOME", "COB", "MUNICÍPIO", "CHAMADA", "STATUS", "INÍCIO", "TÉRMINO", "DURAÇÃO", "EFETIVO_TOTAL", "VTRS_TOTAL"],
      "Cadastro_Efetivo": ["ID_OP", "NOME_OP", "ORD", "STATUS", "INSTITUIÇÃO", "N_BM", "PG", "NOME", "QUALIFICAÇÃO PROFISSIONAL", "UNIDADE", "CONTATO", "FUNÇÃO", "LOCAL_ATUAÇÃO", "INÍCIO_EMPENHO", "TÉRMINO DO EMPENHO", "DIAS_EMPENHO", "DIA_CHEGADA", "LAT", "LONG"],
      "SCO": ["ID_OP", "NOME_OP", "CICLO_INICIO", "CICLO_FIM", "DIA", "RESUMO", "OBJETIVOS", "ENFASE_PERIODO", "SITUAÇÃO", "NOME", "CARGO", "DATA", "AÇÕES", "LAT" , "LONG", "MUNICIPIO"],
      "SCO_Acoes": ["ID_SCO_FK", "HORA", "AÇÃO", "FOTOS", "LAT", "LONG"],
      // Registro de vítimas (novo módulo da plataforma de resgate)
      "Vitimas": ["ID_OP", "NOME_OP", "DIA", "HORA", "NOME", "IDADE", "SEXO", "STATUS", "LOCAL", "DESCRICAO", "ENCAMINHAMENTO", "DATA_REGISTRO", "LAT", "LONG"],
      // Base de apoio para importação de efetivo:
      "EFETIVO_CBMMG": [],
      "VIATURA_CBMMG": []
    };

    for (let nome in abas) {
      let aba = ss.getSheetByName(nome);
      if (!aba) {
        aba = ss.insertSheet(nome);
        // Só aplica formatação de cabeçalho se houver array de colunas definido
        if (abas[nome].length > 0) {
          aba.getRange(1, 1, 1, abas[nome].length).setValues([abas[nome]])
             .setBackground("#0d6efd").setFontColor("white").setFontWeight("bold");
          aba.setFrozenRows(1);
        }
      }
    }
    return "Sistema configurado com sucesso! Abas criadas.";
  } catch (e) {
    return "Erro ao configurar sistema: " + e.message;
  }
}

/**
 * ==========================================
 * 3. NÚCLEO DE LEITURA E GERENCIAMENTO DE CACHE
 * ==========================================
 */

function getDadosAba(nomeAba) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(nomeAba);
  if (cached) { 
    try { return JSON.parse(cached); } catch(e) { cache.remove(nomeAba); } 
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(nomeAba);
  if (!aba) return [];

  const dadosBrutos = aba.getDataRange().getValues();
  const dadosFinais = [];
  let temData = false;

  if (dadosBrutos.length > 1) {
    for (let c = 0; c < dadosBrutos[1].length; c++) {
      if (dadosBrutos[1][c] instanceof Date) { 
        temData = true;
        break; 
      }
    }
  }

  for (let i = 0; i < dadosBrutos.length; i++) {
    const linha = dadosBrutos[i];
    if (linha[0] === "" || linha[0] === null) continue;

    if (temData) {
      const novaLinha = new Array(linha.length);
      for (let j = 0; j < linha.length; j++) {
        novaLinha[j] = (linha[j] instanceof Date) ? linha[j].toISOString() : linha[j];
      }
      dadosFinais.push(novaLinha);
    } else {
      dadosFinais.push(linha);
    }
  }

  try { cache.put(nomeAba, JSON.stringify(dadosFinais), 300); } catch(e) {}
  return dadosFinais;
}

function limparCache(nomeAba) {
  const cache = CacheService.getScriptCache();
  cache.remove(nomeAba);
  cache.remove("pagina_interna_all_keys");
}

function invalidarCacheOp(idOp) {
  if (!idOp) return;
  CacheService.getScriptCache().remove("pagina_interna_" + idOp);
}

/**
 * Compila em uma única chamada leve todos os dados necessários ao abrir a página interna da operação
 */
function carregarDadosPaginaInterna(idOp) {
  try {
    const cache = CacheService.getScriptCache();
    const chaveCache = "pagina_interna_" + idOp;
    const cached = cache.get(chaveCache);
    if (cached) {
      try { return JSON.parse(cached); } catch(e) { cache.remove(chaveCache); }
    }

    const resultado = {
      operacao: buscarDadosOperacaoPorId(idOp),
      efetivo: buscarEfetivo(idOp),
      vitimas: buscarVitimas(idOp)
    };

    try {
      cache.put(chaveCache, JSON.stringify(resultado), 60);
    } catch(e) {}

    return resultado;
  } catch (e) {
    throw new Error("Erro ao compilar dados da página interna: " + e.message);
  }
}

/**
 * ==========================================
 * 4. FUNÇÕES DE BUSCA E CRUDS (BACKEND)
 * ==========================================
 */

function buscarOperacoes() {
  try {
    const dados = getDadosAba('Cadastro_geral');
    if (!dados || dados.length <= 1) return [];

    const linhas = dados.slice(1);
    const agora = new Date();
    
    return lines.map(linha => {
      const dIni = linha[8] ? new Date(linha[8]) : null;
      const dTer = (linha[9] && String(linha[9]).trim() !== "") ? new Date(linha[9]) : null;
      let duracaoFinal = linha[10] || "---";

      if (dIni && !isNaN(dIni.getTime()) && !dTer) {
        const diff = agora - dIni;
        const d = Math.floor(diff / 86400000);
        const h = Math.floor((diff % 86400000) / 3600000);
        const m = Math.floor((diff % 3600000) / 60000);
        duracaoFinal = `${d} d ${h} h ${m} min`;
      }
      const formatar = (d) => (d && !isNaN(d.getTime())) ? Utilities.formatDate(d, "GMT-3", "dd/MM/yyyy HH:mm") : "---";

      return { 
        id: linha[0], 
        nome: linha[1] || "S/N", 
        cob: linha[2] || "---", 
        municipio: linha[3] || "---", 
        grupo: AppScript = linha[4] || "", 
        subgrupo: linha[5] || "", 
        status: linha[7] || "---", 
        inicio: formatar(dIni), 
        termino: formatar(dTer), 
        duracao: duracaoFinal 
      };
    }).reverse();
  } catch (e) { 
    throw new Error("Erro em buscarOperacoes: " + e.message); 
  }
}

function buscarDadosOperacaoPorId(idOp) {
  const dados = getDadosAba('Cadastro_geral');
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0] == idOp) {
      const formatarHTML = (d) => {
        if (!d) return "";
        let dataObj = new Date(d);
        return (!isNaN(dataObj.getTime())) ? Utilities.formatDate(dataObj, "GMT-3", "yyyy-MM-dd'T'HH:mm") : "";
      };
      let inicioNativoStr = "";
      if (dados[i][8]) {
        let tempDate = new Date(dados[i][8]);
        if (!isNaN(tempDate.getTime())) inicioNativoStr = tempDate.toISOString();
      }
      return { 
        id: dados[i][0], nome: dados[i][1], cob: dados[i][2], municipio: dados[i][3], 
        grupo: dados[i][4], subgrupo: dados[i][5] ? String(dados[i][5]) : "", 
        cobrade: dados[i][6], status: dados[i][7], inicio: formatarHTML(dados[i][8]), 
        termino: formatarHTML(dados[i][9]), duracao: dados[i][10], inicio_nativo: inicioNativoStr 
      };
    }
  }
  return null;
}

// --- FUNÇÕES AUXILIARES DE DATA ---
function formatarParaBR(valor) {
  if (!valor) return "";
  if (valor instanceof Date) { // Se o Sheets leu como data nativa
    const dia = String(valor.getDate()).padStart(2, '0');
    const mes = String(valor.getMonth() + 1).padStart(2, '0');
    const ano = valor.getFullYear();
    const hora = String(valor.getHours()).padStart(2, '0');
    const min = String(valor.getMinutes()).padStart(2, '0');
    return `${dia}/${mes}/${ano} ${hora}:${min}`;
  }
  let str = String(valor); // Se leu como texto
  if (str.includes('T')) {
    let d = str.split('T')[0].split('-');
    let t = str.split('T')[1].substring(0,5);
    if (d.length === 3) return `${d[2]}/${d[1]}/${d[0]} ${t}`;
  } else if (str.match(/^\d{4}-\d{2}-\d{2}/)) {
      let p = str.split(' ');
      let d = p[0].split('-');
      let t = p[1] ? p[1].substring(0,5) : '00:00';
      return `${d[2]}/${d[1]}/${d[0]} ${t}`;
  }
  return str;
}

function parseDataGS(valor) {
  if (!valor) return null;
  if (valor instanceof Date) return valor;
  let str = String(valor);
  if (str.includes('/')) { // Desmonta DD/MM/YYYY para calcular
    let parts = str.split(' ');
    let d = parts[0].split('/');
    let t = parts[1] ? parts[1].split(':') : ['00','00'];
    if (d.length === 3) return new Date(d[2], d[1] - 1, d[0], t[0], t[1]);
  }
  return new Date(valor); 
}

// --- ABA 2: CADASTRO EFETIVO ---
function buscarEfetivo(idOp) {
  const dados = getDadosAba("Cadastro_Efetivo");
  if (dados.length <= 1) return [];
  const resultado = [];
  
  for (let i = 1; i < dados.length; i++) {
    const l = dados[i];
    if (l[0] != idOp) continue;
    
    // Calcula o período de trabalho extraindo as datas corretamente
    const iniDate = parseDataGS(l[12]);
    const terDate = parseDataGS(l[13]);
    let periodo = "";
    
    if (iniDate && terDate && !isNaN(iniDate) && !isNaN(terDate)) {
      let diff = terDate.getTime() - iniDate.getTime();
      if (diff > 0) {
        let dias = Math.floor(diff / (1000 * 60 * 60 * 24));
        let horas = Math.floor((diff % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
        periodo = String(dias).padStart(2, '0') + " D " + String(horas).padStart(2, '0') + " H";
      } else {
        periodo = "00 D 00 H";
      }
    }

    resultado.push({
      linha: i + 1, 
      ord: l[2], 
      status: l[3], 
      instituicao: l[4], 
      n_bm: l[5], 
      pg: l[6], 
      nome: l[7],
      qualificacao: l[8], 
      unidade: l[9], 
      contato: l[10], 
      funcao: l[11], 
      inicio_empenho: formatarParaBR(l[12]), // Envia para a tela em formato DD/MM/YYYY
      termino: formatarParaBR(l[13]),        // Envia para a tela em formato DD/MM/YYYY
      dia_chegada: l[14], 
      dias_empenho: periodo, 
      lat: l[15] || "0", 
      long: l[16] || "0", 
      credenciamento: l[17] || "", 
      categoria: l[18] || ""
    });
  }
  return resultado;
}

function adicionarMilitarRapido(idOp, ord, inst, n_bm, pg, nome, und, cred, cat) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("Cadastro_Efetivo");
    const dadosOp = getDadosAba('Cadastro_geral');
    let nomeOperacao = "Não encontrado";
    for (let i = 1; i < dadosOp.length; i++) {
      if (dadosOp[i][0] == idOp) { nomeOperacao = dadosOp[i][1]; break; }
    }
    
    // Array com exatas 19 colunas
    aba.appendRow([
      idOp, nomeOperacao, ord, "DISPONÍVEL", inst, n_bm, pg, nome, "", und, "", "À Designar", 
      "", "", "", "0", "0", cred || "", cat || ""
    ]);
    
    limparCache('Cadastro_Efetivo');
    invalidarCacheOp(idOp);
    return aba.getLastRow();
  } catch(e) { throw new Error(e.message); }
}

function salvarEfetivo(dados) {
  try {
    const sheetEfetivo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cadastro_Efetivo");
    const dadosOp = getDadosAba('Cadastro_geral');
    let nomeOperacao = "Não encontrado";
    for (let i = 1; i < dadosOp.length; i++) { 
      if (dadosOp[i][0] == dados.idOp) { nomeOperacao = dadosOp[i][1]; break; } 
    }
    
    // Array com exatas 19 colunas mapeadas milimetricamente
    const valores = [
      dados.idOp,               // 1: ID_OP
      nomeOperacao,             // 2: NOME_OP
      dados.ord,                // 3: ORD
      dados.status,             // 4: STATUS
      dados.instituicao,        // 5: INSTITUIÇÃO
      dados.n_bm,               // 6: N_BM
      dados.pg,                 // 7: PG
      dados.nome,               // 8: NOME
      dados.qualificacao,       // 9: QUALIFICAÇÃO PROFISSIONAL
      dados.unidade,            // 10: UNIDADE
      dados.contato,            // 11: CONTATO
      dados.funcao,             // 12: FUNÇÃO
      dados.inicio_empenho,     // 13: INÍCIO_EMPENHO
      dados.termino,            // 14: TÉRMINO DO EMPENHO
      dados.dia_chegada,        // 15: DIAS_EMPENHO (Aqui fica o "DIA X")
      "0",                      // 16: LAT
      "0",                      // 17: LONG
      dados.credenciamento || "", // 18: CREDENCIADO
      dados.categoria || ""       // 19: CATEGORIA
    ];
    
    let linhaAfetada = dados.linha;
    
    if (dados.linha && !String(dados.linha).includes('sim-') && !String(dados.linha).includes('tmp-')) { 
      sheetEfetivo.getRange(dados.linha, 1, 1, valores.length).setValues([valores]);
    } else { 
      sheetEfetivo.appendRow(valores); 
      linhaAfetada = sheetEfetivo.getLastRow(); 
    }
    
    SpreadsheetApp.flush();
    limparCache('Cadastro_Efetivo');
    invalidarCacheOp(dados.idOp);
    return linhaAfetada;
    
  } catch (e) { 
    throw new Error(e.message); 
  }
}

function atualizarCelulaEfetivo(linha, coluna, valor) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cadastro_Efetivo").getRange(linha, coluna + 1).setValue(valor);
  limparCache('Cadastro_Efetivo'); 
  return true;
}

function excluirEfetivo(linha) {
  try {
    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cadastro_Efetivo");
    const idOp = aba.getRange(parseInt(linha), 1).getValue(); 
    aba.deleteRow(parseInt(linha));
    SpreadsheetApp.flush(); 
    limparCache('Cadastro_Efetivo');
    invalidarCacheOp(idOp);
    return true;
  } catch(e) { 
    throw new Error("Erro ao excluir: " + e.message); 
  }
}

function salvarEfetivoLote(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("Cadastro_Efetivo");
    const dadosOp = getDadosAba('Cadastro_geral');
    let nomeOperacao = "Operação não encontrada";
    
    for (let i = 1; i < dadosOp.length; i++) { 
      if (dadosOp[i][0] == payload.idOp) { nomeOperacao = dadosOp[i][1]; break; } 
    }
    
    const matrizInserir = payload.lote.map(item => {
       const ordCalculada = item.ord || 99;
       
       let diaTratado = item.dia_chegada || "";
       if (diaTratado && !String(diaTratado).toUpperCase().includes('DIA')) diaTratado = "DIA " + diaTratado;

       // Matriz de exatas 19 colunas!
       return [
         payload.idOp,                // 1. ID_OP
         nomeOperacao,                // 2. NOME_OP
         ordCalculada,                // 3. ORD
         "PRÉ CADASTRO",              // 4. STATUS
         item.instituicao || "CBMMG", // 5. INSTITUIÇÃO
         item.n_bm || "",             // 6. N_BM
         item.pg || "",               // 7. PG
         item.nome || "",             // 8. NOME
         item.qualificacao || "",     // 9. QUALIFICAÇÃO
         item.unidade || "",          // 10. UNIDADE
         item.contato || "",          // 11. CONTATO
         item.funcao || "À Designar", // 12. FUNÇÃO
         item.inicio_empenho || "",   // 13. INÍCIO
         item.termino || "",          // 14. TÉRMINO
         diaTratado,                  // 15. DIA_CHEGADA ("DIA 3")
         "0",                         // 16. LAT (Vazio preenchido com 0)
         "0",                         // 17. LONG (Vazio preenchido com 0)
         item.cred || "",             // 18. CREDENCIAMENTO
         item.cat || ""               // 19. CATEGORIA
       ];
    });

    if (matrizInserir.length > 0) {
       aba.getRange(aba.getLastRow() + 1, 1, matrizInserir.length, 19).setValues(matrizInserir);
    }
    
    SpreadsheetApp.flush();
    limparCache('Cadastro_Efetivo');
    invalidarCacheOp(payload.idOp);
    return matrizInserir.length;
  } catch (e) {
    throw new Error("Erro no servidor: " + e.message);
  }
}



function carregarBaseEfetivoCBMMG() {
  // Ponte direta para a função que você já tem no código (buscarTodoEfetivoCBMMG)
  return buscarTodoEfetivoCBMMG();
}

/**
 * ==========================================
 * 5. INTEGRAÇÃO E IMPORTAÇÃO DE BASES (PRODEMGE)
 * ==========================================
 */

function processEmailAndPopulateSheet() {
  // 1. Alteração Crítica: Utilizar a planilha ATIVA (nova) em vez do ID fixo da antiga
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ABA_DESTINO = "EFETIVO_CBMMG";
  const SEARCH_QUERY = "from:bimg@prodemge.gov.br has:attachment";

  const threads = GmailApp.search(SEARCH_QUERY, 0, 1);
  
  if (threads.length === 0) {
    Logger.log("Nenhum e-mail encontrado.");
    return;
  }

  const message = threads[0].getMessages()[threads[0].getMessages().length - 1];
  const attachments = message.getAttachments();
  let csvBlob = null;

  for (let i = 0; i < attachments.length; i++) {
    if (attachments[i].getName().toLowerCase().endsWith(".csv")) {
      csvBlob = attachments[i];
      break;
    }
  }

  if (!csvBlob) {
    Logger.log("Nenhum arquivo CSV encontrado.");
    return;
  }

  // 2. Proteção de caracteres: Adicionado ISO-8859-1 para evitar quebras de acentuação (ç, ã)
  const csvContent = csvBlob.getDataAsString("ISO-8859-1");
  
  // Nota: Se a Prodemge usar ponto e vírgula, troque para Utilities.parseCsv(csvContent, ";")
  const rawRows = Utilities.parseCsv(csvContent);
  
  const rowsFiltradas = [];
  
  for (let i = 0; i < rawRows.length; i++) {
    let primeiraColuna = String(rawRows[i][0]).trim();
    
    // 3. Alteração Crítica: O "i > 0" impede que o script aborte caso "Número e Dígito" seja o cabeçalho principal
    if ((primeiraColuna.includes("Número e Dígito") && i > 0) || primeiraColuna.includes("Dispensa definitiva")) {
      break; 
    }
    
    if (rawRows[i].join("").trim() !== "") {
      rowsFiltradas.push(rawRows[i]);
    }
  }

  let sheet = ss.getSheetByName(ABA_DESTINO);
  
  if (!sheet) {
    sheet = ss.insertSheet(ABA_DESTINO);
  }

  sheet.clear();
  
  if (rowsFiltradas.length > 0) {
    sheet.getRange(1, 1, rowsFiltradas.length, rowsFiltradas[0].length).setValues(rowsFiltradas);
    
    // Deixa o cabeçalho com visual padrão escuro
    sheet.getRange(1, 1, 1, rowsFiltradas[0].length).setBackground("#343a40").setFontColor("white").setFontWeight("bold");
    sheet.setFrozenRows(1);
    
    // Limpa o cache para as páginas internas enxergarem o novo efetivo instantaneamente
    CacheService.getScriptCache().remove("bdEfetivoCBMMG");
    
    Logger.log("Sucesso! Importadas " + rowsFiltradas.length + " linhas. O bloco de dispensa foi ignorado.");
  } else {
    Logger.log("Nenhum dado válido processado.");
  }
}

function buscarTodoEfetivoCBMMG() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get("bdEfetivoCBMMG");
    if (cached) { try { return JSON.parse(cached); } catch(e) { cache.remove("bdEfetivoCBMMG"); } }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("EFETIVO_CBMMG");
    if (!aba) return {};
    const dados = aba.getDataRange().getValues();
    if (dados.length < 2) return {};

    const headers = dados[0].map(h => String(h).trim().toLowerCase());
    let idxNum = headers.findIndex(h => h.includes("número") || h.includes("digito") || h.includes("bm") || h.includes("num"));
    let idxPG = headers.findIndex(h => h.includes("posto") || h.includes("graduação") || h.includes("código"));
    let idxNome = headers.findIndex(h => h.includes("nome"));
    let idxUnd = headers.findIndex(h => h.includes("unidade") || h.includes("descrição"));
    let idxCred = headers.findIndex(h => h.includes("credenciado"));
    let idxCat = headers.findIndex(h => h.includes("categoria"));

    if (idxNum === -1) idxNum = 2;
    if (idxPG === -1) idxPG = 3;
    if (idxNome === -1) idxNome = 4;
    if (idxUnd === -1) idxUnd = 5;
    if (idxCred === -1) idxCred = 8;
    if (idxCat === -1) idxCat = 9;

    let bd = {};
    for (let i = 1; i < dados.length; i++) {
      let bmLinha = String(dados[i][idxNum]).replace(/\D/g, '');
      if (bmLinha) {
        bd[bmLinha] = { 
          pg: String(dados[i][idxPG] || ""), 
          nome: String(dados[i][idxNome] || ""), 
          und: String(dados[i][idxUnd] || ""),
          credenciamento: String(dados[i][idxCred] || ""),
          categoria: String(dados[i][idxCat] || "")
        };
      }
    }

    try { cache.put("bdEfetivoCBMMG", JSON.stringify(bd), 1800); } catch(e) {}
    return bd;
  } catch (e) { return {}; }
}

function buscarViaturas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VIATURA_CBMMG");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const idxViatura = headers.findIndex(h => h.toString().trim().toUpperCase() === "VIATURA");
  const idxPlaca = headers.findIndex(h => h.toString().trim().toUpperCase() === "PLACA");
  if (idxViatura === -1) return []; 

  const viaturas = [];
  for (let i = 1; i < data.length; i++) {
    let nomeVtr = data[i][idxViatura] ? data[i][idxViatura].toString().trim() : "";
    let placaVtr = (idxPlaca !== -1 && data[i][idxPlaca]) ? data[i][idxPlaca].toString().trim() : "";
    if (nomeVtr !== "") {
      viaturas.push({ viatura: nomeVtr, placa: placaVtr });
    }
  }
  return viaturas;
}

// Função para buscar os municípios do IBGE com Cache (Fornecida por você)
function getMunicipiosMG() {
  const cache = CacheService.getScriptCache();
  const cachedMunicipios = cache.get("MunicipiosMG");
  if (cachedMunicipios) return JSON.parse(cachedMunicipios);
  
  try {
    const res = UrlFetchApp.fetch("https://servicodados.ibge.gov.br/api/v1/localidades/estados/31/municipios");
    const muns = JSON.parse(res.getContentText()).map(m => m.nome).sort();
    cache.put("MunicipiosMG", JSON.stringify(muns), 21600); // Salva por 6 horas
    return muns;
  } catch(e) { 
    return []; 
  }
}

// ========================================================
// ATUALIZAÇÃO: Salva Nova Operação OU Edita (Busca Nativa INFALÍVEL)
// ========================================================
function salvarOperacaoPlanilha(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("Cadastro_geral");
  
  // Limpa qualquer micro-espaço que possa ter vindo do formulário
  const idProcurado = String(dados.id).trim(); 
  
  // A ARMA SECRETA: Usa a busca nativa do Google apenas na Coluna A (Como um Ctrl+F)
  const busca = aba.getRange("A:A").createTextFinder(idProcurado).matchEntireCell(true).findNext();
  
  if (busca) { 
    // ====================================================
    // MODO EDIÇÃO: O ID foi encontrado perfeitamente!
    // ====================================================
    const linhaExata = busca.getRow(); // Pega exatamente o número da linha onde achou
    const linhaExistente = aba.getRange(linhaExata, 1, 1, 13).getValues()[0];
    
    aba.getRange(linhaExata, 1, 1, 13).setValues([[
      dados.id,               // 1. ID (Mantém Intacto)
      dados.nome,             // 2. NOME
      dados.cob,              // 3. COB
      dados.municipios,       // 4. MUNICÍPIO
      dados.chamada,          // 5. CHAMADA
      dados.cmt,              // 6. CMT
      dados.status,           // 7. STATUS
      dados.inicio,           // 8. INÍCIO
      dados.termino,          // 9. TÉRMINO
      dados.duracao,          // 10. DURAÇÃO (Calculada e salva no banco!
      dados.uc,               // 11. UC
      linhaExistente[11],     // 12. EFETIVO (Preserva o valor existente)
      linhaExistente[12]      // 13. VTRS (Preserva o valor existente)
    ]]);
    
  } else {
    // ====================================================
    // MODO CRIAÇÃO: ID não existe. Cria nova linha no final.
    // ====================================================
    aba.appendRow([
      dados.id,
      dados.nome,
      dados.cob,
      dados.municipios,
      dados.chamada,
      dados.cmt,
      dados.status,
      dados.inicio,
      dados.termino,
      dados.duracao,
      dados.uc,
      "",
      ""
    ]);
  }
  
  return true;
}


// ========================================================
// NOVA FUNÇÃO: Busca as UCs de forma automática e inteligente
// ========================================================
function getUCs() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("UC");
    
    // Procura dinamicamente em qual coluna está o cabeçalho "UC"
    const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
    const colIndex = headers.indexOf("UC") + 1; // +1 pois a matriz começa em 0 e as colunas em 1
    
    if (colIndex === 0) return []; // Se não achar a coluna UC, retorna vazio
    
    // Pega todos os nomes da coluna UC (pulando a linha 1 do cabeçalho)
    const ucs = aba.getRange(2, colIndex, aba.getLastRow() - 1, 1).getValues().flat().filter(String);
    
    // Remove duplicadas e organiza em ordem alfabética
    return [...new Set(ucs)].sort(); 
  } catch(e) {
    return [];
  }
}


// ========================================================
// NOVA FUNÇÃO: Busca as operações e blinda contra erros de Data
// ========================================================
function getOperacoes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("Cadastro_geral");
    const ultRow = aba.getLastRow();

    // Se a planilha estiver vazia (só cabeçalho), retorna lista vazia
    if (ultRow <= 1) return []; 

    // Pega os dados pulando o cabeçalho
    const dados = aba.getRange(2, 1, ultRow - 1, 13).getValues();

    const operacoes = dados.map(linha => {
      
      // BLINDAGEM: Garante que se houver uma data real na planilha, ela vira Texto (String)
      let dataInicio = linha[7];
      if (dataInicio instanceof Date) {
        dataInicio = Utilities.formatDate(dataInicio, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      } else {
        dataInicio = String(dataInicio || "");
      }

      let dataTermino = linha[8];
      if (dataTermino instanceof Date) {
        dataTermino = Utilities.formatDate(dataTermino, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      } else {
        dataTermino = String(dataTermino || "-");
      }

      // Constrói o pacote forçando tudo a ser Texto (String) para não travar o envio
      return {
        id: String(linha[0] || ""),
        nome: String(linha[1] || ""),
        cob: String(linha[2] || ""),
        municipio: String(linha[3] || ""),
        chamada: String(linha[4] || ""),
        comandante: String(linha[5] || "Não Informado"),
        status: String(linha[6] || ""),
        inicio: dataInicio,
        termino: dataTermino,
        duracao: String(linha[9] || ""),
        uc: String(linha[10] || ""),
        efetivo: linha[11] || 0,
        vtrs: linha[12] || 0
      };
    });

    return enriquecerOperacoesComRecursos(operacoes.reverse());

  } catch(erro) {
    // Se der qualquer outro erro na planilha, retorna vazio em vez de travar
    return []; 
  }
}

/**
 * ==========================================
 * SISTEMA DE LOGIN E CONTROLE DE PERFIS
 * ==========================================
 */
function validarLoginGS(usuarioInput, senhaInput, cobDesejado) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaUsuarios = ss.getSheetByName("Usuarios");
    
    // Se a aba não existir, avisa
    if (!abaUsuarios) {
      return { sucesso: false, mensagem: "Aba 'Usuarios' não encontrada na planilha." };
    }

    const dados = abaUsuarios.getDataRange().getValues();
    
    // Começa do i = 1 para pular o cabeçalho
    for (let i = 1; i < dados.length; i++) {
      let userPlanilha = String(dados[i][0]).trim();
      let senhaPlanilha = String(dados[i][1]).trim();
      let perfilPlanilha = String(dados[i][2]).trim();

      // Verifica se o usuário bate
      if (userPlanilha === String(usuarioInput).trim()) {
        
        // Verifica se a senha bate
        if (senhaPlanilha === String(senhaInput).trim()) {
          
          // REGRA DE NEGÓCIO: CEB acessa tudo. COB acessa apenas o próprio COB.
          if (perfilPlanilha === "CEB" || perfilPlanilha === cobDesejado) {
            return { sucesso: true, mensagem: "Login aprovado", perfil: perfilPlanilha };
          } else {
            return { 
              sucesso: false, 
              mensagem: `Acesso bloqueado. Seu perfil (${perfilPlanilha}) não permite acessar o painel do ${cobDesejado}.` 
            };
          }
          
        } else {
          return { sucesso: false, mensagem: "Senha incorreta." };
        }
      }
    }
    
    // Se varreu a planilha toda e não achou o usuário
    return { sucesso: false, mensagem: "Usuário não encontrado." };

  } catch (e) {
    return { sucesso: false, mensagem: "Erro no servidor: " + e.message };
  }
}

// ====================================================================
// FUNÇÃO BLINDADA: SOMA EFETIVO E VIATURAS POR OPERAÇÃO
// ====================================================================
function enriquecerOperacoesComRecursos(operacoes) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // -----------------------------------------------------
    // 1. CONTAGEM DE EFETIVO (Aba: Cadastro_Efetivo)
    // -----------------------------------------------------
    let mapaEfetivo = {};
    const abaEfetivo = ss.getSheetByName("Cadastro_Efetivo");
    
    if (abaEfetivo && abaEfetivo.getLastRow() > 1) {
      const dadosEf = abaEfetivo.getDataRange().getValues();
      
      // Limpa os cabeçalhos para evitar erros de digitação (ex: "ID_OP ")
      const headersEf = dadosEf.shift().map(h => String(h).trim().toUpperCase());
      const idxIdOpEf = headersEf.indexOf("ID_OP");
      const idxStatusEf = headersEf.indexOf("STATUS");
      
      if (idxIdOpEf > -1 && idxStatusEf > -1) {
        dadosEf.forEach(linha => {
          let idOperacao = String(linha[idxIdOpEf] || "").trim();
          let status = String(linha[idxStatusEf] || "").trim().toUpperCase();
          
          if (idOperacao !== "" && status !== "PRÉ CADASTRO") {
            mapaEfetivo[idOperacao] = (mapaEfetivo[idOperacao] || 0) + 1;
          }
        });
      }
    }

    // -----------------------------------------------------
    // 2. CONTAGEM DE VIATURAS (Aba: Logistica)
    // -----------------------------------------------------
    let mapaVtrs = {};
    const abaLog = ss.getSheetByName("Logistica"); 
    
    if (abaLog && abaLog.getLastRow() > 1) {
      const dadosLog = abaLog.getDataRange().getValues();
      
      // Limpa cabeçalhos
      const headersLog = dadosLog.shift().map(h => String(h).trim().toUpperCase());
      const idxIdOpLog = headersLog.indexOf("ID_OP"); 
      const idxTipoLog = headersLog.indexOf("TIPO RECURSO"); 
      const idxQtdLog = headersLog.indexOf("QUANTIDADE"); 
      
      if (idxIdOpLog > -1 && idxTipoLog > -1 && idxQtdLog > -1) {
        dadosLog.forEach(linha => {
          let idOperacao = String(linha[idxIdOpLog] || "").trim();
          let tipoRecurso = String(linha[idxTipoLog] || "").trim().toUpperCase();
          let quantidade = parseFloat(linha[idxQtdLog]) || 0;
          
          // Confirma se é VIATURAS
          if (idOperacao !== "" && tipoRecurso === "VIATURAS") {
            mapaVtrs[idOperacao] = (mapaVtrs[idOperacao] || 0) + quantidade;
          }
        });
      }
    }

    // -----------------------------------------------------
    // 3. INJETAR OS TOTAIS E DEVOLVER PARA A TELA
    // -----------------------------------------------------
    operacoes.forEach(op => {
      let idBusca = String(op.id).trim(); // Garante a correspondência do ID
      op.efetivo = mapaEfetivo[idBusca] || 0;
      op.vtrs = mapaVtrs[idBusca] || 0;
    });
    
    return operacoes;
    
  } catch (e) {
    // Se ocorrer um erro estranho, devolve a lista intacta em vez de quebrar o site
    return operacoes;
  }
}

// ==========================================
// FUNÇÃO PARA SALVAR FEEDBACKS DOS USUÁRIOS
// ==========================================
function salvarFeedbackGS(dados) {
  try {
    // Como o script está ligado à planilha, ele já pega ela automaticamente
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Procura pela aba com o nome exato "Feedbacks"
    var abaFeedbacks = ss.getSheetByName('Feedbacks');
    
    // Se por acaso a aba for apagada no futuro, ele recria automaticamente
    if (!abaFeedbacks) {
      abaFeedbacks = ss.insertSheet('Feedbacks');
      abaFeedbacks.appendRow(['DATA / HORA', 'USUÁRIO', 'MÓDULO', 'SUGESTÃO', 'STATUS']);
      abaFeedbacks.getRange("A1:E1").setFontWeight("bold").setBackground("#343a40").setFontColor("white");
    }
    
    // Insere a nova linha de sugestão na aba Feedbacks
    abaFeedbacks.appendRow([
      dados.data,
      dados.usuario,
      dados.modulo,
      dados.sugestao,
      "Pendente" 
    ]);
    
    return true;
  } catch (e) {
    throw new Error("Erro ao salvar feedback: " + e.message);
  }
}
