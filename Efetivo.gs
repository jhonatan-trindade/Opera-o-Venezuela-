/**
 * ==========================================
 * MÓDULO: GESTÃO DE EFETIVO
 * ==========================================
 * Este arquivo concentra todas as regras de negócio, 
 * cálculos e comunicação com o banco de dados (Planilha) 
 * referentes à aba de Cadastro de Efetivo.
 */

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
      inicio_empenho: formatarParaBR(l[12]), 
      termino: formatarParaBR(l[13]),        
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
         diaTratado,                  // 15. DIA_CHEGADA
         "0",                         // 16. LAT
         "0",                         // 17. LONG
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
  return buscarTodoEfetivoCBMMG();
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

    // Fallbacks para as posições que vimos anteriormente
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
