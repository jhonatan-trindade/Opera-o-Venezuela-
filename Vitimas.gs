/**
 * ==========================================
 * MÓDULO: REGISTRO DE VÍTIMAS
 * ==========================================
 * Backend do registro de vítimas da operação de resgate.
 * Aba "Vitimas": ID_OP, NOME_OP, DIA, HORA, NOME, IDADE, SEXO, STATUS,
 * LOCAL, DESCRICAO, ENCAMINHAMENTO, DATA_REGISTRO, LAT, LONG
 * Segue o mesmo padrão de CRUD e cache do módulo de Efetivo.
 */

function buscarVitimas(idOp) {
  const dados = getDadosAba("Vitimas");
  if (dados.length <= 1) return [];
  const resultado = [];

  for (let i = 1; i < dados.length; i++) {
    const l = dados[i];
    if (l[0] != idOp) continue;

    resultado.push({
      linha: i + 1,
      dia: l[2] || "",
      hora: l[3] || "",
      nome: l[4] || "",
      idade: l[5] || "",
      sexo: l[6] || "",
      status: l[7] || "",
      local: l[8] || "",
      descricao: l[9] || "",
      encaminhamento: l[10] || "",
      data_registro: formatarParaBR(l[11]),
      lat: l[12] || "0",
      long: l[13] || "0"
    });
  }
  return resultado;
}

function salvarVitima(dados) {
  try {
    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vitimas");
    const dadosOp = getDadosAba('Cadastro_geral');
    let nomeOperacao = "Não encontrado";
    for (let i = 1; i < dadosOp.length; i++) {
      if (dadosOp[i][0] == dados.idOp) { nomeOperacao = dadosOp[i][1]; break; }
    }

    // 14 colunas na ordem exata da aba Vitimas
    const valores = [
      dados.idOp,                                   // 1: ID_OP
      nomeOperacao,                                 // 2: NOME_OP
      dados.dia || "",                              // 3: DIA
      dados.hora || "",                             // 4: HORA
      dados.nome || "",                             // 5: NOME
      dados.idade || "",                            // 6: IDADE
      dados.sexo || "",                             // 7: SEXO
      dados.status || "",                           // 8: STATUS
      dados.local || "",                            // 9: LOCAL
      dados.descricao || "",                        // 10: DESCRICAO
      dados.encaminhamento || "",                   // 11: ENCAMINHAMENTO
      dados.data_registro || formatarParaBR(new Date()), // 12: DATA_REGISTRO
      dados.lat || "0",                             // 13: LAT
      dados.long || "0"                             // 14: LONG
    ];

    let linhaAfetada = dados.linha;

    if (dados.linha && !String(dados.linha).includes('tmp-')) {
      aba.getRange(dados.linha, 1, 1, valores.length).setValues([valores]);
    } else {
      aba.appendRow(valores);
      linhaAfetada = aba.getLastRow();
    }

    SpreadsheetApp.flush();
    limparCache('Vitimas');
    invalidarCacheOp(dados.idOp);
    return { sucesso: true, linha: linhaAfetada, mensagem: "Vítima registrada com sucesso!" };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao salvar vítima: " + e.message };
  }
}

function excluirVitima(linha) {
  try {
    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vitimas");
    const idOp = aba.getRange(parseInt(linha), 1).getValue();
    aba.deleteRow(parseInt(linha));
    SpreadsheetApp.flush();
    limparCache('Vitimas');
    invalidarCacheOp(idOp);
    return { sucesso: true, mensagem: "Registro de vítima excluído." };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao excluir: " + e.message };
  }
}
