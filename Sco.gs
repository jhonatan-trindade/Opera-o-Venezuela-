/**
 * ==========================================
 * BACKEND: FORMULÁRIO SCO (Sistema de Comando Operacional)
 * ==========================================
 * Aba SCO (cabeçalho) + SCO_Acoes (ações filhas 1:N por token relacional).
 */

/**
 * Lista os nomes do efetivo da operação ativa (datalist do responsável do SCO).
 */
function obterNomesParaPesquisaSCO(idOpAtiva) {
  try {
    const abaEfetivo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cadastro_Efetivo");
    if (!abaEfetivo) return [];
    const dados = abaEfetivo.getDataRange().getValues();
    
    const listaNomes = [];
    for (let i = 1; i < dados.length; i++) {
      let idOpLinha = String(dados[i][0]).trim();
      let nomeMilitar = String(dados[i][7]).trim(); // Coluna H = NOME
      if (idOpLinha === String(idOpAtiva) && nomeMilitar && !listaNomes.includes(nomeMilitar)) {
        listaNomes.push(nomeMilitar);
      }
    }
    return listaNomes.sort();
  } catch (e) { return []; }
}

/**
 * Cria um novo formulário OU atualiza um existente, desmembrando as ações.
 */
function salvarDadosSCO(dados) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaSCO = ss.getSheetByName("SCO");
    const abaAcoes = ss.getSheetByName("SCO_Acoes");
    
    if (!abaSCO || !abaAcoes) throw new Error("As abas 'SCO' ou 'SCO_Acoes' não foram encontradas.");

    const formatarDataBR = function(dataString) {
      if (!dataString) return "";
      if (dataString.includes("/")) return dataString;
      const partes = dataString.split("T");
      if (partes.length !== 2) return dataString;
      const dataPartes = partes[0].split("-");
      return `${dataPartes[2]}/${dataPartes[1]}/${dataPartes[0]} ${partes[1]}`;
    };

    let tokenRelacional = "";

    // MODO EDIÇÃO
    if (dados.modoEdicao && dados.linha && dados.token) {
      tokenRelacional = dados.token;
      
      // Matriz exata de 15 colunas conforme seu banco
      const linhaAtualizada = [
        dados.idOp,                           // 1. ID_OP
        dados.nomeOp,                         // 2. NOME_OP
        formatarDataBR(dados.cicloInicio),    // 3. CICLO_INICIO
        formatarDataBR(dados.cicloFim),       // 4. CICLO_FIM
        dados.dia,                            // 5. DIA
        dados.resumo,                         // 6. RESUMO
        dados.objetivos,                      // 7. OBJETIVO
        dados.enfase,                         // 8. SENFASE_PERIODO
        dados.situacao,                       // 9. SITUACAO
        dados.nomeResp,                       // 10. NOME
        dados.cargo,                          // 11. CARGO
        formatarDataBR(dados.dataRegistro),   // 12. DATA (Data do preenchimento)
        tokenRelacional,                      // 13. AÇÕES (Token FK)
        "0",                                  // 14. LAT
        "0",                                  // 15. LONG
        dados.municipio || ""
      ];
      abaSCO.getRange(dados.linha, 1, 1, linhaAtualizada.length).setValues([linhaAtualizada]);

      // Apaga as ações antigas desse Token para reescrever as novas
      const dadosAcoes = abaAcoes.getDataRange().getValues();
      for (let j = dadosAcoes.length - 1; j > 0; j--) {
        if (String(dadosAcoes[j][0]).trim() === tokenRelacional) abaAcoes.deleteRow(j + 1);
      }
      
    } else {
      // MODO CRIAÇÃO (Validação de duplicidade: operação + dia)
      const dadosSCO = abaSCO.getDataRange().getValues();
      const idNovo = String(dados.idOp).trim();
      const diaNovo = String(dados.dia).trim().toUpperCase();
      
      for (let i = 1; i < dadosSCO.length; i++) {
        let idBanco = String(dadosSCO[i][0]).trim();
        let diaBanco = String(dadosSCO[i][4]).trim().toUpperCase(); // Coluna 5 é o DIA
        if (idBanco === idNovo && diaBanco === diaNovo) {
          return { sucesso: false, mensagem: "Já existe um planejamento preenchido para este DIA." };
        }
      }

      tokenRelacional = "SCO-" + Utilities.getUuid().slice(0, 8).toUpperCase();
      
      const novaLinhaSCO = [
        dados.idOp, dados.nomeOp, formatarDataBR(dados.cicloInicio), formatarDataBR(dados.cicloFim),
        dados.dia, dados.resumo, dados.objetivos, dados.enfase, dados.situacao, dados.nomeResp, 
        dados.cargo, formatarDataBR(dados.dataRegistro), tokenRelacional, "0", "0", dados.municipio || ""
      ];
      abaSCO.appendRow(novaLinhaSCO);
    }

    // Salva as ações filhas (se existirem) na aba SCO_Acoes
    // Coluna FOTOS guarda um JSON [{url,id,nome}] com as imagens anexadas à ação.
    if (dados.listaAcoes && dados.listaAcoes.length > 0) {
      const arrayAcoesInserir = dados.listaAcoes.map(item => [
        tokenRelacional,
        item.hora,
        item.texto,
        JSON.stringify(item.fotos || []),
        0,
        0
      ]);
      abaAcoes.getRange(abaAcoes.getLastRow() + 1, 1, arrayAcoesInserir.length, 6).setValues(arrayAcoesInserir);
    }
    
    return { sucesso: true, mensagem: dados.modoEdicao ? "Formulário atualizado com sucesso!" : "Formulário salvo com sucesso!" };
  } catch (e) { return { sucesso: false, mensagem: "Erro interno: " + e.message }; }
}

/**
 * Carrega o histórico de formulários da operação, juntando as ações ([HH:mm] texto).
 */
function carregarHistoricoSCO(idOpAtiva) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaSCO = ss.getSheetByName("SCO");
    const abaAcoes = ss.getSheetByName("SCO_Acoes");
    
    if (!abaSCO) return [];
    const dadosSCO = abaSCO.getDataRange().getValues();
    if (dadosSCO.length <= 1) return [];

    const mapaAcoes = {};
    if (abaAcoes) {
      const dadosAcoes = abaAcoes.getDataRange().getValues();
      for (let j = 1; j < dadosAcoes.length; j++) {
        let fk = String(dadosAcoes[j][0]).trim();
        let horaBruta = dadosAcoes[j][1];
        let horaFormatada = "";
        
        if (horaBruta instanceof Date) {
          horaFormatada = Utilities.formatDate(horaBruta, Session.getScriptTimeZone(), "HH:mm");
        } else {
          let match = String(horaBruta).match(/\d{2}:\d{2}/);
          horaFormatada = match ? match[0] : String(horaBruta).trim();
        }
        
        let acao = String(dadosAcoes[j][2]).trim();

        // Coluna 4 (índice 3) = FOTOS em JSON. Blindado contra dados legados/inválidos.
        let fotos = [];
        try {
          let brutoFotos = dadosAcoes[j][3];
          if (brutoFotos && String(brutoFotos).trim().charAt(0) === '[') {
            fotos = JSON.parse(brutoFotos);
          }
        } catch (eF) { fotos = []; }

        if (!mapaAcoes[fk]) mapaAcoes[fk] = [];
        mapaAcoes[fk].push({ hora: horaFormatada, texto: acao, fotos: fotos });
      }
    }

    const formatarDataBR = function(val) {
      if (!val) return "";
      if (val instanceof Date) return Utilities.formatDate(val, "GMT-3", "dd/MM/yyyy HH:mm");
      const tentativa = new Date(String(val));
      if (!isNaN(tentativa.getTime())) return Utilities.formatDate(tentativa, "GMT-3", "dd/MM/yyyy HH:mm");
      return String(val);
    };

    const historico = [];
    for (let i = 1; i < dadosSCO.length; i++) {
      if (String(dadosSCO[i][0]).trim() === String(idOpAtiva).trim()) {
        let tokenSCO = String(dadosSCO[i][12]).trim(); // Coluna 13 (índice 12) é o TOKEN
        
        let stringAcoesResumo = "Nenhuma ação registrada.";
        let arrayAcoesOriginal = [];
        
        if (mapaAcoes[tokenSCO]) {
            arrayAcoesOriginal = mapaAcoes[tokenSCO];
            stringAcoesResumo = mapaAcoes[tokenSCO].map(a => `[${a.hora}] ${a.texto}`).join(" | ");
        }

        historico.push({
          linha: i + 1, 
          token: tokenSCO, 
          nomeOp: String(dadosSCO[i][1]), 
          cicloInicio: formatarDataBR(dadosSCO[i][2]), 
          cicloFim: formatarDataBR(dadosSCO[i][3]), 
          dia: String(dadosSCO[i][4]),
          resumo: String(dadosSCO[i][5]), 
          objetivos: String(dadosSCO[i][6]), 
          enfase: String(dadosSCO[i][7]),
          situacao: String(dadosSCO[i][8]), 
          nomeResp: String(dadosSCO[i][9]), 
          cargo: String(dadosSCO[i][10]),
          dataRegistro: formatarDataBR(dadosSCO[i][11]), 
          municipio: String(dadosSCO[i][15] || ""), // <---- INSERIDO AQUI A BUSCA DA COLUNA P
          acoesList: arrayAcoesOriginal,
          acoesStr: stringAcoesResumo
        });
      }
    }
    return historico.reverse();
  } catch (e) { return []; }
}

/**
 * Recebe uma foto em base64 vinda da tela, salva no Google Drive dentro de uma
 * pasta organizada por operação e devolve a URL pública para exibir/imprimir.
 * Chamada por google.script.run a partir do formulário de ações do SCO.
 */
function uploadFotoSCO(payload) {
  try {
    const idOp = String(payload.idOp || "geral").trim();
    const nomeOp = String(payload.nomeOp || "OPERACAO").trim();
    const mime = payload.mime || "image/jpeg";
    const nomeArquivo = payload.nome || ("foto_" + new Date().getTime() + ".jpg");

    // Decodifica o base64 (removendo o cabeçalho data:...;base64, se vier junto)
    let dadosBase64 = String(payload.base64 || "");
    const virgula = dadosBase64.indexOf(",");
    if (virgula > -1) dadosBase64 = dadosBase64.substring(virgula + 1);

    const bytes = Utilities.base64Decode(dadosBase64);
    const blob = Utilities.newBlob(bytes, mime, nomeArquivo);

    // Estrutura de pastas: GeoFOGO_SCO / <NOME_OP (ID)>
    const pastaRaiz = obterOuCriarPasta_(DriveApp.getRootFolder(), "GeoFOGO_SCO");
    const pastaOp = obterOuCriarPasta_(pastaRaiz, nomeOp + " (" + idOp + ")");

    const arquivo = pastaOp.createFile(blob);
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const id = arquivo.getId();
    return {
      sucesso: true,
      id: id,
      nome: nomeArquivo,
      url: "https://drive.google.com/uc?export=view&id=" + id,
      link: arquivo.getUrl()
    };
  } catch (e) {
    return { sucesso: false, mensagem: "Falha no upload da foto: " + e.message };
  }
}

/** Retorna a subpasta pelo nome dentro de "pai", criando-a se não existir. */
function obterOuCriarPasta_(pai, nome) {
  const existentes = pai.getFoldersByName(nome);
  if (existentes.hasNext()) return existentes.next();
  return pai.createFolder(nome);
}

function excluirRegistroSCO(linha, token) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaSCO = ss.getSheetByName("SCO");
    const abaAcoes = ss.getSheetByName("SCO_Acoes");
    
    abaSCO.deleteRow(linha);
    
    if (abaAcoes && token) {
      const dadosAcoes = abaAcoes.getDataRange().getValues();
      for (let j = dadosAcoes.length - 1; j > 0; j--) {
        if (String(dadosAcoes[j][0]).trim() === token) abaAcoes.deleteRow(j + 1);
      }
    }
    return { sucesso: true, mensagem: "Planejamento SCO e ações excluídos com sucesso!" };
  } catch (e) { return { sucesso: false, mensagem: "Erro ao excluir: " + e.message }; }
}
