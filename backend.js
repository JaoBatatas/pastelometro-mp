function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDados = ss.getSheetByName("Dados");
  var sheetBloqueados = ss.getSheetByName("Bloqueados");
  
  // Cria automaticamente a aba de Bloqueados se ela não existir
  if (!sheetBloqueados) {
    sheetBloqueados = ss.insertSheet("Bloqueados");
    sheetBloqueados.appendRow(["Identificador (IP ou UID)"]);
  }
  
  // Captura os dados de identificação enviados pelo frontend
  var userIp = e.parameter.ip || "0.0.0.0";
  var userUid = e.parameter.uid || "anonimo";
  
  // Verifica se o IP ou o UID do usuário estão na lista negra
  var bloqueado = false;
  var lastRowBloqueados = sheetBloqueados.getLastRow();
  if (lastRowBloqueados > 1) {
    var listaBloqueados = sheetBloqueados.getRange(2, 1, lastRowBloqueados - 1, 1).getValues();
    bloqueado = listaBloqueados.some(function(row) { 
      var itemBloqueado = String(row[0]).trim();
      return itemBloqueado === String(userIp).trim() || itemBloqueado === String(userUid).trim(); 
    });
  }
  
  // Bloqueio imediato caso detectado na lista negra
  if (bloqueado) {
    return ContentService.createTextOutput(JSON.stringify({ status: "bloqueado", mensagem: "Acesso negado." }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var action = e.parameter ? (e.parameter.action || "consultar") : "consultar";  
  
  var agora = new Date();
  var dataHoje = Utilities.formatDate(agora, "America/Sao_Paulo", "dd/MM/yyyy");
  var horaFormatada = Utilities.formatDate(agora, "America/Sao_Paulo", "HH:mm");
  
  var ultimaLinha = sheetDados.getLastRow();
  // Lemos até a coluna H (8 colunas) para capturar os novos logs de UID
  var dadosUltimaLinha = ultimaLinha > 1 ? sheetDados.getRange(ultimaLinha, 1, 1, 8).getValues()[0] : ["", "", "", "", "", "", "", ""];
  
  function formatarDataPlanilha(d) {
    if (!d) return "";
    if (d instanceof Date || Object.prototype.toString.call(d) === '[object Date]') {
      return Utilities.formatDate(d, "America/Sao_Paulo", "dd/MM/yyyy");
    }
    return String(d).trim().split(" ")[0];
  }

  function formatarHoraPlanilha(h) {
    if (!h) return "";
    if (h instanceof Date || Object.prototype.toString.call(h) === '[object Date]') {
      return Utilities.formatDate(h, "America/Sao_Paulo", "HH:mm");
    }
    return String(h).trim();
  }

  var dataUltimoRegistro = formatarDataPlanilha(dadosUltimaLinha[0]);
  var horaAcabouUltimo = formatarHoraPlanilha(dadosUltimaLinha[3]); 

  // 1. AÇÃO: Registrar que o pastel ficou PRONTO
  if (action === "registrar_pronto") {
    if (dataUltimoRegistro !== dataHoje) {
      var horas = agora.getHours();
      var minutes = agora.getMinutes();
      var minutosDoDia = (horas * 60) + minutes;
      
      // Nova estrutura de colunas:
      // A: Data | B: Hora | C: Minutos | D: Hora_Acabou | E: IP_Pronto | F: IP_Acabou | G: UID_Pronto | H: UID_Acabou
      sheetDados.appendRow([dataHoje, horaFormatada, minutosDoDia, "", userIp, "", userUid, ""]);
      
      ultimaLinha = sheetDados.getLastRow();
      dataUltimoRegistro = dataHoje;
      horaAcabouUltimo = "";
    }
  }
  
  // 2. AÇÃO: Registrar que o pastel ACABOU
  if (action === "registrar_acabou") {
    if (dataUltimoRegistro === dataHoje && horaAcabouUltimo === "") {
      sheetDados.getRange(ultimaLinha, 4).setValue(horaFormatada); // Coluna D
      sheetDados.getRange(ultimaLinha, 6).setValue(userIp);        // Coluna F
      sheetDados.getRange(ultimaLinha, 8).setValue(userUid);       // Coluna H (UID de quem encerrou)
      SpreadsheetApp.flush(); 
      horaAcabouUltimo = horaFormatada; 
    }
  }
  
  // 3. CONSULTA: Lê os dados históricos para as estatísticas
  var dados = sheetDados.getDataRange().getValues();
  var listaMinutosPronto = [];
  var listaMinutosAcabou = [];
  var statusHoje = "aguardando";
  var horaProntoHoje = "";

  for (var i = 1; i < dados.length; i++) {
    if (dados[i][2] !== "") {
      listaMinutosPronto.push(Number(dados[i][2]));
    }
    var horaAcabouStr = formatarHoraPlanilha(dados[i][3]);
    if (horaAcabouStr !== "") {
       var partes = horaAcabouStr.split(':');
       if (partes.length === 2) {
          var minAcabou = parseInt(partes[0], 10) * 60 + parseInt(partes[1], 10);
          listaMinutosAcabou.push(minAcabou);
       }
    }
  }
  
  if (dados.length > 1) {
    var ultima = dados[dados.length - 1];
    var dataUltimaCalculada = formatarDataPlanilha(ultima[0]);
    
    if (dataUltimaCalculada === dataHoje) {
      horaProntoHoje = formatarHoraPlanilha(ultima[1]); 
      
      if (formatarHoraPlanilha(ultima[3]) === "") {
        statusHoje = "disponivel"; 
      } else {
        statusHoje = "esgotado"; 
      }
    }
  }
  
  var resposta = {
    historico: listaMinutosPronto,
    historicoAcabou: listaMinutosAcabou,
    statusHoje: statusHoje,
    horaPronto: horaProntoHoje
  };
  
  return ContentService.createTextOutput(JSON.stringify(resposta))
    .setMimeType(ContentService.MimeType.JSON);
}