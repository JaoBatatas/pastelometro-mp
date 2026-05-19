function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dados");
  var action = e.parameter ? (e.parameter.action || "consultar") : "consultar";
  
  var agora = new Date();
  var dataHoje = Utilities.formatDate(agora, "America/Sao_Paulo", "dd/MM/yyyy");
  var horaFormatada = Utilities.formatDate(agora, "America/Sao_Paulo", "HH:mm");
  
  var ultimaLinha = sheet.getLastRow();
  var dadosUltimaLinha = ultimaLinha > 1 ? sheet.getRange(ultimaLinha, 1, 1, 4).getValues()[0] : ["", "", "", ""];
  
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

  // 1. Marcar como Pronto
  if (action === "registrar_pronto") {
    if (dataUltimoRegistro !== dataHoje) {
      var horas = agora.getHours();
      var minutos = agora.getMinutes();
      var minutosDoDia = (horas * 60) + minutos;
      sheet.appendRow([dataHoje, horaFormatada, minutosDoDia, ""]);
      
      ultimaLinha = sheet.getLastRow();
      dataUltimoRegistro = dataHoje;
      horaAcabouUltimo = "";
    }
  }
  
  // 2. Marcar como Acabou
  if (action === "registrar_acabou") {
    if (dataUltimoRegistro === dataHoje && horaAcabouUltimo === "") {
      sheet.getRange(ultimaLinha, 4).setValue(horaFormatada);
      SpreadsheetApp.flush(); 
      horaAcabouUltimo = horaFormatada; 
    }
  }
  
  var dados = sheet.getDataRange().getValues();
  var listaMinutosPronto = [];
  var listaMinutosAcabou = []; // Nova lista para os horários de término
  var statusHoje = "aguardando";
  var horaProntoHoje = "";

  for (var i = 1; i < dados.length; i++) {
    // Recolhe coluna C (Minutos Pronto)
    if (dados[i][2] !== "") {
      listaMinutosPronto.push(Number(dados[i][2]));
    }
    // Recolhe coluna D (Hora Acabou) e converte para minutos
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
    historicoAcabou: listaMinutosAcabou, // Envia o novo histórico
    statusHoje: statusHoje,
    horaPronto: horaProntoHoje
  };
  
  return ContentService.createTextOutput(JSON.stringify(resposta))
    .setMimeType(ContentService.MimeType.JSON);
}
