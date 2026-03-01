# Apps Script — Monitor Drive Colheita

Cole esse código no Apps Script e execute `installTrigger`.

```javascript
var TELEGRAM_BOT_TOKEN = "8576612955:AAFGyDzDeAc5TbiL4NUngHEN-aWCV45Nh2s";
var TELEGRAM_CHAT_ID = "1035058177";
var XLSX_ID = "1E7rfOMux-VqPvlhHdibjYqDoBxh6CsPq";
var LAST_MODIFIED_KEY = "lastModified_ColheitaSoja";
var WEBHOOK_URL = "https://terra-bi-app.vercel.app/api/sync?secret=terra-bi-2026";

function checkDriveUpdates() {
  var props = PropertiesService.getScriptProperties();
  var lastChecked = props.getProperty(LAST_MODIFIED_KEY);
  var lastDate = lastChecked ? new Date(lastChecked) : new Date(0);
  var xlsx = DriveApp.getFileById(XLSX_ID);
  if (xlsx.getLastUpdated() <= lastDate) return;
  try {
    var data = extractData();
    UrlFetchApp.fetch(WEBHOOK_URL, {
      method: "post", contentType: "application/json",
      payload: JSON.stringify(data), muteHttpExceptions: true
    });
    props.setProperty(LAST_MODIFIED_KEY, new Date().toISOString());
    sendTelegram("🌾 *Dashboard atualizado!*\n\n• " +
      data.resumo.area_colhida.toFixed(0) + " ha colhidos (" +
      (data.resumo.pct_colhido*100).toFixed(1) + "%)\n• " +
      data.resumo.total_colhido.toFixed(0) + " sacas\n• " +
      data.resumo.media_geral.toFixed(2) + " sc/ha\n\n" +
      "👉 https://terra-bi-app.vercel.app");
  } catch(e) { sendTelegram("⚠️ Erro: " + e.message); }
}

function extractData() {
  var ss = SpreadsheetApp.openById(XLSX_ID);
  var shR = ss.getSheetByName("RESUMO");
  var hdrs = shR.getRange(1,1,1,12).getValues()[0];
  var vals = shR.getRange(2,1,1,12).getValues()[0];
  var resumo = {};
  for(var i=0;i<hdrs.length;i++) resumo[hdrs[i]] = vals[i];

  var shA = ss.getSheetByName("ARMAZEM");
  var armData = shA.getDataRange().getValues();
  var armazem = [];
  for(var i=1;i<armData.length;i++)
    if(armData[i][0]&&armData[i][1]) armazem.push({nome:armData[i][0],total_sc:armData[i][1]});

  var shD = ss.getSheetByName("DATA DE COLHEITA");
  var diaData = shD.getDataRange().getValues();
  var diario = [];
  for(var i=1;i<diaData.length;i++)
    if(diaData[i][0]&&diaData[i][1]) diario.push({
      data:Utilities.formatDate(diaData[i][0],"America/Sao_Paulo","yyyy-MM-dd"),
      total_colhido:diaData[i][1], area_colhida:diaData[i][2]||0, acumulado:diaData[i][3]||0});

  var shT = ss.getSheetByName("PRODUTIVIDADE");
  var talData = shT.getDataRange().getValues();
  var talhoes = [];
  for(var i=2;i<talData.length;i++)
    if(talData[i][0]&&talData[i][0]!='MÉDIA GERAL'&&talData[i][2])
      talhoes.push({talhao:talData[i][0],area:talData[i][2],total_sc:talData[i][3]||0,
        produtividade:talData[i][4]||0,ha_colhido:talData[i][7]||0,
        pct_colhido:talData[i][8]||0,status:talData[i][9]||'INCOMPLETO'});

  return {
    resumo:{total_colhido:resumo['TOTAL COLHIDO'],area_total:resumo['ÁREA TOTAL'],
      area_colhida:resumo['ÁREA COLHIDA'],pct_colhido:resumo['PERCENTUAL COLHIDO'],
      area_nao_colhida:resumo['ÁREA NÃO COLHIDA'],media_geral:resumo['MÉDIA GERAL'],
      media_umidade:resumo['MÉDIA UMIDADE'],media_impureza:resumo['MÉDIA IMPUREZA'],
      total_desconto:resumo['TOTAL DESCONTO'],desconto_sc_ha:resumo['DESCONTO SC/HÁ']},
    armazem:armazem, diario:diario, talhoes:talhoes
  };
}

function sendTelegram(text) {
  UrlFetchApp.fetch("https://api.telegram.org/bot"+TELEGRAM_BOT_TOKEN+"/sendMessage",{
    method:"post",contentType:"application/json",
    payload:JSON.stringify({chat_id:TELEGRAM_CHAT_ID,text:text,parse_mode:"Markdown"})});
}

function installTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for(var i=0;i<triggers.length;i++) ScriptApp.deleteTrigger(triggers[i]);
  ScriptApp.newTrigger("checkDriveUpdates").timeBased().everyMinutes(30).create();
  sendTelegram("✅ *Sistema ativo!*\n\nRafael salva → Dashboard atualiza automaticamente\n\n👉 https://terra-bi-app.vercel.app");
}

function testar() {
  var data = extractData();
  var r = data.resumo;
  sendTelegram("🧪 *Teste OK!*\n\n• " + r.total_colhido.toFixed(0) + " sc total\n• " +
    (r.pct_colhido*100).toFixed(1) + "% colhido\n• " + r.media_geral.toFixed(2) + " sc/ha");
}
```
