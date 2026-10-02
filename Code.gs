// ---- Login com Google ----
const CLIENT_ID = "47231888741-3kbte3jgrpsgd9bqnndd8bp004gsnsq8.apps.googleusercontent.com"; // TODO: mesmo Client ID usado no index.html
const EMAIL_ADMIN = "carlosmantillaneto@gmail.com"; // TODO: mesmo e-mail usado no index.html

// Verifica o token do Google direto com o servidor do Google e devolve o e-mail
// verificado de quem logou, ou null se o token for inválido/de outro app.
function verificarToken(token) {
  if (!token) return null;
  try {
    const resp = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token),
      { muteHttpExceptions: true }
    );
    if (resp.getResponseCode() !== 200) return null;
    const info = JSON.parse(resp.getContentText());
    if (info.aud !== CLIENT_ID) return null; // token de outro app, nao serve
    if (info.email_verified !== 'true' && info.email_verified !== true) return null;
    return info.email;
  } catch (err) {
    return null;
  }
}

function doGet(e) {
  const emailLogado = verificarToken(e.parameter.token);
  if (!emailLogado) {
    return ContentService.createTextOutput(JSON.stringify({ status: "erro", mensagem: "Faça login novamente." }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rows = sheet.getDataRange().getDisplayValues(); 
  if (rows.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const headers = rows[0].map(h => h.trim());
  const data = [];

  for (let i = 1; i < rows.length; i++) {
    let row = rows[i];
    if (!row[0] && !row[1]) continue; 
    let record = {};
    headers.forEach((header, index) => {
      record[header] = row[index] || '-';
    });
    data.push(record);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "success", data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const dados = JSON.parse(e.postData.contents);

  const emailLogado = verificarToken(dados.token);
  if (!emailLogado) {
    return ContentService.createTextOutput(JSON.stringify({ status: "erro", mensagem: "Faça login novamente." }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (emailLogado.toLowerCase() !== EMAIL_ADMIN.toLowerCase()) {
    return ContentService.createTextOutput(JSON.stringify({ status: "erro", mensagem: "Sem permissão para esta ação." }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  const cabecalhos = [
    "ID", "Responsavel", "Setor", "Hostname",
    "Fabricante", "Modelo", "Serial Number", "Patrimonio", "CPU", "Memoria", "Armazenamento", "Observacao",
    "Monitor 1", "Modelo 1", "Serial 1", "Patrimonio 1", 
    "Monitor 2", "Fabricante 2", "Modelo 2", "Serial 2", "Patrimonio 2", 
    "Nobreak", "Fabricante N", "Serial N", "Patrimonio N"
  ];
  
  const rows = sheet.getDataRange().getValues();
  let headerRow = rows.length > 0 ? rows[0].map(h => h.toString().trim()) : [];
  
  if (rows.length === 0 || headerRow[0] !== "ID") {
    sheet.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    headerRow = cabecalhos;
  }

  const val = (v) => (v && v.toString().trim() !== '' ? v.toString().trim() : '-');
  const getColIndex = (name) => headerRow.indexOf(name) + 1;

  if (dados.acao === 'criar') {
    const novoId = new Date().getTime().toString().slice(-6);
    let novaLinha = new Array(cabecalhos.length).fill('-');
    
    novaLinha[getColIndex("ID") - 1] = novoId;
    novaLinha[getColIndex("Responsavel") - 1] = val(dados.responsavel);
    novaLinha[getColIndex("Setor") - 1] = val(dados.setor);
    novaLinha[getColIndex("Hostname") - 1] = val(dados.hostname);
    novaLinha[getColIndex("Fabricante") - 1] = val(dados.fabricante);
    novaLinha[getColIndex("Modelo") - 1] = val(dados.modelo);
    novaLinha[getColIndex("Serial Number") - 1] = val(dados.serial);
    novaLinha[getColIndex("Patrimonio") - 1] = val(dados.patrimonio);
    novaLinha[getColIndex("CPU") - 1] = val(dados.cpu);
    novaLinha[getColIndex("Memoria") - 1] = val(dados.memoria);
    novaLinha[getColIndex("Armazenamento") - 1] = val(dados.armazenamento);
    novaLinha[getColIndex("Observacao") - 1] = val(dados.observacao);
    
    novaLinha[getColIndex("Monitor 1") - 1] = val(dados.monitor1);
    novaLinha[getColIndex("Modelo 1") - 1] = val(dados.mod1);
    novaLinha[getColIndex("Serial 1") - 1] = val(dados.ser1);
    novaLinha[getColIndex("Patrimonio 1") - 1] = val(dados.pat1);

    novaLinha[getColIndex("Monitor 2") - 1] = val(dados.monitor2);
    novaLinha[getColIndex("Fabricante 2") - 1] = val(dados.fab2);
    novaLinha[getColIndex("Modelo 2") - 1] = val(dados.mod2);
    novaLinha[getColIndex("Serial 2") - 1] = val(dados.ser2);
    novaLinha[getColIndex("Patrimonio 2") - 1] = val(dados.pat2);

    novaLinha[getColIndex("Nobreak") - 1] = val(dados.nobreak);
    novaLinha[getColIndex("Fabricante N") - 1] = val(dados.fabN);
    novaLinha[getColIndex("Serial N") - 1] = val(dados.serN);
    novaLinha[getColIndex("Patrimonio N") - 1] = val(dados.patN);
    
    sheet.appendRow(novaLinha);
  } 
  else if (dados.acao === 'editar') {
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0].toString().trim() === dados.id.toString().trim()) {
        const linha = i + 1;
        sheet.getRange(linha, getColIndex("Responsavel")).setValue(val(dados.responsavel));
        sheet.getRange(linha, getColIndex("Setor")).setValue(val(dados.setor));
        sheet.getRange(linha, getColIndex("Hostname")).setValue(val(dados.hostname));
        sheet.getRange(linha, getColIndex("Fabricante")).setValue(val(dados.fabricante));
        sheet.getRange(linha, getColIndex("Modelo")).setValue(val(dados.modelo));
        sheet.getRange(linha, getColIndex("Serial Number")).setValue(val(dados.serial));
        sheet.getRange(linha, getColIndex("Patrimonio")).setValue(val(dados.patrimonio));
        sheet.getRange(linha, getColIndex("CPU")).setValue(val(dados.cpu));
        sheet.getRange(linha, getColIndex("Memoria")).setValue(val(dados.memoria));
        sheet.getRange(linha, getColIndex("Armazenamento")).setValue(val(dados.armazenamento));
        sheet.getRange(linha, getColIndex("Observacao")).setValue(val(dados.observacao));
        
        sheet.getRange(linha, getColIndex("Monitor 1")).setValue(val(dados.monitor1));
        sheet.getRange(linha, getColIndex("Modelo 1")).setValue(val(dados.mod1));
        sheet.getRange(linha, getColIndex("Serial 1")).setValue(val(dados.ser1));
        sheet.getRange(linha, getColIndex("Patrimonio 1")).setValue(val(dados.pat1));

        sheet.getRange(linha, getColIndex("Monitor 2")).setValue(val(dados.monitor2));
        sheet.getRange(linha, getColIndex("Fabricante 2")).setValue(val(dados.fab2));
        sheet.getRange(linha, getColIndex("Modelo 2")).setValue(val(dados.mod2));
        sheet.getRange(linha, getColIndex("Serial 2")).setValue(val(dados.ser2));
        sheet.getRange(linha, getColIndex("Patrimonio 2")).setValue(val(dados.pat2));

        sheet.getRange(linha, getColIndex("Nobreak")).setValue(val(dados.nobreak));
        sheet.getRange(linha, getColIndex("Fabricante N")).setValue(val(dados.fabN));
        sheet.getRange(linha, getColIndex("Serial N")).setValue(val(dados.serN));
        sheet.getRange(linha, getColIndex("Patrimonio N")).setValue(val(dados.patN));
        break;
      }
    }
  }
  else if (dados.acao === 'deletar') {
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0].toString().trim() === dados.id.toString().trim()) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "sucesso" }))
    .setMimeType(ContentService.MimeType.JSON);
}
