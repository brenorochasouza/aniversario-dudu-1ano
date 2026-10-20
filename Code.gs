/**
 * Confirmação de presença — aniversário "Minha primeira volta ao mundo"
 * 1) Crie uma planilha Google Sheets.
 * 2) Extensões > Apps Script; substitua o conteúdo por este código.
 * 3) Execute setup() uma vez e autorize.
 * 4) Implantar > Nova implantação > Aplicativo da Web.
 *    Executar como: você | Quem tem acesso: qualquer pessoa.
 * 5) Copie a URL /exec para a constante SCRIPT_URL do index.html.
 */
const SHEET_NAME = 'Confirmações';

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Data/hora do envio','Nome completo','WhatsApp','Adultos','Crianças','Idades das crianças','Presença','Recadinho']);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,8).setFontWeight('bold').setBackground('#123f66').setFontColor('#ffffff');
    sheet.autoResizeColumns(1,8);
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const nome = String(data.nome || '').trim().slice(0,120);
    const presenca = String(data.presenca || '');
    if (!nome || !['Sim','Não'].includes(presenca)) {
      return ContentService.createTextOutput(JSON.stringify({ok:false,message:'Dados inválidos'})).setMimeType(ContentService.MimeType.JSON);
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) { setup(); sheet = ss.getSheetByName(SHEET_NAME); }
    sheet.appendRow([
      new Date(), nome, String(data.telefone || '').slice(0,30),
      Math.max(0, Math.min(20, Number(data.adultos) || 0)),
      Math.max(0, Math.min(20, Number(data.criancas) || 0)),
      String(data.idades || '').slice(0,80), presenca, String(data.recado || '').slice(0,500)
    ]);
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,message:'Erro ao registrar'})).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
