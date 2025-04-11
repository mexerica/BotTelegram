const TELEGRAM_BOT_TOKEN = '8151731373:AAEUNQlTHLmnKNm6qW0_sHY_3nrdxL16PSo';
const TELEGRAM_API_URL = `https://api.telegram.org/bot${"8151731373:AAEUNQlTHLmnKNm6qW0_sHY_3nrdxL16PSo"}/sendMessage`;
const TELEGRAM_CHAT_ID = '7825989522';
const DATA_SHEET_NAME = 'DadosAtuais';
const LAST_READING_SHEET_NAME = 'UltimaLeitura';

function verificarAtualizacoes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  const lastReadingSheet = ss.getSheetByName(LAST_READING_SHEET_NAME);

  if (!dataSheet || !lastReadingSheet) {
    Logger.log('As abas especificadas não foram encontradas.');
    return;
  }

  const data = getDataFromSheet(dataSheet);
  const lastReadingData = getDataFromSheet(lastReadingSheet);
  const changes = compararValores(data, lastReadingData);

  if (changes.length > 0) {
    const mensagem = formatarMensagem(changes);
    enviarMensagemTelegram(mensagem);
    atualizarUltimaLeitura(data, lastReadingSheet);
  } else {
    Logger.log('Nenhuma mudança detectada.');
  }
}

function getDataFromSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const range = sheet.getRange(1, 1, lastRow, lastColumn); // Assume cabeçalho na primeira linha
  const values = range.getValues();
  const data = {};
  for (let i = 1; i < values.length; i++) { // Começa da segunda linha para ignorar o cabeçalho
    const nome = values[i][0];
    const valor = values[i][1];
    if (nome) {
      data[nome] = valor;
    }
  }
  return data;
}

function compararValores(dataAtual, ultimaLeitura) {
  const changes = [];
  for (const nome in dataAtual) {
    if (dataAtual.hasOwnProperty(nome)) {
      if (!ultimaLeitura.hasOwnProperty(nome) || dataAtual[nome] !== ultimaLeitura[nome]) {
        changes.push({ nome: nome, valorAntigo: ultimaLeitura[nome], valorNovo: dataAtual[nome] });
      }
    }
  }
  return changes;
}

function formatarMensagem(changes) {
  let mensagem = 'Cashback Atualizado:\n';
  for (const change of changes) {
    mensagem += `*${change.nome}:* `;
    if (change.valorAntigo !== undefined) {
      mensagem += `${change.valorAntigo} -> `;
    }
    mensagem += `${change.valorNovo} `;
    let dif = (change.valorNovo / change.valorAntigo) * 100;
    if (change.valorNovo < change.valorAntigo) dif -= 100;
    mensagem += `(${dif.toFixed(2)}%)\n`;
  }
  return mensagem;
}

function enviarMensagemTelegram(mensagem) {
  const payload = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      chat_id: TELEGRAM_CHAT_ID,
      text: mensagem,
      parse_mode: 'Markdown' // Para formatar o texto com *negrito*
    })
  };

  try {
    UrlFetchApp.fetch(TELEGRAM_API_URL, payload);
    Logger.log('Mensagem enviada com sucesso para o Telegram!');
  } catch (error) {
    Logger.log(`Erro ao enviar mensagem para o Telegram: ${error}`);
  }
}

function atualizarUltimaLeitura(dataAtual, lastReadingSheet) {
  lastReadingSheet.clearContents(); // Limpa os dados antigos
  let row = 1; // Inicializa a linha para escrever os dados

  const header = ['Nome', 'Valor']; // Adapte conforme o seu cabeçalho
  lastReadingSheet.getRange(1, 1, 1, header.length).setValues([header]);
  row = 2; // Começa a escrever os dados a partir da segunda linha

  for (const nome in dataAtual) {
    if (dataAtual.hasOwnProperty(nome)) {
      lastReadingSheet.getRange(row, 1).setValue(nome);
      lastReadingSheet.getRange(row, 2).setValue(dataAtual[nome]);
      row++;
    }
  }
  Logger.log('Aba da última leitura foi atualizada.');
}

function agendamentoVerificacao() {
  verificarAtualizacoes();
}
