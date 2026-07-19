/**
 * BACKEND — FESTA TUDO DE HELENA
 * Vincule este código a uma Planilha Google e execute configurarProjeto() uma vez.
 */

const APP_SOURCE = 'tudodehelena-rsvp';
const CONFIG_SHEET = 'CONFIGURACAO';
const RESPONSES_SHEET = 'CONFIRMACOES';
const SPREADSHEET_ID_PROPERTY = 'SPREADSHEET_ID';
const ADMIN_PIN_PROPERTY = 'ADMIN_PIN';

const RESPONSE_HEADERS = [
  'Protocolo',
  'Criado em',
  'Atualizado em',
  'Responsável',
  'WhatsApp',
  'Presença',
  'Quantidade de crianças',
  'Crianças e idades',
  'Restrições / alergias',
  'Observações'
];

const CONFIG_DEFAULTS = {
  eventTitle: 'Festa Tudo de Helena',
  child1Name: 'Heloá',
  child1Age: '',
  child2Name: 'Heitor',
  child2Age: '',
  eventDate: '',
  eventTime: '',
  eventEndTime: '',
  locationName: '',
  address: '',
  eventNote: 'Os detalhes da festa serão informados em breve.',
  deadline: '',
  inviteRule: 'A festa é exclusiva para crianças. Informe apenas as crianças que receberam o convite — não é necessário colocar quantidade de adultos.',
  maxChildrenPerResponse: '10'
};

const CONFIG_DESCRIPTIONS = {
  eventTitle: 'Título exibido no topo do convite',
  child1Name: 'Nome da primeira aniversariante',
  child1Age: 'Idade da primeira aniversariante',
  child2Name: 'Nome do segundo aniversariante',
  child2Age: 'Idade do segundo aniversariante',
  eventDate: 'Data no formato AAAA-MM-DD',
  eventTime: 'Horário de início',
  eventEndTime: 'Horário de término',
  locationName: 'Nome do salão ou local',
  address: 'Endereço da festa',
  eventNote: 'Mensagem exibida abaixo dos detalhes',
  deadline: 'Prazo de confirmação no formato AAAA-MM-DD',
  inviteRule: 'Aviso sobre crianças convidadas e adultos',
  maxChildrenPerResponse: 'Limite de crianças permitido em uma confirmação'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Festa Tudo de Helena')
    .addItem('Preparar / reparar planilha', 'configurarProjeto')
    .addToUi();
}

/**
 * Execute esta função uma vez pelo editor do Apps Script.
 * Ela liga o script à planilha, cria as abas e define o código inicial 1234.
 */
function configurarProjeto() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Abra este script por Extensões > Apps Script dentro da Planilha Google.');
  }

  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(SPREADSHEET_ID_PROPERTY, spreadsheet.getId());
  spreadsheet.setSpreadsheetTimeZone('America/Sao_Paulo');
  if (!properties.getProperty(ADMIN_PIN_PROPERTY)) {
    properties.setProperty(ADMIN_PIN_PROPERTY, '1234');
  }

  ensureSheets_(spreadsheet, true);
  try {
    spreadsheet.toast('Planilha preparada. Código inicial do painel: 1234', 'Festa Tudo de Helena', 8);
    SpreadsheetApp.getUi().alert(
      'Configuração concluída',
      'As abas CONFIGURACAO e CONFIRMACOES foram preparadas.\n\nCódigo inicial do painel: 1234\nTroque-o no primeiro acesso ao organizador.html.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (uiError) {
    console.log('Configuração concluída. Código inicial do painel: 1234');
  }
}

function doGet(e) {
  try {
    const action = String(e && e.parameter && e.parameter.action || 'status');
    if (action === 'getPublicConfig') {
      const result = { ok: true, config: getPublicConfig_() };
      const callback = String(e.parameter.callback || '');
      if (callback) return jsonpOutput_(callback, result);
      return jsonOutput_(result);
    }
    return jsonOutput_({ ok: true, app: 'Festa Tudo de Helena', status: 'online' });
  } catch (error) {
    return jsonOutput_({ ok: false, message: publicErrorMessage_(error) });
  }
}

function doPost(e) {
  const requestId = sanitizeRequestId_(e && e.parameter && e.parameter.requestId);
  let result;
  try {
    const action = String(e && e.parameter && e.parameter.action || '');
    const payload = parsePayload_(e && e.parameter && e.parameter.payload);

    switch (action) {
      case 'submitRsvp':
        result = submitRsvp_(payload);
        break;
      case 'getAdminData':
        result = getAdminData_(payload);
        break;
      case 'saveConfig':
        result = saveConfig_(payload);
        break;
      default:
        throw new Error('Ação não reconhecida.');
    }
  } catch (error) {
    console.error(error && error.stack ? error.stack : error);
    result = { ok: false, message: publicErrorMessage_(error) };
  }
  return iframeMessageOutput_(requestId, result);
}

function submitRsvp_(payload) {
  if (String(payload.website || '').trim()) {
    throw new Error('Envio recusado.');
  }

  const elapsed = Date.now() - Number(payload.startedAt || 0);
  if (elapsed > 0 && elapsed < 700) {
    throw new Error('Aguarde um instante e envie novamente.');
  }

  const config = getPublicConfig_();
  const going = String(payload.going || '').toLowerCase();
  const responsibleName = validateText_(payload.responsibleName, 'Nome do responsável', 2, 100);
  const phone = normalizePhone_(payload.phone);
  const restrictions = sanitizeCell_(limitText_(payload.restrictions, 500));
  const notes = sanitizeCell_(limitText_(payload.notes, 500));

  if (going !== 'sim' && going !== 'nao') {
    throw new Error('Informe se a criança vai ou não à festa.');
  }
  if (phone.length < 10 || phone.length > 13) {
    throw new Error('Informe um WhatsApp válido com DDD.');
  }

  let childrenCount = 0;
  let childrenText = '';
  if (going === 'sim') {
    const children = Array.isArray(payload.children) ? payload.children : [];
    const max = clamp_(Number(config.maxChildrenPerResponse) || 10, 1, 20);
    if (children.length < 1 || children.length > max) {
      throw new Error('Quantidade de crianças inválida.');
    }

    const normalizedChildren = children.map(function(child, index) {
      const name = validateText_(child && child.name, 'Nome da criança ' + (index + 1), 2, 100);
      const rawAge = String(child && child.age || '').trim();
      let age = '';
      if (rawAge !== '') {
        const ageNumber = Number(rawAge);
        if (!Number.isInteger(ageNumber) || ageNumber < 0 || ageNumber > 17) {
          throw new Error('Idade inválida para a criança ' + (index + 1) + '.');
        }
        age = String(ageNumber);
      }
      return { name: sanitizeCell_(name), age: age };
    });

    childrenCount = normalizedChildren.length;
    childrenText = normalizedChildren.map(function(child) {
      if (!child.age) return child.name;
      return child.name + ' (' + child.age + (child.age === '1' ? ' ano)' : ' anos)');
    }).join(' | ');
  } else {
    childrenText = sanitizeCell_(validateText_(payload.declinedChildName, 'Nome da criança convidada', 2, 160));
  }

  const spreadsheet = getSpreadsheet_();
  ensureSheets_(spreadsheet);
  const sheet = spreadsheet.getSheetByName(RESPONSES_SHEET);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const now = new Date();
    const lastRow = sheet.getLastRow();
    let existingRow = 0;
    let protocol = '';
    let createdAt = now;

    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, RESPONSE_HEADERS.length).getValues();
      for (let i = data.length - 1; i >= 0; i -= 1) {
        if (normalizePhone_(data[i][4]) === phone) {
          existingRow = i + 2;
          protocol = String(data[i][0] || '');
          createdAt = data[i][1] instanceof Date ? data[i][1] : now;
          break;
        }
      }
    }

    if (!protocol) protocol = createProtocol_();
    const row = [
      protocol,
      createdAt,
      now,
      sanitizeCell_(responsibleName),
      phone,
      going === 'sim' ? 'Sim' : 'Não',
      childrenCount,
      childrenText,
      restrictions,
      notes
    ];

    if (existingRow) {
      sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return {
      ok: true,
      protocol: protocol,
      updated: Boolean(existingRow),
      childrenCount: childrenCount
    };
  } finally {
    lock.releaseLock();
  }
}

function getAdminData_(payload) {
  assertAdminPin_(payload && payload.pin);
  const spreadsheet = getSpreadsheet_();
  ensureSheets_(spreadsheet);
  const sheet = spreadsheet.getSheetByName(RESPONSES_SHEET);
  const lastRow = sheet.getLastRow();
  const responses = [];

  if (lastRow >= 2) {
    const range = sheet.getRange(2, 1, lastRow - 1, RESPONSE_HEADERS.length);
    const values = range.getValues();
    const displayValues = range.getDisplayValues();

    for (let i = values.length - 1; i >= 0; i -= 1) {
      responses.push({
        protocol: String(values[i][0] || ''),
        createdAt: displayValues[i][1] || '',
        updatedAt: displayValues[i][2] || '',
        responsibleName: String(values[i][3] || ''),
        phone: String(values[i][4] || ''),
        presence: String(values[i][5] || ''),
        childrenCount: Number(values[i][6] || 0),
        children: String(values[i][7] || ''),
        restrictions: String(values[i][8] || ''),
        notes: String(values[i][9] || '')
      });
    }
  }

  const stats = responses.reduce(function(acc, row) {
    acc.total += 1;
    if (row.presence === 'Sim') {
      acc.confirmedFamilies += 1;
      acc.confirmedChildren += Number(row.childrenCount || 0);
    } else if (row.presence === 'Não') {
      acc.declined += 1;
    }
    return acc;
  }, { total: 0, confirmedFamilies: 0, confirmedChildren: 0, declined: 0 });

  return {
    ok: true,
    config: getPublicConfig_(),
    stats: stats,
    responses: responses.slice(0, 1000)
  };
}

function saveConfig_(payload) {
  assertAdminPin_(payload && payload.pin);
  const incoming = payload && payload.config || {};
  const clean = {
    eventTitle: validateText_(incoming.eventTitle, 'Título da festa', 2, 100),
    child1Name: validateText_(incoming.child1Name, 'Nome da primeira aniversariante', 1, 50),
    child1Age: validateOptionalInteger_(incoming.child1Age, 0, 99, 'Idade da primeira aniversariante'),
    child2Name: validateText_(incoming.child2Name, 'Nome do segundo aniversariante', 1, 50),
    child2Age: validateOptionalInteger_(incoming.child2Age, 0, 99, 'Idade do segundo aniversariante'),
    eventDate: validateOptionalDate_(incoming.eventDate, 'Data da festa'),
    eventTime: validateOptionalTime_(incoming.eventTime, 'Horário de início'),
    eventEndTime: validateOptionalTime_(incoming.eventEndTime, 'Horário de término'),
    locationName: limitText_(incoming.locationName, 100),
    address: limitText_(incoming.address, 180),
    eventNote: limitText_(incoming.eventNote, 500),
    deadline: validateOptionalDate_(incoming.deadline, 'Prazo de confirmação'),
    inviteRule: limitText_(incoming.inviteRule, 500),
    maxChildrenPerResponse: String(clamp_(Number(incoming.maxChildrenPerResponse) || 10, 1, 20))
  };

  const spreadsheet = getSpreadsheet_();
  ensureSheets_(spreadsheet);
  writeConfig_(spreadsheet, clean);

  const newPin = String(payload && payload.newPin || '').trim();
  if (newPin) {
    if (newPin.length < 4 || newPin.length > 30) {
      throw new Error('O novo código deve ter entre 4 e 30 caracteres.');
    }
    PropertiesService.getScriptProperties().setProperty(ADMIN_PIN_PROPERTY, newPin);
  }

  return { ok: true, config: getPublicConfig_(), pinChanged: Boolean(newPin) };
}

function getPublicConfig_() {
  const spreadsheet = getSpreadsheet_();
  ensureSheets_(spreadsheet);
  const sheet = spreadsheet.getSheetByName(CONFIG_SHEET);
  const lastRow = sheet.getLastRow();
  const config = Object.assign({}, CONFIG_DEFAULTS);

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
    values.forEach(function(row) {
      const key = String(row[0] || '').trim();
      if (Object.prototype.hasOwnProperty.call(CONFIG_DEFAULTS, key)) {
        config[key] = String(row[1] || '').trim();
      }
    });
  }

  config.maxChildrenPerResponse = clamp_(Number(config.maxChildrenPerResponse) || 10, 1, 20);
  return config;
}

function ensureSheets_(spreadsheet, forceRepair) {
  let configSheet = spreadsheet.getSheetByName(CONFIG_SHEET);
  const configWasCreated = !configSheet;
  if (!configSheet) configSheet = spreadsheet.insertSheet(CONFIG_SHEET);

  let responseSheet = spreadsheet.getSheetByName(RESPONSES_SHEET);
  const responsesWereCreated = !responseSheet;
  if (!responseSheet) responseSheet = spreadsheet.insertSheet(RESPONSES_SHEET);

  const configNeedsRepair = configWasCreated || forceRepair || configSheet.getLastRow() === 0 || String(configSheet.getRange(1, 1).getDisplayValues()[0][0]) !== 'Chave';
  const responseHeaders = responseSheet.getLastRow() >= 1
    ? responseSheet.getRange(1, 1, 1, RESPONSE_HEADERS.length).getDisplayValues()[0]
    : [];
  const responsesNeedRepair = responsesWereCreated || forceRepair || responseHeaders.join('|') !== RESPONSE_HEADERS.join('|');

  if (configNeedsRepair) prepareConfigSheet_(configSheet);
  if (responsesNeedRepair) prepareResponsesSheet_(responseSheet);
}

function prepareConfigSheet_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([['Chave', 'Valor', 'Descrição']]);
  }

  const existing = {};
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues().forEach(function(row) {
      existing[String(row[0] || '')] = String(row[1] || '');
    });
  }

  const rows = Object.keys(CONFIG_DEFAULTS).map(function(key) {
    return [key, Object.prototype.hasOwnProperty.call(existing, key) ? existing[key] : CONFIG_DEFAULTS[key], CONFIG_DESCRIPTIONS[key] || ''];
  });
  sheet.getRange(2, 2, rows.length, 1).setNumberFormat('@');
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#7c4d9e').setFontColor('#ffffff');
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 330);
  sheet.setColumnWidth(3, 420);
  sheet.getRange(2, 1, rows.length, 3).setVerticalAlignment('top').setWrap(true);
}

function prepareResponsesSheet_(sheet) {
  const currentHeaders = sheet.getLastRow() >= 1
    ? sheet.getRange(1, 1, 1, RESPONSE_HEADERS.length).getDisplayValues()[0]
    : [];
  if (currentHeaders.join('|') !== RESPONSE_HEADERS.join('|')) {
    sheet.getRange(1, 1, 1, RESPONSE_HEADERS.length).setValues([RESPONSE_HEADERS]);
  }
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, RESPONSE_HEADERS.length).setFontWeight('bold').setBackground('#e85d8a').setFontColor('#ffffff');
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 155);
  sheet.setColumnWidth(3, 155);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 145);
  sheet.setColumnWidth(6, 95);
  sheet.setColumnWidth(7, 155);
  sheet.setColumnWidth(8, 360);
  sheet.setColumnWidth(9, 300);
  sheet.setColumnWidth(10, 300);
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 2, sheet.getMaxRows() - 1, 2).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  }
}

function writeConfig_(spreadsheet, config) {
  const sheet = spreadsheet.getSheetByName(CONFIG_SHEET);
  const rows = Object.keys(CONFIG_DEFAULTS).map(function(key) {
    return [key, String(config[key] == null ? '' : config[key]), CONFIG_DESCRIPTIONS[key] || ''];
  });
  sheet.getRange(2, 2, rows.length, 1).setNumberFormat('@');
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
}

function getSpreadsheet_() {
  const properties = PropertiesService.getScriptProperties();
  let id = properties.getProperty(SPREADSHEET_ID_PROPERTY);
  if (!id) {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      id = active.getId();
      properties.setProperty(SPREADSHEET_ID_PROPERTY, id);
    }
  }
  if (!id) {
    throw new Error('O projeto ainda não foi configurado. Execute configurarProjeto() na Planilha Google.');
  }
  return SpreadsheetApp.openById(id);
}

function assertAdminPin_(pin) {
  const saved = PropertiesService.getScriptProperties().getProperty(ADMIN_PIN_PROPERTY) || '1234';
  if (String(pin || '').trim() !== saved) {
    throw new Error('Código administrativo incorreto.');
  }
}

function normalizePhone_(value) {
  return String(value == null ? '' : value).replace(/\D/g, '');
}

function validateText_(value, label, min, max) {
  const text = limitText_(value, max);
  if (text.length < min) throw new Error(label + ' não foi preenchido corretamente.');
  return text;
}

function limitText_(value, max) {
  return String(value == null ? '' : value).replace(/\s+/g, ' ').trim().slice(0, max);
}

function validateOptionalInteger_(value, min, max, label) {
  const text = String(value == null ? '' : value).trim();
  if (!text) return '';
  const number = Number(text);
  if (!Number.isInteger(number) || number < min || number > max) throw new Error(label + ' inválida.');
  return String(number);
}

function validateOptionalDate_(value, label) {
  const text = String(value == null ? '' : value).trim();
  if (!text) return '';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) throw new Error(label + ' inválida.');
  const parts = text.split('-').map(Number);
  const date = new Date(parts[0], parts[1] - 1, parts[2]);
  if (date.getFullYear() !== parts[0] || date.getMonth() !== parts[1] - 1 || date.getDate() !== parts[2]) {
    throw new Error(label + ' inválida.');
  }
  return text;
}

function validateOptionalTime_(value, label) {
  const text = String(value == null ? '' : value).trim();
  if (!text) return '';
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(text)) throw new Error(label + ' inválido.');
  return text;
}

function clamp_(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function sanitizeCell_(value) {
  const text = String(value == null ? '' : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function createProtocol_() {
  return 'TDH-' + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
}

function parsePayload_(payloadText) {
  if (!payloadText) return {};
  try {
    return JSON.parse(payloadText);
  } catch (error) {
    throw new Error('Dados enviados em formato inválido.');
  }
}

function sanitizeRequestId_(value) {
  const id = String(value || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 100);
  return id || 'sem_id';
}

function publicErrorMessage_(error) {
  const message = String(error && error.message || 'Ocorreu um erro inesperado.');
  return message.replace(/Exception:\s*/i, '').slice(0, 300);
}

function jsonOutput_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonpOutput_(callback, data) {
  const safeCallback = String(callback || '');
  if (!/^[A-Za-z_$][0-9A-Za-z_$\.]{0,120}$/.test(safeCallback)) {
    return jsonOutput_({ ok: false, message: 'Callback inválido.' });
  }
  return ContentService
    .createTextOutput(safeCallback + '(' + JSON.stringify(data).replace(/</g, '\\u003c') + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function iframeMessageOutput_(requestId, payload) {
  const message = {
    source: APP_SOURCE,
    requestId: requestId,
    payload: payload
  };
  const safeJson = JSON.stringify(message)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');

  return HtmlService
    .createHtmlOutput(
      '<!doctype html><html><head><meta charset="utf-8"></head><body>' +
      '<script>window.parent.postMessage(' + safeJson + ', "*");<\/script>' +
      '<noscript>Resposta processada.</noscript></body></html>'
    )
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
