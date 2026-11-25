/** @OnlyCurrentDoc */
/**
 * CRM Eleve: cadastro e acompanhamento de agendamentos da CRC.
 * Os dados ficam na aba "Produtividade" desta planilha; o CRM só lê e grava nela.
 *
 * Implantação: Implantar → Nova implantação → App da Web
 *   Executar como: "Usuário que acessa o app da Web"
 *   Quem pode acessar: "Qualquer pessoa com uma Conta do Google"
 * Só quem tem permissão de EDIÇÃO na planilha consegue usar o CRM.
 */

const SHEET_NAME = 'Produtividade';
const NCOLS = 16; // A..P

// Colunas (1 = A)
const C = {
  dataAgend: 1, nome: 2, telefone: 3, campanha: 4, crc: 5, nascimento: 6, clinica: 7,
  dataConsulta: 8, hora: 9, confirmacao: 10, meio: 11, status: 12, fechou: 13,
  orcamento: 14, valor: 15, obs: 16,
};

const FALLBACK = {
  campanha: ['Paciente', 'Indicação', 'Aniversariantes', 'Base face', 'Lead novo', 'Facebook', 'Implante', 'Protocolo', 'Prótese flex', 'INSS', 'Limpeza', 'Voucher 250', 'Voucher 500', 'Campanha'],
  crc: ['Thaís'],
  clinica: ['Eleve', 'Penha', 'Grimaldi', 'Teotônio', 'Vila Diva', 'Pira', 'Star'],
  meio: ['WhatsApp', 'Ligação', 'Callix', '3C'],
  status: ['CONFIRMOU', 'COMPARECEU', 'FALTOU', 'DESMARCOU', 'REAGENDADO', 'CANCELOU', 'CANCELOU/ NÃO RETORNAR'],
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('CRM Eleve')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function sheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Aba "' + SHEET_NAME + '" não encontrada.');
  return sh;
}

function tz_() {
  return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'America/Sao_Paulo';
}

function fmtDate_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  const s = String(v || '').trim();
  const m = s.match(/^(\d{1,2})\/+(\d{1,2})\/+(\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}

function parseDate_(s) {
  if (!s) return '';
  return Utilities.parseDate(s + ' 12:00', tz_(), 'yyyy-MM-dd HH:mm');
}

function fmtHora_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz_(), "HH'h'mm");
  return String(v || '').trim();
}

function str_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return String(Math.round(v) === v ? v : v);
  return String(v).trim();
}

function rowToRecord_(r, rowNumber) {
  return {
    row: rowNumber,
    dataAgend: fmtDate_(r[C.dataAgend - 1]),
    nome: str_(r[C.nome - 1]),
    telefone: str_(r[C.telefone - 1]),
    campanha: str_(r[C.campanha - 1]),
    crc: str_(r[C.crc - 1]),
    nascimento: fmtDate_(r[C.nascimento - 1]),
    clinica: str_(r[C.clinica - 1]),
    dataConsulta: fmtDate_(r[C.dataConsulta - 1]),
    hora: fmtHora_(r[C.hora - 1]),
    confirmacao: str_(r[C.confirmacao - 1]),
    meio: str_(r[C.meio - 1]),
    status: str_(r[C.status - 1]),
    fechou: str_(r[C.fechou - 1]),
    orcamento: str_(r[C.orcamento - 1]),
    valor: str_(r[C.valor - 1]),
    obs: str_(r[C.obs - 1]),
  };
}

function isRecord_(r) {
  return str_(r[C.nome - 1]) !== '' || r[C.dataAgend - 1] instanceof Date;
}

/** Opções das listas suspensas da planilha (com cache de 6 h). */
function options_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('options_v1');
  if (hit) return JSON.parse(hit);

  const sh = sheet_();
  const last = Math.max(2, sh.getLastRow());
  const read = (col) => {
    const out = [];
    const seen = {};
    const add = (v) => {
      const s = String(v || '').trim();
      const k = s.toLowerCase();
      if (!s || seen[k] || /^\d$/.test(s) || /^(meio de agendamento|consultor \(crc\))$/i.test(s)) return;
      seen[k] = true;
      out.push(s);
    };
    const dvs = sh.getRange(2, col, last - 1, 1).getDataValidations();
    const rulesSeen = {};
    dvs.forEach((row) => {
      const dv = row[0];
      if (!dv || dv.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return;
      const vals = dv.getCriteriaValues()[0] || [];
      const key = vals.join('|');
      if (rulesSeen[key]) return;
      rulesSeen[key] = true;
      vals.forEach(add);
    });
    sh.getRange(2, col, last - 1, 1).getValues().forEach((r) => add(r[0]));
    return out;
  };
  const opts = {
    campanha: read(C.campanha),
    crc: read(C.crc),
    clinica: read(C.clinica),
    meio: read(C.meio),
    status: read(C.status),
  };
  Object.keys(FALLBACK).forEach((k) => { if (!opts[k].length) opts[k] = FALLBACK[k]; });
  cache.put('options_v1', JSON.stringify(opts), 21600);
  return opts;
}

/** Carga inicial do app. */
function bootstrap() {
  const sh = sheet_();
  const last = sh.getLastRow();
  const records = [];
  if (last >= 2) {
    const values = sh.getRange(2, 1, last - 1, NCOLS).getValues();
    values.forEach((r, i) => { if (isRecord_(r)) records.push(rowToRecord_(r, i + 2)); });
  }
  const props = PropertiesService.getUserProperties();
  return {
    email: Session.getActiveUser().getEmail(),
    meuNome: props.getProperty('crcNome') || '',
    options: options_(),
    records: records,
    today: Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'),
  };
}

function setMeuNome(nome) {
  PropertiesService.getUserProperties().setProperty('crcNome', String(nome || '').trim());
  return true;
}

/**
 * Cria (rec.row vazio) ou atualiza um agendamento.
 * rec.orig = nome original da linha, usado para detectar se a planilha foi reordenada.
 */
function saveRecord(rec) {
  if (!rec || !String(rec.nome || '').trim()) throw new Error('Informe o nome do paciente.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    let row = Number(rec.row) || 0;
    let current = null;

    if (row) {
      current = sh.getRange(row, 1, 1, NCOLS).getValues()[0];
      if (str_(current[C.nome - 1]) !== String(rec.orig || '').trim()) {
        throw new Error('Essa linha foi alterada na planilha por outra pessoa. Recarregue o CRM e tente de novo.');
      }
    } else {
      row = nextEmptyRow_(sh);
      current = new Array(NCOLS).fill('');
    }

    const out = current.slice();
    out[C.dataAgend - 1] = rec.dataAgend ? parseDate_(rec.dataAgend) : (current[C.dataAgend - 1] || parseDate_(Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd')));
    out[C.nome - 1] = String(rec.nome).trim();
    out[C.telefone - 1] = String(rec.telefone || '').trim();
    out[C.campanha - 1] = String(rec.campanha || '').trim();
    out[C.crc - 1] = String(rec.crc || '').trim();
    out[C.nascimento - 1] = dateOrKeep_(rec.nascimento, current[C.nascimento - 1]);
    out[C.clinica - 1] = String(rec.clinica || '').trim();
    out[C.dataConsulta - 1] = dateOrKeep_(rec.dataConsulta, current[C.dataConsulta - 1]);
    out[C.hora - 1] = String(rec.hora || '').trim();
    out[C.confirmacao - 1] = String(rec.confirmacao || '').trim();
    out[C.meio - 1] = String(rec.meio || '').trim();
    out[C.status - 1] = String(rec.status || '').trim();
    out[C.fechou - 1] = String(rec.fechou || '').trim();
    out[C.orcamento - 1] = String(rec.orcamento || '').trim();
    out[C.valor - 1] = money_(rec.valor);
    out[C.obs - 1] = String(rec.obs || '').trim();

    // Telefone e hora como texto, para o Sheets não converter em número/horário.
    sh.getRange(row, C.telefone).setNumberFormat('@');
    sh.getRange(row, C.hora).setNumberFormat('@');
    sh.getRange(row, C.dataAgend).setNumberFormat('dd/mm/yyyy');
    sh.getRange(row, C.nascimento).setNumberFormat('dd/mm/yyyy');
    sh.getRange(row, C.dataConsulta).setNumberFormat('dd/mm/yyyy');
    sh.getRange(row, 1, 1, NCOLS).setValues([out]);
    SpreadsheetApp.flush();

    return rowToRecord_(sh.getRange(row, 1, 1, NCOLS).getValues()[0], row);
  } finally {
    lock.releaseLock();
  }
}

/** Data nova; se veio vazia e a célula tinha um texto fora do padrão, preserva o texto. */
function dateOrKeep_(iso, cur) {
  if (iso) return parseDate_(iso);
  if (cur !== '' && cur !== null && !fmtDate_(cur)) return cur;
  return '';
}

/** Número (aceita 1.500,00 / R$ 1500); texto não reconhecido é mantido como está. */
function money_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (typeof v === 'number') return v;
  let s = String(v).replace(/[R$\s]/g, '');
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  else if (/^\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  const n = Number(s);
  return isFinite(n) && s !== '' ? n : String(v);
}

/** Primeira linha totalmente vazia (A..P) abaixo do último agendamento. */
function nextEmptyRow_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return 2;
  const values = sh.getRange(2, 1, last - 1, NCOLS).getValues();
  let lastRecord = 1;
  values.forEach((r, i) => { if (isRecord_(r)) lastRecord = i + 2; });
  for (let i = lastRecord - 1; i < values.length; i++) {
    if (values[i].every((v) => v === '' || v === null)) return i + 2;
  }
  return last + 1;
}

/** Limpa o conteúdo da linha (A..P). Formatação e listas suspensas são mantidas. */
function deleteRecord(row, orig) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const current = sh.getRange(row, 1, 1, NCOLS).getValues()[0];
    if (str_(current[C.nome - 1]) !== String(orig || '').trim()) {
      throw new Error('Essa linha foi alterada na planilha por outra pessoa. Recarregue o CRM.');
    }
    sh.getRange(row, 1, 1, NCOLS).clearContent();
    return true;
  } finally {
    lock.releaseLock();
  }
}
