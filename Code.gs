const APP = {
  SHEETS: {
    CONFIG: 'CONFIG',
    LISTAS: 'LISTAS',
    INDICADORES: 'INDICADORES',
    HISTORICO: 'HISTORICO',
    PLANO_ACAO: 'PLANO_ACAO'
  }
};

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Gestão de Indicadores')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const initialized = isSystemInitialized_(ss);

  if (!initialized) {
    return {
      initialized: false,
      indicators: [],
      recentHistory: [],
      recentActions: [],
      availableYears: [new Date().getFullYear()],
      counts: { indicadores: 0, historico: 0, acoes: 0 }
    };
  }

  const indicators = getIndicators_();
  const history = getHistory_();
  const actions = getActionPlans_();
  const indicatorMap = Object.fromEntries(indicators.map(i => [String(i.id), i.nome]));

  const recentHistory = history
    .sort((a, b) => sortYearMonthDesc_(a, b))
    .slice(0, 12)
    .map(row => Object.assign({}, row, { indicadorNome: indicatorMap[String(row.indicadorId)] || '' }));

  const recentActions = actions
    .sort((a, b) => sortYearMonthDesc_(a, b))
    .slice(0, 12)
    .map(row => Object.assign({}, row, { indicadorNome: indicatorMap[String(row.indicadorId)] || '' }));

  const availableYears = [...new Set(history.map(h => Number(h.ano)).filter(v => !isNaN(v)))]
    .sort((a, b) => b - a);

  if (!availableYears.length) availableYears.push(new Date().getFullYear());

  return {
    initialized: true,
    indicators,
    recentHistory,
    recentActions,
    availableYears,
    counts: {
      indicadores: indicators.filter(i => i.ativo).length,
      historico: history.length,
      acoes: actions.length
    }
  };
}

function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createConfigSheet_(ss);
  createListsSheet_(ss);
  createIndicatorsSheet_(ss);
  createHistorySheet_(ss);
  createActionPlanSheet_(ss);

  setConfigValue_('APP_INITIALIZED', 'TRUE');
  setConfigValue_('NEXT_INDICATOR_ID', '1');
  setConfigValue_('NEXT_HISTORY_ID', '1');
  setConfigValue_('NEXT_ACTION_ID', '1');
  setConfigValue_('VERSION', '1.0');

  return 'Estrutura criada com sucesso. As abas CONFIG, LISTAS, INDICADORES, HISTORICO e PLANO_ACAO já estão prontas.';
}

function saveIndicator(payload) {
  ensureSystemReady_();

  const nome = sanitizeText_(payload.nome);
  const metaValor = toNumber_(payload.metaValor);
  const polaridade = sanitizeText_(payload.polaridade) || 'MENOR_MELHOR';
  const responsavel = sanitizeText_(payload.responsavel);
  const metaDescricao = sanitizeText_(payload.metaDescricao);
  const descricao = sanitizeText_(payload.descricao);

  if (!nome) throw new Error('Informe o nome do indicador.');
  if (metaValor === null) throw new Error('Informe uma meta numérica válida.');

  const sh = getSheet_(APP.SHEETS.INDICADORES);
  const id = getNextId_('NEXT_INDICATOR_ID');
  sh.appendRow([
    id,
    nome,
    metaValor,
    polaridade,
    responsavel,
    metaDescricao,
    descricao,
    'TRUE',
    new Date(),
    new Date()
  ]);

  return 'Indicador salvo com sucesso.';
}

function saveMonthlyRecord(payload) {
  ensureSystemReady_();

  const indicadorId = sanitizeText_(payload.indicadorId);
  const ano = toInt_(payload.ano);
  const mes = toInt_(payload.mes);
  const numerador = toNumber_(payload.numerador);
  const denominador = toNumber_(payload.denominador);
  const observacao = sanitizeText_(payload.observacao);

  if (!indicadorId) throw new Error('Selecione um indicador.');
  if (!ano || ano < 2000 || ano > 2100) throw new Error('Ano inválido.');
  if (!mes || mes < 1 || mes > 12) throw new Error('Mês inválido.');
  if (numerador === null) throw new Error('Numerador inválido.');
  if (denominador === null || denominador === 0) throw new Error('Denominador inválido. Não pode ser zero.');

  const result = numerador / denominador;
  const sh = getSheet_(APP.SHEETS.HISTORICO);
  const data = sh.getDataRange().getValues();

  let foundRow = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(indicadorId) && Number(data[i][2]) === ano && Number(data[i][3]) === mes) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow) {
    sh.getRange(foundRow, 5, 1, 6).setValues([[
      numerador,
      denominador,
      result,
      '',
      observacao,
      new Date()
    ]]);
  } else {
    const id = getNextId_('NEXT_HISTORY_ID');
    sh.appendRow([
      id,
      indicadorId,
      ano,
      mes,
      numerador,
      denominador,
      result,
      '',
      observacao,
      new Date(),
      new Date()
    ]);
  }

  recomputeMedianForIndicator_(indicadorId);
  return 'Lançamento salvo com sucesso.';
}

function saveActionPlan(payload) {
  ensureSystemReady_();

  const indicadorId = sanitizeText_(payload.indicadorId);
  const items = Array.isArray(payload.items) ? payload.items : [];

  if (!indicadorId) throw new Error('Selecione um indicador.');

  const sh = getSheet_(APP.SHEETS.PLANO_ACAO);
  ensureActionPlanSchema_(sh);

  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]) === String(indicadorId)) {
      sh.deleteRow(i + 1);
    }
  }

  items.forEach(item => {
    const ano = toInt_(item.ano);
    const mes = toInt_(item.mes);
    const resultadoRef = toNumber_(item.resultadoRef);
    const fato = sanitizeText_(item.fato);
    const causa = sanitizeText_(item.causa);
    const acoes = sanitizeText_(item.acoes);
    const responsavel = sanitizeText_(item.responsavel);
    const status = sanitizeText_(item.status) || 'PENDENTE';
    const prazo = sanitizeText_(item.prazo);

    if (!ano || ano < 2000 || ano > 2100) return;
    if (!mes || mes < 1 || mes > 12) return;
    if (!fato && !causa && !acoes && !responsavel && !prazo && resultadoRef === null) return;

    const id = getNextId_('NEXT_ACTION_ID');
    sh.appendRow([
      id,
      indicadorId,
      ano,
      mes,
      resultadoRef,
      fato,
      causa,
      acoes,
      responsavel,
      status,
      prazo,
      new Date(),
      new Date()
    ]);
  });

  return 'Plano de ação salvo com sucesso.';
}

function getDashboardData(indicatorId, year, compareYears) {
  ensureSystemReady_();
  const indicators = getIndicators_();
  const history = getHistory_();
  const indicator = indicators.find(i => String(i.id) === String(indicatorId));

  if (!indicator) throw new Error('Indicador não encontrado.');

  const currentYear = toInt_(year) || new Date().getFullYear();
  const compareCount = toInt_(compareYears) || 2;

  const currentYearRows = history.filter(h => String(h.indicadorId) === String(indicatorId) && Number(h.ano) === currentYear);
  const currentYearMonthly = createMonthlyArray_(currentYearRows);

  const compareSeries = [];
  compareSeries.push({ year: currentYear, values: currentYearMonthly });
  for (let i = 1; i <= compareCount; i++) {
    const y = currentYear - i;
    const rows = history.filter(h => String(h.indicadorId) === String(indicatorId) && Number(h.ano) === y);
    compareSeries.push({ year: y, values: createMonthlyArray_(rows) });
  }

  return {
    indicatorId: indicator.id,
    indicatorName: indicator.nome,
    metaValor: indicator.metaValor,
    currentYear,
    currentYearMonthly,
    compareSeries
  };
}

function createConfigSheet_(ss) {
  let sh = ss.getSheetByName(APP.SHEETS.CONFIG);
  if (!sh) sh = ss.insertSheet(APP.SHEETS.CONFIG);
  sh.clear();
  sh.getRange(1, 1, 1, 2).setValues([['CHAVE', 'VALOR']]);
  styleHeader_(sh, 1, 2);
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, 2, 220);
}

function createListsSheet_(ss) {
  let sh = ss.getSheetByName(APP.SHEETS.LISTAS);
  if (!sh) sh = ss.insertSheet(APP.SHEETS.LISTAS);
  sh.clear();

  sh.getRange('A1').setValue('MESES');
  sh.getRange('A2:A13').setValues([['JAN'],['FEV'],['MAR'],['ABR'],['MAI'],['JUN'],['JUL'],['AGO'],['SET'],['OUT'],['NOV'],['DEZ']]);

  sh.getRange('C1').setValue('POLARIDADE');
  sh.getRange('C2:C3').setValues([['MENOR_MELHOR'],['MAIOR_MELHOR']]);

  sh.getRange('E1').setValue('STATUS_ACAO');
  sh.getRange('E2:E4').setValues([['PENDENTE'],['ANDAMENTO'],['CONCLUIDO']]);

  styleSimpleListTitles_(sh, ['A1','C1','E1']);
  sh.setColumnWidths(1, 6, 140);
}

function createIndicatorsSheet_(ss) {
  let sh = ss.getSheetByName(APP.SHEETS.INDICADORES);
  if (!sh) sh = ss.insertSheet(APP.SHEETS.INDICADORES);
  sh.clear();
  sh.getRange(1, 1, 1, 10).setValues([[
    'ID', 'NOME', 'META_VALOR', 'POLARIDADE', 'RESPONSAVEL', 'META_DESCRICAO',
    'DESCRICAO', 'ATIVO', 'CREATED_AT', 'UPDATED_AT'
  ]]);
  styleHeader_(sh, 1, 10);
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, 10, 170);
}

function createHistorySheet_(ss) {
  let sh = ss.getSheetByName(APP.SHEETS.HISTORICO);
  if (!sh) sh = ss.insertSheet(APP.SHEETS.HISTORICO);
  sh.clear();
  sh.getRange(1, 1, 1, 11).setValues([[
    'ID', 'INDICADOR_ID', 'ANO', 'MES', 'NUMERADOR', 'DENOMINADOR',
    'RESULTADO', 'MEDIANA_6', 'OBSERVACAO', 'CREATED_AT', 'UPDATED_AT'
  ]]);
  styleHeader_(sh, 1, 11);
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, 11, 150);
}

function createActionPlanSheet_(ss) {
  let sh = ss.getSheetByName(APP.SHEETS.PLANO_ACAO);
  if (!sh) sh = ss.insertSheet(APP.SHEETS.PLANO_ACAO);
  sh.clear();
  sh.getRange(1, 1, 1, 13).setValues([[
    'ID', 'INDICADOR_ID', 'ANO', 'MES', 'RESULTADO_REF', 'FATO', 'CAUSA',
    'ACOES', 'RESPONSAVEL', 'STATUS', 'PRAZO', 'CREATED_AT', 'UPDATED_AT'
  ]]);
  styleHeader_(sh, 1, 13);
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, 13, 170);
}

function styleHeader_(sheet, row, numCols) {
  const range = sheet.getRange(row, 1, 1, numCols);
  range.setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
}

function styleSimpleListTitles_(sheet, a1s) {
  a1s.forEach(a1 => {
    sheet.getRange(a1).setFontWeight('bold').setBackground('#dbeafe');
  });
}

function isSystemInitialized_(ss) {
  const sh = ss.getSheetByName(APP.SHEETS.CONFIG);
  if (!sh) return false;
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === 'APP_INITIALIZED' && String(values[i][1]) === 'TRUE') return true;
  }
  return false;
}

function ensureSystemReady_() {
  if (!isSystemInitialized_(SpreadsheetApp.getActiveSpreadsheet())) {
    throw new Error('A estrutura do sistema ainda não foi criada. Clique primeiro em "Criar estrutura automática na planilha".');
  }
}

function getIndicators_() {
  const sh = getSheet_(APP.SHEETS.INDICADORES);
  const data = sh.getDataRange().getValues();
  return data.slice(1).filter(r => r[0] !== '').map(r => ({
    id: r[0],
    nome: r[1],
    metaValor: toNumber_(r[2]),
    polaridade: r[3],
    responsavel: r[4],
    metaDescricao: r[5],
    descricao: r[6],
    ativo: String(r[7]).toUpperCase() === 'TRUE'
  }));
}

function getHistory_() {
  const sh = getSheet_(APP.SHEETS.HISTORICO);
  const data = sh.getDataRange().getValues();
  return data.slice(1).filter(r => r[0] !== '').map(r => ({
    id: r[0],
    indicadorId: r[1],
    ano: toInt_(r[2]),
    mes: toInt_(r[3]),
    numerador: toNumber_(r[4]),
    denominador: toNumber_(r[5]),
    resultado: toNumber_(r[6]),
    mediana6: toNumber_(r[7]),
    observacao: r[8]
  }));
}

function getActionPlans_() {
  const sh = getSheet_(APP.SHEETS.PLANO_ACAO);
  ensureActionPlanSchema_(sh);
  const data = sh.getDataRange().getValues();
  return data.slice(1).filter(r => r[0] !== '').map(r => ({
    id: r[0],
    indicadorId: r[1],
    ano: toInt_(r[2]),
    mes: toInt_(r[3]),
    resultadoRef: toNumber_(r[4]),
    fato: r[5],
    causa: r[6],
    acoes: r[7],
    responsavel: r[8],
    status: r[9],
    prazo: formatDateIso_(r[10])
  }));
}

function getActionPlansByIndicator(indicatorId) {
  ensureSystemReady_();

  const indicadorId = sanitizeText_(indicatorId);
  if (!indicadorId) throw new Error('Selecione um indicador.');

  const rows = getActionPlans_()
    .filter(r => String(r.indicadorId) === String(indicadorId))
    .sort((a, b) => {
      if (Number(b.ano) !== Number(a.ano)) return Number(b.ano) - Number(a.ano);
      return Number(b.mes) - Number(a.mes);
    });

  return {
    indicadorId,
    items: rows.map(r => ({
      ano: r.ano,
      mes: r.mes,
      resultadoRef: r.resultadoRef,
      fato: r.fato || '',
      causa: r.causa || '',
      acoes: r.acoes || '',
      responsavel: r.responsavel || '',
      status: r.status || 'PENDENTE',
      prazo: r.prazo || ''
    }))
  };
}

function ensureActionPlanSchema_(sheet) {
  const sh = sheet || getSheet_(APP.SHEETS.PLANO_ACAO);
  const lastCol = sh.getLastColumn();
  if (lastCol >= 13) return;

  if (lastCol === 12) {
    sh.insertColumnAfter(10);
    sh.getRange(1, 11).setValue('PRAZO');
    sh.getRange(1, 12).setValue('CREATED_AT');
    sh.getRange(1, 13).setValue('UPDATED_AT');
    styleHeader_(sh, 1, 13);
    sh.setColumnWidth(11, 170);
  }
}

function formatDateIso_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const txt = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(txt)) return txt;
  return '';
}

function getSheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('A aba "' + name + '" não existe.');
  return sh;
}

function setConfigValue_(key, value) {
  const sh = getSheet_(APP.SHEETS.CONFIG);
  const data = sh.getDataRange().getValues();
  let foundRow = null;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(key)) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow) {
    sh.getRange(foundRow, 2).setValue(value);
  } else {
    sh.appendRow([key, value]);
  }
}

function getConfigValue_(key, defaultValue) {
  const sh = getSheet_(APP.SHEETS.CONFIG);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(key)) {
      return data[i][1];
    }
  }
  return defaultValue;
}

function getNextId_(key) {
  const current = Number(getConfigValue_(key, 1)) || 1;
  setConfigValue_(key, String(current + 1));
  return current;
}

function recomputeMedianForIndicator_(indicatorId) {
  const sh = getSheet_(APP.SHEETS.HISTORICO);
  const data = sh.getDataRange().getValues();

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(indicatorId) && data[i][0] !== '') {
      rows.push({
        rowIndex: i + 1,
        ano: toInt_(data[i][2]),
        mes: toInt_(data[i][3]),
        resultado: toNumber_(data[i][6])
      });
    }
  }

  rows.sort((a, b) => a.ano - b.ano || a.mes - b.mes);

  rows.forEach((row, index) => {
    const previousSix = rows.slice(Math.max(0, index - 6), index)
      .map(r => r.resultado)
      .filter(v => v !== null);

    const median = previousSix.length ? median_(previousSix) : '';
    sh.getRange(row.rowIndex, 8).setValue(median);
  });
}

function median_(arr) {
  if (!arr.length) return '';
  const sorted = arr.slice().sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2
    ? sorted[mid]
    : (sorted[mid - 1] + sorted[mid]) / 2;
}

function createMonthlyArray_(rows) {
  const arr = Array(12).fill(null);
  rows.forEach(r => {
    if (r.mes >= 1 && r.mes <= 12) arr[r.mes - 1] = r.resultado;
  });
  return arr;
}

function sortYearMonthDesc_(a, b) {
  return Number(b.ano) - Number(a.ano) || Number(b.mes) - Number(a.mes);
}

function sanitizeText_(value) {
  return String(value || '').trim();
}

function toNumber_(value) {
  if (value === null || value === undefined || value === '') return null;
  const normalized = String(value)
    .replace(/\s/g, '')
    .replace(/\.(?=\d{3}(\D|$))/g, '')
    .replace(',', '.')
    .replace(/[^\d.-]/g, '');
  const num = Number(normalized);
  return isNaN(num) ? null : num;
}

function toInt_(value) {
  const num = parseInt(value, 10);
  return isNaN(num) ? null : num;
}
