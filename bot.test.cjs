const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

const sampleFoodBudget = (6.5 + 3) * 2;
const greekEuro = amount => amount.toFixed(2).replace('.', ',') + '€';

const props = new Map(Object.entries({
  TELEGRAM_BOT_TOKEN: '123:fake_test_token',
  SPREADSHEET_ID: 'test',
  PAIR_CODE: 'a-test-code-123456',
}));
const sheets = [['Date','Type','Category','Amount','Description','Status','ID']];
const sends = [];
let inbox = [];
let failSend = false;
let modelIntent;
let modelCalls = 0;
const modelStatuses = [];
const delays = [];
const ledger = {
  getLastRow: () => sheets.length,
  getRange(row, column, height) {
    return {
      setValues(values) { values.forEach((line,index) => { sheets[row - 1 + index] = line; }); },
      getValues() { return sheets.slice(row - 1, row - 1 + height); },
      setValue(value) { sheets[row - 1][column - 1] = value; },
      setNumberFormat() {},
      createTextFinder: key => ({
        matchEntireCell: () => ({
          findNext() {
            const index = sheets.findIndex(values => values[6] === key);
            return index < 0 ? null : { getRow: () => index + 1 };
          },
        }),
      }),
    };
  },
};
const book = {
  getSpreadsheetTimeZone: () => 'Etc/UTC',
  getSheetByName(name) {
    if (name === 'Telegram transactions') return ledger;
    if (name === 'Actuals vs plan') return dashboard;
    return {getRange: address => ({
      getValue: () => ({A10: '2032-11-19', B4: 8000, B8: sampleFoodBudget}[address]),
      getValues: () => [['2033-01-01',null,null,null,null,null,null,sampleFoodBudget,0,0,0,0,0,0,0]],
    })};
  },
};
const dashboardFormulas = {};
const dashboardValues = {};
const dashboard = {
  getRange(address) {
    const range = {
      merge: () => range, setBackground: () => range, setFontColor: () => range, setFontWeight: () => range,
      setValue() {}, setNumberFormat() {},
      setValues: values => { dashboardValues[address] = values; },
      getValues: () => address === 'I6:K6' ? [dashboardValues['I6:K32']?.[0] || ['', '', '']] : [],
      isBlank: () => !dashboardValues[address],
      setFormula: value => { dashboardFormulas[address] = value; },
      setFormulas: value => { dashboardFormulas[address] = value; },
    };
    return range;
  },
  setColumnWidth() {}, setColumnWidths() {}, setFrozenRows() {},
};
const context = vm.createContext({
  LockService: {getScriptLock: () => ({tryLock: () => true, releaseLock() {}})},
  PropertiesService: {getScriptProperties: () => ({
    getProperty: key => props.get(key), setProperty: (key, value) => props.set(key, value),
    deleteProperty: key => props.delete(key), setProperties: values => {
      for (const [key, value] of Object.entries(values)) props.set(key, value);
    },
  })},
  SpreadsheetApp: {openById: () => book, flush() {}},
  Utilities: {formatDate: () => '2032-11-19', sleep: delay => delays.push(delay)},
  UrlFetchApp: {fetch(url, options) {
    const params = JSON.parse(options.payload);
    if (url.endsWith(':generateContent')) {
      modelCalls++;
      assert.equal(options.headers['x-goog-api-key'], 'test-gemini-key');
      assert.equal(params.contents.length, 1);
      assert.equal(params.generationConfig.responseMimeType, 'application/json');
    assert.doesNotMatch(JSON.stringify(params.contents), /8000|5000|fake_test_token/);
      const status = modelStatuses.shift() || 200;
      return {getResponseCode: () => status, getContentText: () => JSON.stringify({
        candidates: [{finishReason: 'STOP', content: {parts: [{text: JSON.stringify(modelIntent)}]}}],
      })};
    }
    let result;
    if (url.endsWith('/getUpdates')) {
      result = inbox.filter(entry => entry.update_id >= (params.offset || 0));
    } else {
      if (failSend) throw new Error('simulated send failure');
      sends.push(params);
      result = true;
    }
    return {getContentText: () => JSON.stringify({ok: true, result})};
  }},
});
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8'), context);
const {parseBudgetMessage, pollBudgetBot, makeDashboard} = context;
const day = '2032-11-19';
for (const [text, amount, type, category, date = day] of [
  ['Spent €12.50 on petrol.',12.5,'Expense','fuel'],
  ['Πλήρωσα 11,50 € διόδια',11.5,'Expense','tolls'],
  ['275eur insurance',275,'Expense','insurance'],
  ['Received €675 for October invoice',675,'Income','work'],
  ['Received €125 in support',125,'Income','support income'],
  ['Paid back €45',45,'Debt repayment','other debt'],
  ['/repay 1.250,50 credit card',1250.5,'Debt repayment','credit card'],
  ['Paid €96 credit card',96,'Debt repayment','credit card'],
  ['Moved €400 to savings',400,'Savings','savings'],
  ['/withdraw 60',60,'Savings withdrawal','savings'],
  ['/expense 25 groceries 2032-11-17',25,'Expense','groceries','2032-11-17'],
  ['Χθες πλήρωσα 85 ευρώ για ασφάλιση εργασίας',85,'Expense','work_insurance','2032-11-18'],
  ['Subscription 9.50 euros',9.5,'Expense','subscription'],
  ['Συνδρομή 9,50 ευρώ',9.5,'Expense','subscription'],
  ['/expense 9.50 subscriptions',9.5,'Expense','subscription'],
  ['Paid 28 euros for internet',28,'Expense','telecommunication_bill'],
  ['Πλήρωσα 28 ευρώ για ίντερνετ',28,'Expense','telecommunication_bill'],
  ['Paid 85 euros for work insurance',85,'Expense','work_insurance'],
  ['/expense 28 telecommunication_bill',28,'Expense','telecommunication_bill'],
  ['/expense 85 work_insurance',85,'Expense','work_insurance'],
  ['Sandwich 6.50 euros',6.5,'Expense','groceries'],
  ['Paid €42 in road tax',42,'Expense','car'],
]) {
  const result = parseBudgetMessage(text, day);
  assert.ok(result, text);
  assert.deepEqual([result.amount,result.type,result.category,result.day], [amount,type,category,date], text);
}
for (const text of ['Maybe €100 back pay', 'Need to pay €275 insurance', '€25', 'Spent €25 food and €20 fuel', 'Paid $12 for fuel', '/save -10', '/expense 20 fuel 2032-11-20']) {
  assert.equal(parseBudgetMessage(text, day), null, text);
}
function update(update_id, from, text, type = 'private') {
  return {update_id, message: {message_id: update_id, date: Date.parse(day + 'T12:00:00Z') / 1000, chat: {id: from, type}, from: {id: from}, text}};
}
inbox = [
  update(1, 99, '/pair wrong'), update(2, 11, '/pair a-test-code-123456'),
  update(3, 99, 'Spent €99 on fuel'), update(4, 11, 'Spent €99 on fuel', 'group'),
  update(5, 11, 'Spent €12.50 on petrol'),
];
pollBudgetBot();
assert.equal(props.get('OWNER_ID'), '11');
assert.equal(props.has('PAIR_CODE'), false);
assert.equal(sends.length, 2);
assert.equal(sheets.length, 2);
assert.equal(sheets[1][3], 12.5);
inbox.push(update(6, 11, 'Spent €11.50 on tolls'));
failSend = true;
assert.throws(() => pollBudgetBot(), /simulated send failure|Telegram request failed/);
assert.equal(sheets.length, 3);
failSend = false;
pollBudgetBot();
assert.equal(sheets.length, 3);
assert.match(sends.at(-1).text, /Αποθήκευσα τα έξοδα για διόδια με αξία 11,50€/);
inbox.push(update(7, 11, '/undo tg-11-5'));
pollBudgetBot();
assert.equal(sheets[1][5], 'Voided');
makeDashboard({getSheetByName: () => null, insertSheet: () => dashboard});
assert.match(dashboardFormulas.B4, /Cash baseline/);
assert.match(dashboardFormulas['A7:H32'][0][7], /"work_insurance"/);
assert.equal(dashboardFormulas['A7:H32'].length, 26);
assert.equal(dashboardFormulas['A7:H32'][2][0], "='24-month plan'!A19");
assert.equal(dashboardFormulas['A7:H32'][25][0], "='24-month plan'!A42");
props.set('GEMINI_API_KEY', 'test-gemini-key');
modelIntent = {kind: 'transactions', items: [
  {type: 'Expense', category: 'fuel', amount: 20, day: '2032-11-18'},
  {type: 'Expense', category: 'tolls', amount: 11.5, day: '2032-11-18'},
]};
inbox.push(update(8, 11, '/ai Yesterday I spent €20 on fuel and €11.50 on tolls'));
pollBudgetBot();
assert.match(sends.at(-1).text, /\/confirm 8/);
assert.equal(sheets.length, 3);
inbox.push(update(9, 11, '/confirm 8'));
failSend = true;
assert.throws(() => pollBudgetBot(), /Telegram request failed/);
assert.equal(sheets.length, 5);
failSend = false;
pollBudgetBot();
assert.match(sends.at(-1).text, /Αποθήκευσα τα έξοδα για καύσιμα με αξία 20,00€/);
assert.match(sends.at(-1).text, /Αποθήκευσα τα έξοδα για διόδια με αξία 11,50€/);
assert.equal(sheets.length, 5);
modelIntent = {kind: 'question', query: {metric: 'spending', scope: 'car', start: '2032-11-01', end: day}};
inbox.push(update(10, 11, '/ask How much did I spend on the car this month?'));
pollBudgetBot();
assert.match(sends.at(-1).text, /^€43.00 in recorded spending/);
inbox.push(update(11, 11, '/undo tg-11-8-2'));
pollBudgetBot();
assert.equal(sheets[4][5], 'Voided');
modelIntent = {kind: 'question', query: {metric: 'cash balance', scope: 'all', start: day, end: day}};
inbox.push(update(12, 11, 'What is my cash balance?'));
pollBudgetBot();
assert.match(sends.at(-1).text, /€7988.50/);
inbox.push(update(13, 11, 'Spent €8 on petrol'));
pollBudgetBot();
assert.equal(modelCalls, 1);
assert.equal(sheets.length, 6);
modelIntent = {kind: 'transactions', items: [{type: 'Expense', category: 'fuel', amount: 99, day: '2032-11-20'}]};
inbox.push(update(14, 11, '/ai Planning to spend €99 on fuel tomorrow'));
pollBudgetBot();
assert.match(sends.at(-1).text, /couldn’t identify completed/);
assert.equal(sheets.length, 6);
modelIntent = {kind: 'query', query: {}};
inbox.push(update(15, 11, '/ask Bad response'));
pollBudgetBot();
assert.match(sends.at(-1).text, /invalid response/);
const previousModelCalls = modelCalls;
inbox.push(update(16, 11, 'Διόδια 11.20'));
pollBudgetBot();
assert.equal(modelCalls, previousModelCalls);
assert.match(sends.at(-1).text, /^Αποθήκευσα τα έξοδα για διόδια με αξία 11,20€\./);
assert.equal(sheets.at(-1)[3], 11.20);
assert.equal(sheets.at(-1)[2], 'tolls');
assert.match(sends.at(-1).text, /\/undo tg-11-16/);
const expectedCount = sheets.length;
sheets.at(-1)[3] = 12;
assert.match(context.handleBudgetMessage(ledger, 'Διόδια 11.20', day, 'tg-11-16', undefined, book), /Δεν μπόρεσα να επιβεβαιώσω/);
assert.equal(sheets.length, expectedCount);
const greek = context.localBudgetQuestion('Πόσα ξόδεψα σε βενζίνη τον προηγούμενο μήνα;', '2033-01-06');
assert.deepEqual([greek.scope, greek.start, greek.end], ['fuel', '2032-12-01', '2032-12-31']);
assert.equal(context.localBudgetQuestion('How much did I spend yesterday on the car or groceries?', day), null);
modelIntent = {kind: 'question', query: {metric: 'spending', scope: 'car', start: '2032-11-01', end: day}};
modelStatuses.push(503, 503, 200);
inbox.push(update(17, 11, '/ask Total for my car expenses so far in November'));
pollBudgetBot();
assert.match(sends.at(-1).text, /recorded spending \(car\)/);
assert.equal(delays.length, 2);
assert.ok(delays[0] >= 1000 && delays[0] < 1250);
assert.ok(delays[1] >= 2000 && delays[1] < 2250);
modelStatuses.push(503, 503, 503);
inbox.push(update(18, 11, '/ask Total for my car expenses so far in November'));
pollBudgetBot();
assert.match(sends.at(-1).text, /HTTP 503/);
assert.equal(delays.length, 4);
modelStatuses.push(403);
inbox.push(update(19, 11, '/ask Total for my car expenses so far in November'));
pollBudgetBot();
assert.match(sends.at(-1).text, /HTTP 403/);
assert.equal(delays.length, 4);
const foodModelCalls = modelCalls;
inbox.push(update(20, 11, 'σάντουιτς 6,50 ευρώ'), update(21, 11, 'καφές 3 ευρώ'));
pollBudgetBot();
assert.equal(modelCalls, foodModelCalls);
assert.deepEqual(sheets.slice(-2).map(row => [row[2], row[3]]), [['groceries', 6.5], ['groceries', 3]]);
assert.match(sends.at(-1).text, /φαγητό με αξία 3,00€/);
inbox.push(update(22, 11, 'Πόσα μου μένουν για φαγητό αυτόν τον μήνα;'));
pollBudgetBot();
assert.ok(sends.at(-1).text.includes(greekEuro(6.5 + 3) + ' από ' + greekEuro(sampleFoodBudget) + '. ' + greekEuro(sampleFoodBudget - 6.5 - 3) + ' απομένουν.'));
assert.equal(dashboardValues['I6:K32'][1][0], "='24-month plan'!$B$8");
assert.match(dashboardValues['I6:K32'][3][1], /"groceries"/);
assert.equal(dashboardValues['I6:K32'][3][2], '=ROUND(I9-J9,2)');
modelIntent = {kind: 'transactions', items: [{type: 'Expense', category: 'groceries', amount: 18, day}]};
inbox.push(update(23, 11, 'Μεζεδοπωλείο 18 ευρώ'));
pollBudgetBot();
assert.equal(sheets.at(-1)[3], 18);
assert.match(sends.at(-1).text, /φαγητό με αξία 18,00€/);
const postCount = sheets.length;
modelIntent = {kind: 'transactions', items: [{type: 'Expense', category: 'groceries', amount: 18, day}]};
inbox.push(update(24, 11, 'Μεζεδοπωλείο 19 ευρώ'));
pollBudgetBot();
assert.equal(sheets.length, postCount);
assert.match(sends.at(-1).text, /Δεν μπόρεσα να ταυτοποιήσω/);
props.set('SUBSCRIPTION_KEYWORDS', ' Example Stream, Μουσική Δείγμα ');
const callsBeforeAliases = modelCalls;
inbox.push(update(25, 11, 'Example Stream 6.25 euros'), update(26, 11, 'μουσική δείγμα 8,10 ευρώ'));
pollBudgetBot();
assert.equal(modelCalls, callsBeforeAliases);
assert.deepEqual(sheets.slice(-2).map(row => [row[2], row[3]]), [['subscription', 6.25], ['subscription', 8.1]]);
assert.equal(parseBudgetMessage('Example Streamer 6.25 euros', day, ['example stream']), null);
console.log('Passed food auto-categorization, food-budget tracking, verified Greek receipts, Gemini recovery, and existing bot checks.');
