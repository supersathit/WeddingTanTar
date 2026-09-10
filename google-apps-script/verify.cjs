const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const path = require('node:path');
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
for (const script of html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)) new Function(script[1]);
const tables = {RSVP:[], Slips:[]};
let files = 0;
const properties = {};
const fileNames = [];
const sheets = Object.fromEntries(Object.entries(tables).map(([name, rows]) => [name, {
  getLastRow:() => rows.length + 1,
  getRange:(row, column, count) => ({
    getValues:() => rows.slice(row - 2, row - 2 + count),
    getValue:() => rows[row - 2][column - 1],
    setValue:value => { rows[row - 2][column - 1] = value; },
    createTextFinder:id => ({matchEntireCell:() => ({findNext:() => {
      const index = rows.findIndex(entry => entry[0] === id);
      return index < 0 ? null : {getRow:() => index + 2};
    }})})
  }),
  appendRow:row => rows.push(row)
}]));
const context = vm.createContext({
  console:{error:() => {}},
  PropertiesService:{getScriptProperties:() => ({getProperty:key => properties[key], setProperty:(key, value) => { properties[key] = value; }})},
  SpreadsheetApp:{openById:() => ({getSheetByName:name => sheets[name]}), flush:() => {}},
  LockService:{getScriptLock:() => ({waitLock:() => {}, releaseLock:() => {}})},
  Utilities:{base64Decode:value => [...Buffer.from(value, 'base64')], newBlob:(...args) => args},
  DriveApp:{getFolderById:() => ({createFile:blob => { files++; fileNames.push(blob[2]); return {getUrl:() => 'private-drive-file', setTrashed:() => {}}; }})},
  ContentService:{MimeType:{JSON:'json'}, createTextOutput:text => ({setMimeType:() => JSON.parse(text)})}
});
vm.runInContext(fs.readFileSync(path.join(__dirname, 'Code.gs'), 'utf8'), context);
const post = data => context.doPost({postData:{contents:JSON.stringify(data)}});
const rsvp = {type:'rsvp', requestId:'test-request-00001', name:'=IMPORTXML("x")', attend:'มา', guests:'2', session:'ทั้งวัน', wish:'ยินดีด้วย'};
assert.equal(post(rsvp).ok, true);
assert.equal(tables.RSVP[0][2][0], "'");
assert.equal(post(rsvp).ok, true);
assert.equal(tables.RSVP.length, 1, 'Retry must not duplicate a row');
assert.equal(post({...rsvp, requestId:'test-request-00002', guests:'21'}).ok, false);
assert.equal(post({...rsvp, requestId:'test-request-00003', attend:'ไม่มา', guests:undefined, session:undefined}).ok, true);
assert.equal(tables.RSVP[1][4], 0);
const slip = {type:'slip', requestId:'test-request-00004', name:'Test guest', fileName:'slip.pdf', mimeType:'application/pdf', base64:Buffer.from('%PDF-1.4 mock').toString('base64')};
assert.equal(post(slip).ok, true);
assert.equal(post(slip).ok, true);
assert.equal(files, 1, 'Retry must not create a second Drive file');
assert.equal(post({...slip, requestId:'test-request-00005', base64:Buffer.from('not a PDF').toString('base64')}).ok, false);
assert.equal(files, 1);
assert.match(post(slip).receiptId, /^\d{5}$/);
assert.equal(fileNames[0], `${post(slip).receiptId}_Test guest.pdf`);
assert.equal(tables.Slips[0][3], fileNames[0]);
assert.notEqual(post(slip).receiptId, post(rsvp).receiptId);
const thaiSlip = {...slip, requestId:'test-request-00006', name:'สมชาย ใจดี/ทดสอบ'};
const thaiReceipt = post(thaiSlip).receiptId;
assert.equal(fileNames[1], `${thaiReceipt}_สมชาย ใจดี_ทดสอบ.pdf`);
assert.equal(post(thaiSlip).receiptId, thaiReceipt);
console.log('PASS: browser script syntax, RSVP validation, formula escaping, retry deduplication, slip validation and file deduplication.');
const publicPayload = {...rsvp, requestId:'test-public-000001', name:'Private full name', publicName:'เพื่อนเจ้าสาว', publishWish:'yes', wish:'มีความสุขมาก ๆ นะ', approved:true};
assert.equal(post(publicPayload).ok, true);
const publicRow = tables.RSVP.at(-1);
const getWishes = () => context.doGet({parameter:{action:'wishes'}});
assert.equal(getWishes().wishes.length, 1, 'Consented wish publishes immediately; legacy rows stay private');
assert.deepEqual(Object.keys(getWishes().wishes[0]).sort(), ['name', 'wish']);
assert.equal(getWishes().wishes[0].name, 'เพื่อนเจ้าสาว');
assert.equal(JSON.stringify(getWishes()).includes('Private full name'), false);
publicRow[8] = false;
assert.equal(getWishes().wishes.length, 0, 'Non-consenting wishes stay private');
publicRow[8] = true;
publicRow[10] = false;
assert.equal(getWishes().wishes.length, 1, 'Legacy approval column no longer affects publication');
publicRow[9] = '';
assert.equal(getWishes().wishes.length, 0, 'Public name is required');
publicRow[9] = 'เพื่อนเจ้าสาว';
publicRow[6] = ' ';
assert.equal(getWishes().wishes.length, 0, 'Empty wishes stay hidden');
assert.equal(context.doGet({parameter:{action:'slips'}}).ok, false);
assert.equal(post({...publicPayload, requestId:'test-public-000002', publicName:''}).ok, false);
console.log('PASS: immediate publication with consent and public name; private fields, empty wishes and non-consenting rows excluded.');

async function verifyClient() {
  const start = html.indexOf('const GAS_URL =');
  const end = html.indexOf("document.getElementById('sendSlip').addEventListener", start);
  const sent = [];
  let mode = 'network';
  const browser = vm.createContext({
    crypto:undefined, AbortController, setTimeout, clearTimeout, TypeError, SyntaxError,
    console:{error:() => {}},
    fetch:async (url, options) => {
      const body = JSON.parse(options.body);
      sent.push(body);
      assert.equal(options.mode, undefined, 'Must not use no-cors');
      if (mode === 'network') throw new TypeError('Failed to fetch');
      return {ok:true, json:async () => mode === 'wrong-id'
        ? {ok:true, requestId:'wrong-request-id'}
        : mode === 'rejected' ? {ok:false, message:'rejected'}
        : post(body)};
    }
  });
  vm.runInContext(html.slice(start, end), browser);
  const payload = {type:'rsvp', name:'Client test', attend:'มา', guests:'2', session:'ทั้งวัน', wish:'Test'};
  await assert.rejects(browser.sendWeddingForm(payload));
  mode = 'wrong-id';
  await assert.rejects(browser.sendWeddingForm(payload));
  mode = 'rejected';
  await assert.rejects(browser.sendWeddingForm(payload));
  mode = 'success';
  const receipt = await browser.sendWeddingForm(payload);
  assert.match(receipt, /^\d{5}$/);
  assert.equal(receipt, post(sent[0]).receiptId);
  assert.equal(new Set(sent.map(body => body.requestId)).size, 1, 'Retry must preserve request ID');
  assert.equal(tables.RSVP.at(-1)[2], 'Client test');
  assert.equal(tables.RSVP.at(-1)[4], 2);
  console.log('PASS: frontend-to-backend payload, absent crypto, failed network, rejected response, wrong receipt and successful retry.');
}
verifyClient().catch(error => { console.error(error); process.exitCode = 1; });
