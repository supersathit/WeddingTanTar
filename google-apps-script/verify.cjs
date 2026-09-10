const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const path = require('node:path');
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
for (const script of html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)) new Function(script[1]);
const tables = {RSVP:[], Slips:[]};
let files = 0;
const sheets = Object.fromEntries(Object.entries(tables).map(([name, rows]) => [name, {
  getLastRow:() => rows.length + 1,
  getRange:() => ({createTextFinder:id => ({matchEntireCell:() => ({findNext:() => rows.find(row => row[0] === id)})})}),
  appendRow:row => rows.push(row)
}]));
const context = vm.createContext({
  console:{error:() => {}},
  PropertiesService:{getScriptProperties:() => ({getProperty:key => key})},
  SpreadsheetApp:{openById:() => ({getSheetByName:name => sheets[name]}), flush:() => {}},
  LockService:{getScriptLock:() => ({waitLock:() => {}, releaseLock:() => {}})},
  Utilities:{base64Decode:value => [...Buffer.from(value, 'base64')], newBlob:(...args) => args},
  DriveApp:{getFolderById:() => ({createFile:() => { files++; return {getUrl:() => 'private-drive-file', setTrashed:() => {}}; }})},
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
console.log('PASS: browser script syntax, RSVP validation, formula escaping, retry deduplication, slip validation and file deduplication.');

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
  assert.equal(receipt, sent[0].requestId);
  assert.equal(new Set(sent.map(body => body.requestId)).size, 1, 'Retry must preserve request ID');
  assert.equal(tables.RSVP.at(-1)[2], 'Client test');
  assert.equal(tables.RSVP.at(-1)[4], 2);
  console.log('PASS: frontend-to-backend payload, absent crypto, failed network, rejected response, wrong receipt and successful retry.');
}
verifyClient().catch(error => { console.error(error); process.exitCode = 1; });
