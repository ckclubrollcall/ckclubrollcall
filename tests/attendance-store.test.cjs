const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const path=require('node:path');
function context() {
  const c=vm.createContext({Date,Utilities:{base64Decode:s=>Array.from(Buffer.from(s,'base64')),base64Encode:b=>Buffer.from(b).toString('base64')}});
  vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas/attendance_store.js'),'utf8'),c);
  return c;
}
const png='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAADklEQVR4nGNgYGD4D8IABgMB/8+HxnAAAAAASUVORK5CYII=';
test('valid PNG accepted',()=>assert.ok(context().validateSignaturePng(png).length>57));
for(const [name,value] of [['KEEP','KEEP'],['empty',''],['missing',undefined],['object',{}],['URL','https://drive.google.com/file/d/1234567890/view'],['large','x'.repeat(200001)],['bad base64','data:image/png;base64,AA=A'],['fake PNG','data:image/png;base64,'+Buffer.alloc(100).toString('base64')]]) {
  test('reject '+name,()=>assert.throws(()=>context().validateSignaturePng(value)));
}
test('CRC tampering rejected',()=>{
  const b=Buffer.from(png.split(',')[1],'base64');b[20]^=1;
  assert.throws(()=>context().validateSignaturePng('data:image/png;base64,'+b.toString('base64')));
});
test('headers reordered; legacy columns allowed; duplicates and missing fields rejected',()=>{
  const c=context(),h=Array.from(c.ATTENDANCE_FIELDS,f=>f[1]);
  assert.equal(c.attendanceHeaderMap(h.slice(0,11),true).recordType,undefined);
  assert.equal(c.attendanceHeaderMap(h.slice().reverse()).signatureUrl,1);
  assert.throws(()=>c.attendanceHeaderMap(h.concat('老師簽名')));
  assert.throws(()=>c.attendanceHeaderMap(h.slice(1)));
});
test('real chip payload, canonical URL, hostile domains rejected',()=>{
  const c=context(),chip=c.signatureChip('1234567890_abcd');
  assert.equal(chip.userEnteredValue.stringValue,'@');
  assert.equal(chip.chipRuns[0].chip.richLinkProperties.uri,'https://drive.google.com/file/d/1234567890_abcd/view');
  assert.equal(c.signatureFileId('https://drive.google.com.evil.example/file/d/1234567890'),null);
});
test('timestamp then row determines latest',()=>{
  const c=context(),date='2026-01-02';
  assert.equal(c.selectLatestAttendance([{date,_ts:2,_row:3},{date,_ts:1,_row:4},{date,_ts:2,_row:5}])[date]._row,5);
});
test('summary insertion does not shift teacher lookup',()=>{
  const c=context();
  assert.equal(c.summaryHeaderColumn(['最新的簽名圖片','社團名稱','最後修改日期','指導老師姓名'],['指導老師姓名']),3);
});
test('user strings are not formulas',()=>{
  assert.equal(context().attendanceCell('=malicious()').userEnteredValue.stringValue,'=malicious()');
});
