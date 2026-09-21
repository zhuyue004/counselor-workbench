import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { unzipSync, strFromU8 } from './vendor/fflate.js';
import { APP_VERSION, emptyState, dateFromId, gradeStats, previewWorkbook, exportSheets, isPasscode } from './data.js';
import { COLUMNS, STUDENT_EXPORT_COLUMNS, blankTemplate, createWorkbook } from './xlsx.js';
import { makePackage, readPackage, mergePackage, packageName } from './transfer.js';

test('fixed workbook has four worksheets and text identifiers', async () => {
  const zip = unzipSync(new Uint8Array(await blankTemplate().arrayBuffer()));
  const workbook = strFromU8(zip['xl/workbook.xml']);
  assert.match(workbook, /name="学生"/);
  assert.match(workbook, /name="成绩"/);
  assert.match(workbook, /name="资助"/);
  assert.match(workbook, /name="工作记录"/);
  assert.match(strFromU8(zip['xl/worksheets/sheet1.xml']), /t="inlineStr"/);
  assert.equal(Object.keys(COLUMNS).length, 4);
  assert.deepEqual(COLUMNS.学生, ['序号', '专业', '班级', '学号', '姓名', '性别', '民族', '身份证号', '宿舍', '本人电话', '家长电话', '家庭住址']);
});

test('Excel preview matches by student ID, adds custom fields and avoids duplicates', () => {
  const workbook = {
    学生: [[...COLUMNS.学生, '自定义_宿舍楼'], ['1','', '一班','00123','','','','','','','','','6栋'], ['', '', '', '', '', '', '', '', '', '', '', '', '']],
    成绩: [COLUMNS.成绩, ['', '00123','2025-2026-2','MATH01','高数','55','是','补考通过','']],
    资助: [COLUMNS.资助], 工作记录: [COLUMNS.工作记录]
  };
  const first = previewWorkbook(emptyState(),workbook);
  assert.equal(first.errors.length,0);
  assert.equal(first.next.students['00123'].name,'');
  assert.equal(first.next.students['00123'].ethnicity,'');
  assert.equal(first.next.students['00123'].custom['自定义_宿舍楼'],'6栋');
  assert.equal(gradeStats(first.next,'00123','2025-2026-2').cumulative,1);
  assert.equal(gradeStats(first.next,'00123','2025-2026-2').resolved,1);
  const second = previewWorkbook(first.next,workbook);
  assert.equal(second.next.grades.length,1);
  assert.deepEqual(exportSheets(second.next).学生[0].slice(0, STUDENT_EXPORT_COLUMNS.length), STUDENT_EXPORT_COLUMNS);
  assert.equal(exportSheets(second.next).学生[1][3],'00123');
  assert.equal(exportSheets(second.next).学生[1][4],'');
});

test('birth date only comes from a valid mainland ID checksum', () => {
  assert.equal(dateFromId('11010519491231002X'),'1949-12-31');
  assert.equal(dateFromId('110105194912310021'),'');
});

test('local unlock password requires exactly six digits', () => {
  assert.equal(isPasscode('123456'), true);
  assert.equal(isPasscode('12345'), false);
  assert.equal(isPasscode('1234567'), false);
  assert.equal(isPasscode('12a456'), false);
});

test('encryption finishes before IndexedDB write transactions begin', () => {
  const source = readFileSync(new URL('./data.js', import.meta.url), 'utf8');
  assert.match(source, /const encryptedState = await encrypt\(utf8\.encode\(JSON\.stringify\(state\)\)\);\s*const db = await openDb\(\);\s*await request\(db\.transaction\('app', 'readwrite'\)\.objectStore\('app'\)\.put\(encryptedState, 'state'\)\)/s);
  assert.match(source, /const encryptedFile = await encodeFile\(file\);\s*const db = await openDb\(\);\s*await request\(db\.transaction\('files', 'readwrite'\)\.objectStore\('files'\)\.put\(encryptedFile\)\)/s);
});

test('encrypted ZIP roundtrip preserves attachments and named version timestamp', async () => {
  const state = emptyState(); state.students['00123'] = { id:'00123',name:'张三',updatedAt:'2026-09-20T00:00:00.000Z' };
  const files = [{ id:'photo1',studentId:'00123',kind:'photo',name:'一班_张三_证件照_20260920_1.jpg',type:'image/jpeg',createdAt:'2026-09-20T00:00:00.000Z',blob:new Blob([new Uint8Array([1,2,3])],{type:'image/jpeg'}) }];
  const blob = await makePackage(state,files,'test-passphrase');
  const zip = unzipSync(new Uint8Array(await blob.arrayBuffer()));
  assert.deepEqual(Object.keys(zip).sort(),['manifest.json','payload.enc']);
  const restored = await readPackage(blob,'test-passphrase');
  assert.equal(restored.state.students['00123'].name,'张三');
  assert.deepEqual([...new Uint8Array(await restored.files[0].blob.arrayBuffer())],[1,2,3]);
  await assert.rejects(readPackage(blob,'wrong-password'),/密码不正确/);
  assert.match(packageName(new Date(2026,8,20,14,5,6)),new RegExp(`v${APP_VERSION.replaceAll('.','\\.')}_20260920_140506\\.zip$`));
});

test('package merge keeps newer local records', () => {
  const current = emptyState(), incoming = emptyState();
  current.students.a = {id:'a',name:'本机',updatedAt:'2026-09-20T12:00:00Z'};
  incoming.students.a = {id:'a',name:'旧包',updatedAt:'2026-09-19T12:00:00Z'};
  incoming.students.b = {id:'b',name:'新增',updatedAt:'2026-09-20T12:00:00Z'};
  const result = mergePackage(current,incoming);
  assert.equal(result.merged.students.a.name,'本机');
  assert.equal(result.merged.students.b.name,'新增');
});
