import { COLUMNS, STUDENT_EXPORT_COLUMNS, STUDENT_KNOWN_COLUMNS } from './xlsx.js';

export const APP_VERSION = '1.0.22';
export const emptyState = () => ({ schema: 1, students: {}, grades: [], funding: [], records: [], settings: { threshold: 1 }, updatedAt: new Date().toISOString() });
export const uid = () => crypto.randomUUID?.() || [...crypto.getRandomValues(new Uint8Array(16))].map(v => v.toString(16).padStart(2, '0')).join('');
export const now = () => new Date().toISOString();
export const clean = value => String(value ?? '').trim();
export const isPasscode = value => /^\d{6}$/.test(String(value ?? ''));
export const dateFromId = id => {
  const text = clean(id);
  if (!/^\d{17}[\dXx]$/.test(text)) return '';
  const weights = [7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2];
  const checksum = '10X98765432'[[...text.slice(0,17)].reduce((sum, digit, index) => sum + Number(digit) * weights[index], 0) % 11];
  if (text[17].toUpperCase() !== checksum) return '';
  const year = +text.slice(6, 10), month = +text.slice(10, 12), day = +text.slice(12, 14);
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day ? `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}` : '';
};
export const excelDate = value => {
  const text = clean(value);
  if (!/^\d+(\.\d+)?$/.test(text) || Number(text) < 20000 || Number(text) > 80000) return text;
  return new Date(Date.UTC(1899, 11, 30) + Number(text) * 86400000).toISOString().slice(0, 10);
};
export const isFailed = grade => {
  const explicit = clean(grade.failed).toLowerCase();
  if (['是', '1', 'true', '挂科', '不及格'].includes(explicit)) return true;
  if (['否', '0', 'false', '及格'].includes(explicit)) return false;
  const score = Number(grade.score);
  return clean(grade.score) !== '' && Number.isFinite(score) && score < 60;
};
export const courseKey = grade => clean(grade.courseCode) || clean(grade.courseName);
export function gradeStats(state, studentId, term) {
  const entries = state.grades.filter(g => g.studentId === studentId);
  const failed = entries.filter(isFailed);
  return {
    semester: new Set(failed.filter(g => g.term === term).map(courseKey).filter(Boolean)).size,
    cumulative: new Set(failed.map(courseKey).filter(Boolean)).size,
    resolved: new Set(failed.filter(g => ['已通过', '补考通过', '重修通过'].includes(clean(g.retakeStatus))).map(courseKey)).size
  };
}
export function terms(state) { return [...new Set(state.grades.map(g => g.term).filter(Boolean))].sort().reverse(); }

const request = req => new Promise((resolve, reject) => { req.onsuccess = () => resolve(req.result); req.onerror = () => reject(req.error); });
let dbPromise;
let sessionKey;
const utf8 = new TextEncoder();
const decode = new TextDecoder();
const derive = async (password, salt) => {
  const raw = await crypto.subtle.importKey('raw', utf8.encode(password), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations: 250000, hash: 'SHA-256' }, raw, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
};
const encrypt = async bytes => {
  if (!sessionKey) throw new Error('请先解锁应用');
  const iv = crypto.getRandomValues(new Uint8Array(12));
  return { iv, data: await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, sessionKey, bytes) };
};
const decrypt = async envelope => {
  if (!sessionKey) throw new Error('请先解锁应用');
  return new Uint8Array(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: envelope.iv }, sessionKey, envelope.data));
};
const decryptWith = async (key, envelope) => new Uint8Array(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: envelope.iv }, key, envelope.data));
export function openDb() {
  if (!dbPromise) dbPromise = new Promise((resolve, reject) => {
    const req = indexedDB.open('counselor-workbench-v1', 1);
    req.onupgradeneeded = () => {
      req.result.createObjectStore('app');
      req.result.createObjectStore('files', { keyPath: 'id' });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbPromise;
}
export async function hasSecurity() { const db = await openDb(); return !!await request(db.transaction('app').objectStore('app').get('security')); }
export async function usesLegacySecurity() {
  const db = await openDb(); const security = await request(db.transaction('app').objectStore('app').get('security'));
  return !!security && security.mode !== 'pin6';
}
export async function setupSecurity(password) {
  if (!isPasscode(password)) throw new Error('解锁密码必须是 6 位数字');
  if (await hasSecurity()) throw new Error('已设置解锁密码');
  const salt = crypto.getRandomValues(new Uint8Array(16));
  sessionKey = await derive(password, salt);
  const verifier = await encrypt(utf8.encode('counselor-workbench-v1'));
  const db = await openDb();
  await request(db.transaction('app', 'readwrite').objectStore('app').put({ mode: 'pin6', salt, verifier }, 'security'));
  await saveState(emptyState());
}
export async function unlockSecurity(password) {
  const db = await openDb();
  const security = await request(db.transaction('app').objectStore('app').get('security'));
  if (!security) throw new Error('尚未设置解锁密码');
  if (security.mode === 'pin6' && !isPasscode(password)) throw new Error('请输入 6 位数字密码');
  sessionKey = await derive(password, security.salt);
  try { if (decode.decode(await decrypt(security.verifier)) !== 'counselor-workbench-v1') throw new Error(); }
  catch { sessionKey = undefined; throw new Error('解锁密码不正确'); }
}
export function lockSecurity() { sessionKey = undefined; }
export async function openLocalFiles() {
  const db = await openDb();
  const stored = await request(db.transaction('app').objectStore('app').get('file-device-key'));
  if (stored?.key) { sessionKey = stored.key; return; }
  sessionKey = await crypto.subtle.generateKey({ name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
  await request(db.transaction('app', 'readwrite').objectStore('app').put({ key: sessionKey }, 'file-device-key'));
}
export async function loadState() {
  const db = await openDb();
  const payload = await request(db.transaction('app').objectStore('app').get('state'));
  return payload ? JSON.parse(decode.decode(await decrypt(payload))) : emptyState();
}
export async function loadAccountState(accountId) {
  const db = await openDb(); const payload = await request(db.transaction('app').objectStore('app').get(`account-state:${accountId}`));
  return payload ? JSON.parse(decode.decode(await decrypt(payload))) : undefined;
}
export async function saveAccountState(accountId, state) {
  state.updatedAt = now(); const payload = await encrypt(utf8.encode(JSON.stringify(state)));
  const db = await openDb(); await request(db.transaction('app', 'readwrite').objectStore('app').put(payload, `account-state:${accountId}`));
}
export async function saveState(state) {
  state.updatedAt = now();
  const encryptedState = await encrypt(utf8.encode(JSON.stringify(state)));
  const db = await openDb();
  await request(db.transaction('app', 'readwrite').objectStore('app').put(encryptedState, 'state'));
}
const encodeFile = async file => {
  const meta = utf8.encode(JSON.stringify({ id: file.id, studentId: file.studentId, kind: file.kind, name: file.name, type: file.type, createdAt: file.createdAt }));
  const bytes = new Uint8Array(await file.blob.arrayBuffer());
  const combined = new Uint8Array(4 + meta.length + bytes.length);
  new DataView(combined.buffer).setUint32(0, meta.length, true);
  combined.set(meta, 4);
  combined.set(bytes, 4 + meta.length);
  return { id: file.id, ...(await encrypt(combined)) };
};
const decodeFile = async record => {
  if (!record) return undefined;
  const bytes = await decrypt(record);
  const length = new DataView(bytes.buffer).getUint32(0, true);
  const meta = JSON.parse(decode.decode(bytes.subarray(4, 4 + length)));
  return { ...meta, blob: new Blob([bytes.subarray(4 + length)], { type: meta.type }) };
};
export async function putFile(file) {
  const encryptedFile = await encodeFile(file);
  const db = await openDb();
  await request(db.transaction('files', 'readwrite').objectStore('files').put(encryptedFile));
}
export async function getFile(id) { const db = await openDb(); return decodeFile(await request(db.transaction('files').objectStore('files').get(id))); }
export async function allFiles() { const db = await openDb(); return Promise.all((await request(db.transaction('files').objectStore('files').getAll())).map(decodeFile)); }
export async function putFiles(files) {
  const records = await Promise.all(files.map(encodeFile));
  const db = await openDb(); const tx = db.transaction('files', 'readwrite');
  for (const record of records) tx.objectStore('files').put(record);
  await new Promise((resolve, reject) => { tx.oncomplete = resolve; tx.onerror = () => reject(tx.error); });
}
export async function changeSecurity(currentPassword, nextPassword) {
  if (!isPasscode(nextPassword)) throw new Error('新密码必须是 6 位数字');
  const db = await openDb();
  const security = await request(db.transaction('app').objectStore('app').get('security'));
  if (!security) throw new Error('尚未设置解锁密码');
  const currentKey = await derive(currentPassword, security.salt);
  try {
    if (decode.decode(await decryptWith(currentKey, security.verifier)) !== 'counselor-workbench-v1') throw new Error();
  } catch { throw new Error('当前解锁密码不正确'); }
  const oldSessionKey = sessionKey;
  try {
    sessionKey = currentKey;
    const currentState = await loadState();
    const currentFiles = await allFiles();
    const salt = crypto.getRandomValues(new Uint8Array(16));
    sessionKey = await derive(nextPassword, salt);
    const verifier = await encrypt(utf8.encode('counselor-workbench-v1'));
    const stateRecord = await encrypt(utf8.encode(JSON.stringify(currentState)));
    const fileRecords = await Promise.all(currentFiles.map(encodeFile));
    const tx = db.transaction(['app', 'files'], 'readwrite');
    const requests = [tx.objectStore('app').put({ mode: 'pin6', salt, verifier }, 'security'), tx.objectStore('app').put(stateRecord, 'state')];
    for (const record of fileRecords) requests.push(tx.objectStore('files').put(record));
    await Promise.all(requests.map(request));
    await new Promise((resolve, reject) => { tx.oncomplete = resolve; tx.onerror = () => reject(tx.error); });
  } catch (error) {
    sessionKey = oldSessionKey;
    throw error;
  }
}

function rowsToObjects(rows) {
  if (!rows?.length) return [];
  const headers = rows[0].map(clean);
  return rows.slice(1).filter(row => row.some(v => clean(v))).map((row, index) => ({ rowNumber: index + 2, data: Object.fromEntries(headers.map((key, i) => [key, clean(row[i])])) }));
}
const stable = parts => parts.map(clean).join('\u001f');
export function previewWorkbook(state, workbook) {
  const next = structuredClone(state);
  const errors = [];
  const counts = { students: 0, grades: 0, funding: 0, records: 0, updated: 0 };
  for (const { rowNumber, data } of rowsToObjects(workbook.学生)) {
    const id = data.学号, name = data.姓名;
    if (!id) { errors.push(`学生表第 ${rowNumber} 行缺少学号，已跳过`); continue; }
    const old = next.students[id];
    const custom = { ...(old?.custom || {}) };
    for (const [key, value] of Object.entries(data)) if (key && !STUDENT_KNOWN_COLUMNS.has(key) && value) custom[key] = value;
    next.students[id] = { ...old, id, name: name || old?.name || '', className: data.班级 || old?.className || '', classRole: data.班级职务 || old?.classRole || '', major: data.专业 || old?.major || '', dorm: data.宿舍 || old?.dorm || '', phone: data.本人电话 || old?.phone || '', parentPhone: data.家长电话 || old?.parentPhone || '', address: data.家庭住址 || old?.address || '', idNumber: data.身份证号 || old?.idNumber || '', gender: data.性别 || old?.gender || '', ethnicity: data.民族 || old?.ethnicity || '', birthDate: excelDate(data.出生日期) || dateFromId(data.身份证号 || old?.idNumber) || old?.birthDate || '', custom, updatedAt: now() };
    counts[old ? 'updated' : 'students']++;
  }
  const configs = [
    ['成绩', 'grades', data => ({ id: data.记录ID || uid(), studentId: data.学号, term: data.学期, courseCode: data.课程代码, courseName: data.课程名称, score: data.成绩, failed: data.是否挂科, retakeStatus: data.补考或重修状态, note: data.备注 }) , data => stable([data.学号, data.学期, data.课程代码 || data.课程名称])],
    ['资助', 'funding', data => ({ id: data.记录ID || uid(), studentId: data.学号, hardship: data.困难等级, family: data.家庭情况, program: data.资助项目, paidAt: excelDate(data.发放日期), amount: data.发放金额, note: data.备注 }), data => stable([data.学号, data.资助项目, data.发放日期, data.发放金额])],
    ['工作记录', 'records', data => ({ id: data.记录ID || uid(), studentId: data.学号, type: data.类型 || '谈话', date: excelDate(data.日期), subject: data.对象, attention: data.关注等级, source: data.来源, summary: data.摘要, todo: data.后续待办, outcome: data.处理结论, dueDate: excelDate(data.到期日期), done: ['是', '已完成', 'true', '1'].includes(clean(data.完成状态).toLowerCase()), attachments: [] }), data => stable([data.学号, data.类型, data.日期, data.对象, data.摘要])]
  ];
  for (const [sheetName, key, make, naturalKey] of configs) {
    for (const { rowNumber, data } of rowsToObjects(workbook[sheetName])) {
      if (!data.学号) { errors.push(`${sheetName}表第 ${rowNumber} 行缺少学号，已跳过`); continue; }
      if (!next.students[data.学号]) { errors.push(`${sheetName}表第 ${rowNumber} 行的学号 ${data.学号} 未在学生表中找到，已跳过`); continue; }
      const candidate = make(data);
      const oldIndex = next[key].findIndex(item => item.id === data.记录ID || (!data.记录ID && naturalKey(reverseRow(sheetName, item)) === naturalKey(data)));
      if (oldIndex < 0) { next[key].push({ ...candidate, updatedAt: now() }); counts[key]++; }
      else { next[key][oldIndex] = { ...next[key][oldIndex], ...candidate, id: next[key][oldIndex].id, attachments: next[key][oldIndex].attachments, updatedAt: now() }; counts.updated++; }
    }
  }
  return { next, errors, counts };
}

function reverseRow(name, item) {
  if (name === '成绩') return { 学号: item.studentId, 学期: item.term, 课程代码: item.courseCode, 课程名称: item.courseName };
  if (name === '资助') return { 学号: item.studentId, 资助项目: item.program, 发放日期: item.paidAt, 发放金额: item.amount };
  return { 学号: item.studentId, 类型: item.type, 日期: item.date, 对象: item.subject, 摘要: item.summary };
}

export function exportSheets(state) {
  const students = Object.values(state.students);
  const customKeys = [...new Set(students.flatMap(s => Object.keys(s.custom || {})))].sort();
  return {
    学生: [[...STUDENT_EXPORT_COLUMNS, ...customKeys], ...students.map((s, index) => [index + 1, s.major, s.className, s.id, s.name, s.classRole, s.gender, s.ethnicity, s.idNumber, s.dorm, s.phone, s.parentPhone, s.address, s.birthDate, ...customKeys.map(k => s.custom?.[k] || '')])],
    成绩: [COLUMNS.成绩, ...state.grades.map(g => [g.id, g.studentId, g.term, g.courseCode, g.courseName, g.score, g.failed, g.retakeStatus, g.note])],
    资助: [COLUMNS.资助, ...state.funding.map(f => [f.id, f.studentId, f.hardship, f.family, f.program, f.paidAt, f.amount, f.note])],
    工作记录: [COLUMNS.工作记录, ...state.records.map(r => [r.id, r.studentId, r.type, r.date, r.subject, r.attention, r.source, r.summary, r.todo, r.outcome, r.dueDate, r.done ? '已完成' : '未完成'])]
  };
}
