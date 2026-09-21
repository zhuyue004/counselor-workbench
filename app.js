import { APP_VERSION, uid, now, clean, dateFromId, gradeStats, terms, isFailed, emptyState, hasSecurity, usesLegacySecurity, setupSecurity, unlockSecurity, changeSecurity, lockSecurity, loadState, saveState, putFile, getFile, allFiles, putFiles, previewWorkbook, exportSheets } from './data.js';
import { COLUMNS, blankTemplate, createWorkbook, readWorkbook } from './xlsx.js';
import { makePackage, readPackage, mergePackage, packageName, stamp } from './transfer.js';
import { unzipSync } from './vendor/fflate.js';

const app = document.querySelector('#app');
const modalRoot = document.querySelector('#modal-root');
const toastEl = document.querySelector('#toast');
const inputs = { xlsx: document.querySelector('#xlsx-input'), photozip: document.querySelector('#photozip-input'), package: document.querySelector('#package-input') };
let state, view = 'home', selectedId = '', detailTab = 'info', search = '', selectedTerm = '', workModule = '', workSearch = '', todoFilter = 'all', photoUrl = '', toastTimer;
const today = () => new Date().toISOString().slice(0, 10);
const esc = value => String(value ?? '').replace(/[&<>"']/g, char => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[char]));
const fmtDate = value => clean(value) ? esc(clean(value).slice(0, 10)) : '—';
const maskId = value => value ? `${esc(value.slice(0, 4))}**********${esc(value.slice(-4))}` : '—';
const student = () => state.students[selectedId];
const displayName = value => clean(value) || '—';
const studentName = id => state.students[id] ? displayName(state.students[id].name) : id;
const studentInitial = s => displayName(s?.name).slice(0, 1);
const courseLabel = grade => grade.courseName || grade.courseCode || '未命名课程';
const releases = [
  { version: '1.0.7', updatedAt: '2026-09-21 16:45', notes: ['工作模块支持选择学生后直接新增成绩、资助及各类工作记录。', '完善待办筛选、资料完整度、导入反馈、备份状态和心理工作字段。'] },
  { version: '1.0.6', updatedAt: '2026-09-21 16:31', notes: ['底部导航调整为首页、学生、工作、待办和设置。', '新增工作中心，集中查看组织发展、学业成绩、资助、心理和家校联系。'] },
  { version: '1.0.5', updatedAt: '2026-09-21 16:24', notes: ['学生详情中的本人电话和家长电话支持点击拨号。'] },
  { version: '1.0.4', updatedAt: '2026-09-21 16:19', notes: ['“数据”模块改为“设置”。', '新增当前版本、更新时间和历史更新记录。'] },
  { version: '1.0.3', updatedAt: '2026-09-21 16:12', notes: ['顶部改为紧凑栏，移除英文标题和冗余提示。'] },
  { version: '1.0.2', updatedAt: '2026-09-21 16:04', notes: ['学生模板调整为序号、专业、班级、学号、姓名、性别、民族等列。', '除学号外的学生字段可留空。'] },
  { version: '1.0.1', updatedAt: '2026-09-21 15:50', notes: ['修复手工保存学生档案时 IndexedDB 写入失败的问题。', '解锁密码改为 6 位数字。'] },
  { version: '1.0.0', updatedAt: '2026-09-21 14:53', notes: ['完成 iPhone 版学生档案、成绩预警、资助和工作记录。'] }
];
async function normalizedPhoto(blob) {
  const url = URL.createObjectURL(blob);
  try {
    const image = new Image(); image.src = url; await image.decode();
    const scale = Math.min(1, 900 / Math.max(image.naturalWidth, image.naturalHeight));
    const canvas = document.createElement('canvas'); canvas.width = Math.max(1, Math.round(image.naturalWidth * scale)); canvas.height = Math.max(1, Math.round(image.naturalHeight * scale));
    canvas.getContext('2d').drawImage(image,0,0,canvas.width,canvas.height);
    const result = await new Promise(resolve => canvas.toBlob(resolve,'image/jpeg',.85));
    return result || blob;
  } catch { return blob; }
  finally { URL.revokeObjectURL(url); }
}
function notify(message, error = false) { toastEl.textContent = message; toastEl.className = error ? 'show error' : 'show'; clearTimeout(toastTimer); toastTimer = setTimeout(() => toastEl.className = '', 3600); }
function fileName(s, type, date, seq, ext) { return `${(s.className || '未分班').replace(/[\\/:*?"<>|]/g, '_')}_${displayName(s.name).replace(/[\\/:*?"<>|]/g, '_')}_${type}_${(date || today()).replaceAll('-', '')}_${seq}.${ext}`; }
function download(blob, name) { const url = URL.createObjectURL(blob); const link = document.createElement('a'); link.href = url; link.download = name; document.body.append(link); link.click(); link.remove(); setTimeout(() => URL.revokeObjectURL(url), 60000); }
function showModal(title, body, submitText, onSubmit) {
  modalRoot.innerHTML = `<div class="modal-shade"><div class="modal" role="dialog" aria-modal="true" aria-label="${esc(title)}"><div class="modal-head"><h2>${esc(title)}</h2><button type="button" class="close" data-close>×</button></div><form id="modal-form">${body}<div class="modal-actions"><button type="button" class="btn ghost" data-close>取消</button><button class="btn" type="submit">${esc(submitText)}</button></div></form></div></div>`;
  modalRoot.querySelectorAll('[data-close]').forEach(button => button.addEventListener('click', closeModal));
  modalRoot.querySelector('#modal-form').addEventListener('submit', async event => {
    event.preventDefault();
    const form = event.currentTarget, button = form.querySelector('[type=submit]');
    button.disabled = true;
    try { await onSubmit(new FormData(form), form); if (modalRoot.contains(form)) closeModal(); }
    catch (error) { notify(error.message || String(error), true); button.disabled = false; }
  });
  modalRoot.querySelector('input,select,textarea')?.focus();
}
function closeModal() { modalRoot.innerHTML = ''; }
const field = (label, name, value = '', type = 'text', extra = '') => `<div class="field"><label for="f-${esc(name)}">${esc(label)}</label><input id="f-${esc(name)}" name="${esc(name)}" type="${type}" value="${esc(value)}" ${extra}></div>`;
const area = (label, name, value = '') => `<div class="field"><label for="f-${esc(name)}">${esc(label)}</label><textarea id="f-${esc(name)}" name="${esc(name)}">${esc(value)}</textarea></div>`;
const choose = (label, name, options, current = '') => `<div class="field"><label for="f-${esc(name)}">${esc(label)}</label><select id="f-${esc(name)}" name="${esc(name)}">${options.map(option => `<option value="${esc(option)}" ${option === current ? 'selected' : ''}>${esc(option)}</option>`).join('')}</select></div>`;
async function persist() { await saveState(state); render(); }
function setView(next) { view = next; if (next !== 'work') workModule = ''; render(); window.scrollTo(0, 0); }

async function gate() {
  const exists = await hasSecurity();
  const legacy = exists && await usesLegacySecurity();
  const passwordAttrs = legacy ? 'required autocomplete="current-password"' : 'required inputmode="numeric" pattern="[0-9]{6}" maxlength="6" autocomplete="current-password"';
  app.innerHTML = `<div class="lock-page"><div class="lock-card"><div class="lock-mark">档</div><h1>${exists ? '欢迎回来' : '建立你的工作台'}</h1><p>${exists ? (legacy ? '输入原解锁密码后，可在“数据”页面改为 6 位数字密码。' : '输入 6 位数字密码，查看保存在这台设备上的学生资料。') : '学生资料仅保存在这台设备中。请设置 6 位数字密码，并定期导出加密数据包备份。'}</p><form id="gate-form">${field(legacy ? '原解锁密码' : '6 位数字密码', 'password', '', 'password', passwordAttrs)}${exists ? '' : field('再次输入 6 位数字密码', 'confirm', '', 'password', 'required inputmode="numeric" pattern="[0-9]{6}" maxlength="6" autocomplete="new-password"')}<button class="btn" type="submit">${exists ? '解锁' : '创建并进入'}</button></form><div class="lock-foot">忘记密码无法解密本机资料。请妥善保存密码和导出的加密备份包。</div></div></div>`;
  app.querySelector('#gate-form').addEventListener('submit', async event => {
    event.preventDefault(); const form = new FormData(event.currentTarget); const password = String(form.get('password'));
    try {
      if (exists) await unlockSecurity(password);
      else { if (password !== form.get('confirm')) throw new Error('两次输入的密码不一致'); await setupSecurity(password); }
      state = await loadState(); selectedTerm = terms(state)[0] || ''; render();
    } catch (error) { notify(error.message, true); }
  });
}

function shell(body) {
  app.innerHTML = `<div class="shell"><header class="topbar"><div class="toprow"><div class="brand">辅导员工作台</div><button class="icon-btn" data-action="lock" aria-label="锁定">⌁</button></div></header><main class="content">${body}</main><nav class="tabbar five-tabs" aria-label="主导航">${[['home','⌂','首页'],['students','◫','学生'],['work','▦','工作'],['alerts','✓','待办'],['data','⚙','设置']].map(([key, icon, label]) => `<button data-view="${key}" class="${view === key ? 'active' : ''}"><span class="tab-icon">${icon}</span><span>${label}</span></button>`).join('')}</nav></div>`;
  if (photoUrl) { URL.revokeObjectURL(photoUrl); photoUrl = ''; }
  if (view === 'students' && selectedId) loadPhoto(selectedId);
}

function homeView() {
  const total = Object.keys(state.students).length, threshold = Number(state.settings.threshold || 1);
  const warnings = selectedTerm ? Object.keys(state.students).filter(id => gradeStats(state, id, selectedTerm).semester >= threshold) : [];
  const due = state.records.filter(r => r.todo && !r.done && r.dueDate && r.dueDate <= today()).sort((a,b) => a.dueDate.localeCompare(b.dueDate));
  const recent = [...state.records].sort((a,b) => (b.date || '').localeCompare(a.date || '')).slice(0, 4), incomplete = Object.values(state.students).filter(s => profileMissing(s).length);
  return `<div class="hero"><div class="kicker">今日概览</div><strong>${total} 名学生</strong></div><div class="stats"><button class="stat" data-view="students"><strong>${total}</strong><span>学生档案</span></button><button class="stat warn" data-view="work"><strong>${warnings.length}</strong><span>学业预警</span></button><button class="stat due" data-view="alerts"><strong>${due.length}</strong><span>到期待办</span></button></div><div class="section-title"><h2>优先跟进</h2></div><div class="panel list">${due.length ? due.slice(0, 4).map(r => `<button class="list-row" data-student="${esc(r.studentId)}"><div class="avatar">◷</div><div class="row-main"><strong>${esc(studentName(r.studentId))} · ${esc(r.todo)}</strong><small>到期 ${fmtDate(r.dueDate)} · ${esc(r.type)}</small></div><span class="badge red">待办</span></button>`).join('') : empty('暂无待跟进事项')}</div><div class="section-title"><h2>资料待补</h2><small>${incomplete.length} 人</small></div><div class="panel list">${incomplete.length ? incomplete.slice(0,3).map(s => `<button class="list-row" data-student="${esc(s.id)}"><div class="avatar">${esc(studentInitial(s))}</div><div class="row-main"><strong>${esc(studentName(s.id))}</strong><small>缺少：${esc(profileMissing(s).join('、'))}</small></div><span class="chevron">›</span></button>`).join('') : empty('资料已完整')}</div><div class="section-title"><h2>最近记录</h2><small>${state.records.length} 条</small></div><div class="panel list">${recent.length ? recent.map(r => `<button class="list-row" data-student="${esc(r.studentId)}"><div class="avatar">✎</div><div class="row-main"><strong>${esc(studentName(r.studentId))} · ${esc(r.type)}</strong><small>${fmtDate(r.date)} · ${esc(r.summary || '无摘要')}</small></div><span class="chevron">›</span></button>`).join('') : empty('还没有工作记录')}</div>`;
}
function empty(title) { return `<div class="empty"><strong>${esc(title)}</strong></div>`; }
function studentsView() {
  if (selectedId && student()) return detailView();
  const query = search.toLowerCase(), filtered = Object.values(state.students).filter(s => [s.id,s.name,s.className,s.major,s.dorm].some(v => clean(v).toLowerCase().includes(query))).sort((a,b) => clean(a.className).localeCompare(clean(b.className),'zh-CN') || clean(a.name).localeCompare(clean(b.name),'zh-CN'));
  return `<div class="toolbar"><h2>学生档案 <span class="muted tiny">${filtered.length}</span></h2><button class="btn small" data-action="add-student">＋ 新增</button></div><div class="search-wrap"><input class="search" id="student-search" type="search" placeholder="搜索姓名、学号、班级、专业" value="${esc(search)}"></div><div class="panel list">${filtered.length ? filtered.map(s => `<button class="list-row" data-student="${esc(s.id)}"><div class="avatar">${esc(studentInitial(s))}</div><div class="row-main"><strong>${esc(displayName(s.name))}</strong><small>${esc(s.className || '未分班')} · ${esc(s.id)} · ${esc(s.major)}</small></div><span class="chevron">›</span></button>`).join('') : empty(search ? '没有匹配的学生' : '还没有学生档案', search ? '试试学号或班级关键词。' : '可手工新增，或从固定 Excel 模板批量导入。')}</div>`;
}
const dialNumber = value => clean(value).replace(/[^\d+]/g, '');
const profileMissing = student => ['姓名','班级','专业','宿舍','本人电话','家长电话','家庭住址','身份证号','证件照'].filter(label => ({ 姓名: student.name, 班级: student.className, 专业: student.major, 宿舍: student.dorm, 本人电话: student.phone, 家长电话: student.parentPhone, 家庭住址: student.address, 身份证号: student.idNumber, 证件照: student.photoId })[label] ? false : true);
function infoItem(label, value) {
  const phone = ['本人电话', '家长电话'].includes(label) && dialNumber(value);
  const content = phone ? `<a class="phone-link" href="tel:${esc(phone)}">${esc(value)}</a>` : esc(value || '—');
  return `<div class="info-item"><small>${esc(label)}</small><strong>${content}</strong></div>`;
}
function detailView() {
  const s = student(); const grades = state.grades.filter(g => g.studentId === s.id), funding = state.funding.filter(f => f.studentId === s.id), records = state.records.filter(r => r.studentId === s.id).sort((a,b) => (b.date || '').localeCompare(a.date || ''));
  const tabs = [['info','信息'],['grades','成绩'],['funding','资助'],['records','记录']];
  let body = '';
  if (detailTab === 'info') {
    const fields = [['学号',s.id],['姓名',s.name],['班级',s.className],['专业',s.major],['宿舍',s.dorm],['本人电话',s.phone],['家长电话',s.parentPhone],['家庭住址',s.address],['性别',s.gender],['民族',s.ethnicity],['出生日期',s.birthDate]];
    body = `<div class="panel pad"><div class="info-grid">${fields.map(([label,value]) => infoItem(label,value)).join('')}<div class="info-item"><small>身份证号</small><strong id="id-number">${maskId(s.idNumber)}</strong><button class="btn small ghost" data-action="toggle-id" style="margin-top:7px">显示 / 隐藏</button></div>${Object.entries(s.custom || {}).map(([k,v]) => infoItem(k,v)).join('')}</div></div><div class="btn-row" style="margin-top:12px"><button class="btn secondary" data-action="edit-student">编辑档案</button><button class="btn ghost" data-action="photo">上传证件照</button></div>`;
  } else if (detailTab === 'grades') {
    const stats = gradeStats(state,s.id,selectedTerm);
    body = `<div class="stats"><div class="stat warn"><strong>${stats.semester}</strong><span>本学期挂科</span></div><div class="stat"><strong>${stats.cumulative}</strong><span>历史累计</span></div><div class="stat"><strong>${stats.resolved}</strong><span>后来已通过</span></div></div><div class="section-title"><h2>成绩记录</h2><button class="btn small" data-action="add-grade">＋ 添加</button></div><div class="panel">${grades.length ? grades.sort((a,b) => b.term.localeCompare(a.term)).map(g => `<div class="record"><div class="record-top"><strong>${esc(courseLabel(g))}</strong><span class="badge ${isFailed(g) ? 'red' : 'green'}">${isFailed(g) ? '曾挂科' : '及格'}</span></div><p>${esc(g.term)} · 成绩 ${esc(g.score || '—')} ${g.retakeStatus ? `· ${esc(g.retakeStatus)}` : ''}</p><div class="actions"><button class="btn small ghost" data-action="edit-grade" data-id="${esc(g.id)}">编辑</button></div></div>`).join('') : empty('暂无成绩', '可手工添加或导入 Excel。')}</div>`;
  } else if (detailTab === 'funding') {
    const total = funding.reduce((sum,f) => sum + (Number(f.amount) || 0),0);
    body = `<div class="hero"><div class="kicker">累计发放金额</div><strong>¥ ${total.toLocaleString('zh-CN')}</strong><p>基于已录入的资助记录</p></div><div class="section-title"><h2>资助记录</h2><button class="btn small" data-action="add-funding">＋ 添加</button></div><div class="panel">${funding.length ? funding.map(f => `<div class="record"><div class="record-top"><strong>${esc(f.program || '资助项目')}</strong><span class="badge">¥ ${esc(f.amount || '0')}</span></div><p>${fmtDate(f.paidAt)} · 困难等级 ${esc(f.hardship || '未填')}</p>${f.family ? `<p>家庭情况：${esc(f.family)}</p>` : ''}<div class="actions"><button class="btn small ghost" data-action="edit-funding" data-id="${esc(f.id)}">编辑</button></div></div>`).join('') : empty('暂无资助记录', '可记录困难等级、家庭情况和历次资助。')}</div>`;
  } else {
    body = `<div class="toolbar"><h2>工作记录</h2><button class="btn small" data-action="add-record">＋ 记录</button></div><div class="panel">${records.length ? records.map(r => `<div class="record"><div class="record-top"><strong>${esc(r.type)} · ${fmtDate(r.date)}</strong><span class="badge ${r.done ? 'green' : r.todo ? 'red' : 'gray'}">${r.done ? '已完成' : r.todo ? '待跟进' : '已记录'}</span></div><p>对象：${esc(r.subject || '—')}</p><p>${esc(r.summary || '无摘要')}</p>${r.todo ? `<p>后续：${esc(r.todo)} · 到期 ${fmtDate(r.dueDate)}</p>` : ''}<div class="actions"><button class="btn small ghost" data-action="edit-record" data-id="${esc(r.id)}">编辑</button>${r.todo ? `<button class="btn small secondary" data-action="toggle-done" data-id="${esc(r.id)}">${r.done ? '恢复待办' : '标记完成'}</button>` : ''}${(r.attachments || []).map((id,i) => `<button class="btn small ghost" data-action="open-file" data-id="${esc(id)}">截图 ${i+1}</button>`).join('')}</div></div>`).join('') : empty('暂无工作记录', '谈话和家校联系可在这里记录并附截图。')}</div>`;
  }
  return `<button class="back" data-action="back-students">← 返回学生列表</button><div class="panel detail-head"><div class="avatar" id="detail-avatar">${esc(studentInitial(s))}</div><div><h2>${esc(displayName(s.name))}</h2><p>${esc(s.className || '未分班')} · ${esc(s.major || '未填专业')}</p><span class="pill">学号 ${esc(s.id)}</span></div></div><div class="segmented">${tabs.map(([key,label]) => `<button data-detail-tab="${key}" class="${detailTab === key ? 'active' : ''}">${label}</button>`).join('')}</div>${body}`;
}
async function loadPhoto(id) {
  const s = state.students[id]; if (!s?.photoId) return;
  try { const file = await getFile(s.photoId); if (!file || selectedId !== id) return; photoUrl = URL.createObjectURL(file.blob); const avatar = document.querySelector('#detail-avatar'); if (avatar) avatar.innerHTML = `<img src="${photoUrl}" alt="${esc(s.name)}证件照">`; }
  catch (error) { notify(error.message,true); }
}
function alertsView() {
  const all = state.records.filter(r => r.todo), due = all.filter(r => todoFilter === 'done' ? r.done : todoFilter === 'overdue' ? !r.done && r.dueDate && r.dueDate < today() : todoFilter === 'today' ? !r.done && r.dueDate === today() : !r.done).sort((a,b) => (a.dueDate || '9999').localeCompare(b.dueDate || '9999'));
  return `<div class="toolbar"><h2>待办</h2><small class="muted">${due.length} 项</small></div><div class="segmented todo-filter">${[['all','待处理'],['today','今日'],['overdue','已逾期'],['done','已完成']].map(([key,label]) => `<button data-action="todo-filter" data-filter="${key}" class="${todoFilter === key ? 'active' : ''}">${label}</button>`).join('')}</div><div class="panel list">${due.length ? due.map(r => `<button class="list-row" data-student="${esc(r.studentId)}"><div class="avatar">✓</div><div class="row-main"><strong>${esc(studentName(r.studentId))} · ${esc(r.todo)}</strong><small>${r.dueDate ? `到期 ${fmtDate(r.dueDate)}` : '未设到期日'} · ${esc(r.type)}</small></div><span class="badge ${r.done ? 'green' : r.dueDate && r.dueDate < today() ? 'red' : ''}">${r.done ? '已完成' : r.dueDate && r.dueDate < today() ? '已逾期' : '待办'}</span></button>`).join('') : empty('暂无待办')}</div>`;
}

const workModules = [
  ['organization', '◇', '组织发展'],
  ['grades', '≋', '学业成绩'],
  ['funding', '¥', '资助工作'],
  ['mental', '♡', '心理工作'],
  ['contact', '⌁', '家校联系']
];
function workView() {
  if (workModule) return workDetailView();
  return `<div class="toolbar"><h2>工作</h2></div><div class="work-grid">${workModules.map(([key, icon, label]) => `<button class="work-card" data-action="work-module" data-module="${key}"><span>${icon}</span><strong>${label}</strong><i>›</i></button>`).join('')}</div>`;
}
function workRecordList(type, emptyTitle) {
  const query = clean(workSearch).toLowerCase(), records = state.records.filter(record => record.type === type && [studentName(record.studentId), state.students[record.studentId]?.className, record.date, record.summary, record.subject].some(value => clean(value).toLowerCase().includes(query))).sort((a,b) => (b.date || '').localeCompare(a.date || ''));
  return `<div class="panel">${records.length ? records.map(record => `<button class="list-row" data-student="${esc(record.studentId)}"><div class="avatar">${type === '心理工作' ? '♡' : type === '组织发展' ? '◇' : '⌁'}</div><div class="row-main"><strong>${esc(studentName(record.studentId))}</strong><small>${fmtDate(record.date)} · ${esc(record.summary || '无摘要')}</small></div><span class="chevron">›</span></button>`).join('') : empty(emptyTitle)}</div>`;
}
function workDetailView() {
  const module = workModules.find(([key]) => key === workModule);
  if (!module) { workModule = ''; return workView(); }
  const [, , title] = module;
  let body = '';
  if (workModule === 'grades') {
    const query = clean(workSearch).toLowerCase(), threshold = Number(state.settings.threshold || 1), warnings = selectedTerm ? Object.values(state.students).map(s => ({ s, count: gradeStats(state,s.id,selectedTerm).semester })).filter(item => item.count >= threshold && [studentName(item.s.id), item.s.className, item.s.major].some(value => clean(value).toLowerCase().includes(query))).sort((a,b) => b.count - a.count) : [];
    body = `<div class="notice">本学期挂科达到 <strong>${threshold} 门</strong>时提醒。</div><div class="section-title"><h2>学业预警</h2><select id="term-select" class="mini-select"><option value="">选择学期</option>${terms(state).map(term => `<option value="${esc(term)}" ${term === selectedTerm ? 'selected' : ''}>${esc(term)}</option>`).join('')}</select></div><div class="panel list">${warnings.length ? warnings.map(({s,count}) => `<button class="list-row" data-student="${esc(s.id)}"><div class="avatar">${esc(studentInitial(s))}</div><div class="row-main"><strong>${esc(studentName(s.id))}</strong><small>${esc(s.className)} · 历史累计 ${gradeStats(state,s.id,selectedTerm).cumulative} 门</small></div><span class="badge red">${count} 门</span></button>`).join('') : empty('暂无学业预警')}</div>`;
  } else if (workModule === 'funding') {
    const query = clean(workSearch).toLowerCase(), rows = state.funding.filter(item => [studentName(item.studentId), state.students[item.studentId]?.className, item.paidAt, item.program].some(value => clean(value).toLowerCase().includes(query))).sort((a,b) => (b.paidAt || '').localeCompare(a.paidAt || ''));
    body = `<div class="panel list">${rows.length ? rows.map(item => `<button class="list-row" data-student="${esc(item.studentId)}"><div class="avatar">¥</div><div class="row-main"><strong>${esc(studentName(item.studentId))} · ${esc(item.program || '资助项目')}</strong><small>${fmtDate(item.paidAt)} · ¥ ${esc(item.amount || '0')}</small></div><span class="chevron">›</span></button>`).join('') : empty('暂无资助记录')}</div>`;
  } else if (workModule === 'organization') body = workRecordList('组织发展', '暂无组织发展记录');
  else if (workModule === 'mental') body = workRecordList('心理工作', '暂无心理工作记录');
  else body = workRecordList('家校联系', '暂无家校联系记录');
  const action = workModule === 'grades' ? 'work-add-grade' : workModule === 'funding' ? 'work-add-funding' : 'work-add-record';
  return `<button class="back" data-action="back-work">← 返回工作</button><div class="toolbar"><h2>${esc(title)}</h2><button class="btn small" data-action="${action}" data-type="${esc(title)}">＋ 新增</button></div><div class="search-wrap"><input class="search" id="work-search" type="search" placeholder="搜索学生、班级、日期或内容" value="${esc(workSearch)}"></div>${body}`;
}
function dataView() {
  const current = releases[0];
  return `<div class="toolbar"><h2>设置</h2></div><div class="section-title"><h2>本机解锁</h2></div><div class="panel"><div class="file-card"><h3>6 位数字密码</h3><button class="btn ghost" data-action="change-passcode">修改解锁密码</button></div></div><div class="section-title"><h2>Excel 模板</h2></div><div class="panel"><div class="file-card"><h3>空白固定模板</h3><button class="btn secondary" data-action="template">下载空白模板</button></div><div class="file-card"><h3>导入 Excel</h3><button class="btn secondary" data-action="import-xlsx">选择 Excel 文件</button></div><div class="file-card"><h3>导出当前数据</h3><button class="btn secondary" data-action="export-xlsx">导出 Excel</button></div></div><div class="section-title"><h2>照片与传输</h2></div><div class="panel"><div class="file-card"><h3>证件照批量导入</h3><button class="btn ghost" data-action="import-photos">选择照片 ZIP</button></div><div class="file-card"><h3>导出加密数据包</h3><p>上次备份：${state.settings.lastExportedAt ? esc(state.settings.lastExportedAt.replace('T',' ').slice(0,16)) : '尚未备份'}</p><button class="btn" data-action="export-package">生成传输包 / 备份</button></div><div class="file-card"><h3>导入加密数据包</h3><button class="btn ghost" data-action="import-package">选择数据包 ZIP</button></div></div><div class="section-title"><h2>版本</h2></div><button class="version-card" data-action="version-history"><span><strong>v${esc(APP_VERSION)}</strong><small>更新于 ${esc(current.updatedAt)}</small></span><span class="chevron">›</span></button>`;
}

function versionHistory() {
  showModal('历史更新记录', `<div class="release-list">${releases.map(release => `<section class="release"><div><strong>v${esc(release.version)}</strong><small>${esc(release.updatedAt)}</small></div><ul>${release.notes.map(note => `<li>${esc(note)}</li>`).join('')}</ul></section>`).join('')}</div>`, '关闭', async () => {});
}
function render() {
  if (!state) return gate();
  const body = view === 'home' ? homeView() : view === 'students' ? studentsView() : view === 'work' ? workView() : view === 'alerts' ? alertsView() : dataView();
  shell(body);
  const searchInput = document.querySelector('#student-search');
  if (searchInput) searchInput.addEventListener('input', event => { const pos = event.target.selectionStart; search = event.target.value; render(); const input = document.querySelector('#student-search'); input.focus(); input.setSelectionRange(pos,pos); });
  const workInput = document.querySelector('#work-search');
  if (workInput) workInput.addEventListener('input', event => { workSearch = event.target.value; render(); document.querySelector('#work-search')?.focus(); });
  document.querySelector('#term-select')?.addEventListener('change', event => { selectedTerm = event.target.value; render(); });
}

function studentForm(existing) {
  const s = existing || {};
  showModal(existing ? '编辑学生档案' : '新增学生', `${field('学号 *','id',s.id,'text',existing ? 'readonly required' : 'required')}${field('姓名','name',s.name)}${field('班级','className',s.className)}${field('专业','major',s.major)}${field('宿舍','dorm',s.dorm)}${field('本人电话','phone',s.phone,'tel')}${field('家长电话','parentPhone',s.parentPhone,'tel')}${area('家庭住址','address',s.address)}${field('身份证号','idNumber',s.idNumber)}${choose('性别','gender',['','男','女','其他'],s.gender)}${field('民族','ethnicity',s.ethnicity)}${field('出生日期','birthDate',s.birthDate,'date')}${area('自定义字段（每行：字段名=内容）','custom',Object.entries(s.custom || {}).map(([k,v]) => `${k}=${v}`).join('\n'))}<p class="inline-note">学号用于关联其他数据，必须填写；其余字段可留空。有效的 18 位身份证号可自动生成出生日期。</p>`, '保存档案', async form => {
    const id = clean(form.get('id')), name = clean(form.get('name')); if (!id) throw new Error('请填写学号');
    const custom = {}; for (const line of String(form.get('custom') || '').split(/\r?\n/)) { const at = line.indexOf('='); if (at > 0) custom[clean(line.slice(0,at))] = clean(line.slice(at+1)); }
    const idNumber = clean(form.get('idNumber'));
    state.students[id] = { ...s, id, name, className: clean(form.get('className')), major: clean(form.get('major')), dorm: clean(form.get('dorm')), phone: clean(form.get('phone')), parentPhone: clean(form.get('parentPhone')), address: clean(form.get('address')), idNumber, gender: clean(form.get('gender')), ethnicity: clean(form.get('ethnicity')), birthDate: clean(form.get('birthDate')) || dateFromId(idNumber), custom, updatedAt: now() };
    selectedId = id; detailTab = 'info'; view = 'students'; await persist(); notify('档案已保存');
  });
}
function pickStudent(next) { showModal('选择学生', choose('学生','studentId',['', ...Object.values(state.students).sort((a,b) => studentName(a.id).localeCompare(studentName(b.id),'zh-CN')).map(s => `${s.id}｜${studentName(s.id)}`)], ''), '下一步', async form => { const id = clean(form.get('studentId')).split('｜')[0]; if (!state.students[id]) throw new Error('请选择学生'); selectedId = id; setTimeout(next, 0); }); }
function gradeForm(existing) {
  const g = existing || {};
  showModal(existing ? '编辑成绩' : '添加成绩', `${field('学期 *','term',g.term || selectedTerm,'text','required placeholder="例如 2025-2026-2"')}${field('课程代码','courseCode',g.courseCode)}${field('课程名称 *','courseName',g.courseName,'text','required')}${field('成绩','score',g.score,'text','inputmode="decimal"')}${choose('是否挂科','failed',['自动判断','是','否'],g.failed || '自动判断')}${choose('补考或重修状态','retakeStatus',['未处理','补考中','重修中','补考通过','重修通过','已通过'],g.retakeStatus || '未处理')}${area('备注','note',g.note)}`, '保存成绩', async form => {
    const next = { ...g, id: g.id || uid(), studentId: selectedId, term: clean(form.get('term')), courseCode: clean(form.get('courseCode')), courseName: clean(form.get('courseName')), score: clean(form.get('score')), failed: form.get('failed') === '自动判断' ? '' : form.get('failed'), retakeStatus: clean(form.get('retakeStatus')), note: clean(form.get('note')), updatedAt: now() };
    if (existing) Object.assign(existing,next); else state.grades.push(next);
    if (!selectedTerm) selectedTerm = clean(form.get('term')); await persist(); notify('成绩已添加');
  });
}
function fundingForm(existing) {
  const f = existing || {};
  showModal(existing ? '编辑资助记录' : '添加资助记录', `${choose('困难等级','hardship',['','特别困难','困难','一般困难','其他'],f.hardship)}${area('家庭情况','family',f.family)}${field('资助项目 *','program',f.program,'text','required')}${field('发放日期','paidAt',f.paidAt || today(),'date')}${field('发放金额','amount',f.amount,'number','min="0" step="0.01"')}${area('备注','note',f.note)}`, '保存资助', async form => {
    const next = { ...f, id: f.id || uid(), studentId: selectedId, hardship: clean(form.get('hardship')), family: clean(form.get('family')), program: clean(form.get('program')), paidAt: clean(form.get('paidAt')), amount: clean(form.get('amount')), note: clean(form.get('note')), updatedAt: now() };
    if (existing) Object.assign(existing,next); else state.funding.push(next);
    await persist(); notify('资助记录已保存');
  });
}
function recordForm(existing) {
  const r = existing || {};
  showModal(existing ? '编辑工作记录' : '添加工作记录', `${choose('类型','type',['谈话','家校联系','组织发展','心理工作'],r.type || (workModule === 'organization' ? '组织发展' : workModule === 'mental' ? '心理工作' : workModule === 'contact' ? '家校联系' : '谈话'))}${field('日期','date',r.date || today(),'date','required')}${field('对象','subject',r.subject,'text','placeholder="例如 学生本人 / 母亲"')}${field('关注等级','attention',r.attention)}${field('来源','source',r.source)}${area('摘要 *','summary',r.summary)}${area('后续待办','todo',r.todo)}${area('处理结论','outcome',r.outcome)}${field('到期日期','dueDate',r.dueDate,'date')}${field('追加截图（可多选）','images','','file','accept="image/*" multiple')}`, '保存记录', async (form,data) => {
    const summary = clean(form.get('summary')); if (!summary) throw new Error('请填写摘要');
    const id = uid(), date = clean(form.get('date')), type = clean(form.get('type')), s = student();
    const images = [...data.querySelector('[name=images]').files]; const attachments = [...(r.attachments || [])];
    for (const [index,image] of images.entries()) {
      const ext = image.name.split('.').pop()?.toLowerCase() || 'jpg', fileId = uid();
      await putFile({ id: fileId, studentId: s.id, kind: 'screenshot', name: fileName(s, `${type}截图`, date, attachments.length+1, ext), type: image.type || 'image/jpeg', createdAt: now(), blob: image });
      attachments.push(fileId);
    }
    const next = { ...r, id: r.id || id, studentId: s.id, type, date, subject: clean(form.get('subject')), attention: clean(form.get('attention')), source: clean(form.get('source')), summary, todo: clean(form.get('todo')), outcome: clean(form.get('outcome')), dueDate: clean(form.get('dueDate')), done: r.done || false, attachments, updatedAt: now() };
    if (existing) Object.assign(existing,next); else state.records.push(next);
    await persist(); notify('工作记录已保存');
  });
}
function photoForm() {
  showModal('上传证件照', `${field('选择照片','photo','','file','accept="image/*" required')}<p class="inline-note">建议使用 JPG 或 PNG。照片将加密保存在本机。</p>`, '保存照片', async (_,form) => {
    const image = form.querySelector('[name=photo]').files[0], s = student(); if (!image) throw new Error('请选择照片');
    const id = uid(), compressed = await normalizedPhoto(image), isJpeg = compressed.type === 'image/jpeg', ext = isJpeg ? 'jpg' : image.name.split('.').pop()?.toLowerCase() || 'jpg';
    await putFile({ id, studentId: s.id, kind: 'photo', name: fileName(s,'证件照',today(),1,ext), type: compressed.type || image.type || 'image/jpeg', createdAt: now(), blob: compressed });
    s.photoId = id; s.updatedAt = now(); await persist(); notify('证件照已保存');
  });
}
function thresholdForm() { showModal('设置挂科预警', `${field('本学期挂科达到几门时提醒','threshold',state.settings.threshold,'number','required min="1" max="30"')}`, '保存规则', async form => { state.settings.threshold = Math.max(1, Math.min(30, Number(form.get('threshold')) || 1)); await persist(); notify('预警规则已更新'); }); }
function passcodeForm() {
  showModal('修改解锁密码', `${field('当前解锁密码','currentPassword','','password','required autocomplete="current-password"')}${field('新 6 位数字密码','nextPassword','','password','required inputmode="numeric" pattern="[0-9]{6}" maxlength="6" autocomplete="new-password"')}${field('再次输入新密码','confirmPassword','','password','required inputmode="numeric" pattern="[0-9]{6}" maxlength="6" autocomplete="new-password"')}`, '保存并重新加密', async form => {
    const next = String(form.get('nextPassword')); if (next !== form.get('confirmPassword')) throw new Error('两次输入的新密码不一致');
    await changeSecurity(String(form.get('currentPassword')), next); notify('解锁密码已改为 6 位数字');
  });
}
function passwordForm(title, text, onPassword) { showModal(title, `${text ? `<p class="notice">${esc(text)}</p>` : ''}${field('数据包密码','password','','password','required minlength="8" autocomplete="off"')}`, '继续', async form => onPassword(String(form.get('password')))); }

async function importXlsx(file) {
  const workbook = await readWorkbook(file), preview = previewWorkbook(state, workbook);
  const c = preview.counts;
  showModal('确认 Excel 导入', `<div class="notice">新增学生 ${c.students} 人、成绩 ${c.grades} 条、资助 ${c.funding} 条、工作记录 ${c.records} 条；更新 ${c.updated} 条。</div>${preview.errors.length ? `<p class="inline-note">发现 ${preview.errors.length} 行错误，这些行会跳过：</p><div class="panel pad tiny" style="max-height:160px;overflow:auto">${preview.errors.slice(0,20).map(esc).join('<br>')}</div><button type="button" class="btn small ghost" id="download-import-errors" style="margin-top:10px">下载错误清单</button>` : '<p class="tiny muted">未发现格式错误。</p>'}`, '确认导入', async () => { state = preview.next; if (!selectedTerm) selectedTerm = terms(state)[0] || ''; await persist(); notify('Excel 已导入'); });
  modalRoot.querySelector('#download-import-errors')?.addEventListener('click', () => download(new Blob([`\uFEFF问题\n${preview.errors.map(error => `"${error.replaceAll('"','""')}"`).join('\n')}`], { type: 'text/csv;charset=utf-8' }), `Excel导入错误_${stamp()}.csv`));
}
async function importPhotoZip(file) {
  if (file.size > 50 * 1024 * 1024) throw new Error('照片 ZIP 超过 50 MB，请分成多个压缩包导入');
  const zip = unzipSync(new Uint8Array(await file.arrayBuffer()));
  const matches = [], unknown = [], duplicate = [];
  const seen = new Set();
  for (const [path, bytes] of Object.entries(zip)) {
    if (path.endsWith('/')) continue;
    const name = path.split('/').pop(), m = name.match(/^(.+)\.(jpe?g|png|webp)$/i); if (!m) continue;
    const id = m[1]; if (seen.has(id)) { duplicate.push(name); continue; } seen.add(id);
    if (!state.students[id]) unknown.push(name); else matches.push({ studentId: id, bytes, ext: m[2].toLowerCase() });
  }
  showModal('确认照片导入', `<div class="notice">匹配学生 ${matches.length} 张；无对应学号 ${unknown.length} 张；重复学号 ${duplicate.length} 张。</div>${unknown.length || duplicate.length ? `<p class="tiny muted">未匹配：${esc([...unknown,...duplicate].slice(0,12).join('、'))}</p>` : ''}`, '确认导入', async () => {
    for (const item of matches) { const s = state.students[item.studentId], id = uid(), sourceType = item.ext === 'png' ? 'image/png' : item.ext === 'webp' ? 'image/webp' : 'image/jpeg';
      const photo = await normalizedPhoto(new Blob([item.bytes], { type: sourceType })); const ext = photo.type === 'image/jpeg' ? 'jpg' : item.ext;
      await putFile({ id, studentId: s.id, kind: 'photo', name: fileName(s,'证件照',today(),1,ext), type: photo.type || sourceType, createdAt: now(), blob: photo }); s.photoId = id; s.updatedAt = now(); }
    await persist(); notify(`已导入 ${matches.length} 张证件照`);
  });
}
async function exportEncrypted(password) {
  const blob = await makePackage(state, await allFiles(), password), name = packageName();
  state.settings.lastExportedAt = now(); await persist();
  closeModal();
  modalRoot.innerHTML = `<div class="modal-shade"><div class="modal"><div class="modal-head"><h2>数据包已生成</h2><button class="close" data-close>×</button></div><p class="notice">${esc(name)}<br>文件名包含应用版本号和导出日期时间。请妥善保管密码。</p><div class="modal-actions"><button class="btn ghost" data-close>关闭</button><button class="btn" id="save-package">保存 / 分享 ZIP</button></div></div></div>`;
  modalRoot.querySelectorAll('[data-close]').forEach(button => button.addEventListener('click',closeModal));
  modalRoot.querySelector('#save-package').addEventListener('click', async () => {
    const file = new File([blob], name, { type: 'application/zip' });
    try { if (navigator.canShare?.({ files: [file] })) await navigator.share({ files: [file], title: name }); else download(blob,name); }
    catch (error) { if (error.name !== 'AbortError') { download(blob,name); notify('已改用浏览器下载'); } }
  });
}
async function importEncrypted(file,password) {
  const incoming = await readPackage(file,password), result = mergePackage(state,incoming.state), existing = new Set((await allFiles()).map(f => f.id)), newFiles = incoming.files.filter(f => !existing.has(f.id));
  showModal('确认数据包导入', `<div class="notice">导出时间：${esc(incoming.manifest.exportedAt)}<br>新增记录 ${result.added} 条，更新较旧记录 ${result.updated} 条，新增附件 ${newFiles.length} 个。</div><p class="tiny muted">本机较新的记录会保留。相同记录以更新时间较新的版本为准。</p>`, '确认合并', async () => { await putFiles(newFiles); state = result.merged; selectedTerm = terms(state)[0] || selectedTerm; await persist(); notify('数据包已合并'); });
}

app.addEventListener('click', async event => {
  const viewButton = event.target.closest('[data-view]'); if (viewButton) { selectedId = ''; setView(viewButton.dataset.view); return; }
  const row = event.target.closest('[data-student]'); if (row) { selectedId = row.dataset.student; detailTab = 'info'; setView('students'); return; }
  const tab = event.target.closest('[data-detail-tab]'); if (tab) { detailTab = tab.dataset.detailTab; render(); return; }
  const button = event.target.closest('[data-action]'); if (!button) return;
  try {
    switch(button.dataset.action) {
      case 'lock': lockSecurity(); state = undefined; selectedId = ''; gate(); break;
      case 'back-students': selectedId = ''; render(); break;
      case 'back-work': workModule = ''; render(); break;
      case 'work-module': workModule = button.dataset.module; render(); window.scrollTo(0, 0); break;
      case 'todo-filter': todoFilter = button.dataset.filter; render(); break;
      case 'work-add-grade': pickStudent(() => gradeForm()); break;
      case 'work-add-funding': pickStudent(() => fundingForm()); break;
      case 'work-add-record': pickStudent(() => recordForm()); break;
      case 'add-student': studentForm(); break;
      case 'edit-student': studentForm(student()); break;
      case 'add-grade': gradeForm(); break;
      case 'edit-grade': gradeForm(state.grades.find(g => g.id === button.dataset.id)); break;
      case 'add-funding': fundingForm(); break;
      case 'edit-funding': fundingForm(state.funding.find(f => f.id === button.dataset.id)); break;
      case 'add-record': recordForm(); break;
      case 'edit-record': recordForm(state.records.find(r => r.id === button.dataset.id)); break;
      case 'photo': photoForm(); break;
      case 'threshold': thresholdForm(); break;
      case 'change-passcode': passcodeForm(); break;
      case 'version-history': versionHistory(); break;
      case 'toggle-id': { const el = document.querySelector('#id-number'); el.textContent = el.textContent.includes('*') ? student().idNumber || '—' : (student().idNumber ? `${student().idNumber.slice(0,4)}**********${student().idNumber.slice(-4)}` : '—'); break; }
      case 'toggle-done': { const record = state.records.find(r => r.id === button.dataset.id); if (record) { record.done = !record.done; record.updatedAt = now(); await persist(); } break; }
      case 'open-file': { const file = await getFile(button.dataset.id); if (!file) throw new Error('附件不存在'); const url = URL.createObjectURL(file.blob); modalRoot.innerHTML = `<div class="modal-shade"><div class="modal"><div class="modal-head"><h2>${esc(file.name)}</h2><button class="close" data-close>×</button></div><img class="photo-preview" src="${url}" alt="截图"><div class="modal-actions"><button class="btn ghost" data-close>关闭</button><a class="btn" href="${url}" download="${esc(file.name)}" style="text-align:center;text-decoration:none">保存图片</a></div></div></div>`; modalRoot.querySelectorAll('[data-close]').forEach(b => b.addEventListener('click', () => { closeModal(); URL.revokeObjectURL(url); })); break; }
      case 'template': download(blankTemplate(), `辅导员工作台_空白模板_v${APP_VERSION}.xlsx`); break;
      case 'export-xlsx': download(createWorkbook(exportSheets(state)), `辅导员工作台_数据导出_v${APP_VERSION}_${stamp()}.xlsx`); break;
      case 'import-xlsx': inputs.xlsx.click(); break;
      case 'import-photos': inputs.photozip.click(); break;
      case 'export-package': passwordForm('导出加密数据包','请设置本次传输包密码，至少 8 位。',exportEncrypted); break;
      case 'import-package': inputs.package.click(); break;
    }
  } catch (error) { notify(error.message || String(error), true); }
});

inputs.xlsx.addEventListener('change', async event => { const file = event.target.files[0]; event.target.value = ''; if (!file) return; try { await importXlsx(file); } catch(error) { notify(error.message,true); } });
inputs.photozip.addEventListener('change', async event => { const file = event.target.files[0]; event.target.value = ''; if (!file) return; try { await importPhotoZip(file); } catch(error) { notify(error.message,true); } });
inputs.package.addEventListener('change', async event => { const file = event.target.files[0]; event.target.value = ''; if (!file) return; passwordForm('导入加密数据包','输入导出时设置的数据包密码。',password => importEncrypted(file,password)); });

let hiddenAt = 0;
document.addEventListener('visibilitychange', () => {
  if (document.hidden) hiddenAt = Date.now();
  else if (state && hiddenAt && Date.now() - hiddenAt > 5 * 60 * 1000) { lockSecurity(); state = undefined; selectedId = ''; closeModal(); gate(); }
});
if ('serviceWorker' in navigator && location.protocol !== 'file:') navigator.serviceWorker.register('./sw.js').catch(() => {});
gate().catch(error => { app.innerHTML = `<div class="empty"><strong>应用无法启动</strong><p>${esc(error.message)}</p></div>`; });
