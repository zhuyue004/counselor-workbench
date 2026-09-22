import { APP_VERSION, uid, now, clean, dateFromId, gradeStats, terms, isFailed, emptyState, hasSecurity, unlockSecurity, loadState, openLocalFiles, putFile, getFile, allFiles, putFiles, previewWorkbook, exportSheets } from './data.js';
import { COLUMNS, blankTemplate, createWorkbook, readWorkbook } from './xlsx.js';
import { makePackage, readPackage, mergePackage, packageName, stamp } from './transfer.js';
import { unzipSync } from './vendor/fflate.js';
import { account, restoreAccount, signIn, register, sendResetEmail, signOut, loadCloudState, saveCloudState } from './firebase.js';

const app = document.querySelector('#app');
const modalRoot = document.querySelector('#modal-root');
const toastEl = document.querySelector('#toast');
const inputs = { xlsx: document.querySelector('#xlsx-input'), photozip: document.querySelector('#photozip-input'), package: document.querySelector('#package-input') };
let state, view = 'home', selectedId = '', detailTab = 'info', search = '', selectedTerm = '', workModule = '', workSearch = '', todoFilter = 'all', classFilter = '', sensitiveVisible = { phone: false, parentPhone: false, idNumber: false, address: false }, photoUrl = '', toastTimer, localMigrationAvailable = false, cloudSyncStatus = '未同步', cloudSyncAt = '';
const today = () => new Date().toISOString().slice(0, 10);
const esc = value => String(value ?? '').replace(/[&<>"']/g, char => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[char]));
const fmtDate = value => clean(value) ? esc(clean(value).slice(0, 10)) : '—';
const maskId = value => value ? `${esc(value.slice(0, 4))}**********${esc(value.slice(-4))}` : '—';
const student = () => state.students[selectedId];
const displayName = value => clean(value) || '—';
const studentName = id => state.students[id] ? displayName(state.students[id].name) : id;
const studentInitial = s => displayName(s?.name).slice(0, 1);
const courseLabel = grade => grade.courseName || grade.courseCode || '未命名课程';
const flatIcon = name => ({
  home: '<svg viewBox="0 0 24 24"><path d="m3 10 9-7 9 7v10H3z"/><path d="M9 21v-6h6v6"/></svg>',
  students: '<svg viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="2.5"/><path d="M5.5 17c.8-2.2 2-3.3 3.5-3.3s2.7 1.1 3.5 3.3M15 9h3M15 13h3M15 17h2"/></svg>',
  work: '<svg viewBox="0 0 24 24"><path d="M3 7h7l2 2h9v10a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><path d="M3 7V5a2 2 0 0 1 2-2h5l2 2h7a2 2 0 0 1 2 2v2"/></svg>',
  todo: '<svg viewBox="0 0 24 24"><rect x="4" y="3" width="16" height="18" rx="2"/><path d="m8 10 2 2 4-4M8 16h8"/></svg>',
  settings: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.9l.1.1-2.1 2.1-.1-.1a1.7 1.7 0 0 0-1.9-.3 1.7 1.7 0 0 0-1 1.5v.2h-3v-.2a1.7 1.7 0 0 0-1-1.5 1.7 1.7 0 0 0-1.9.3l-.1.1-2.1-2.1.1-.1A1.7 1.7 0 0 0 7 15a1.7 1.7 0 0 0-1.5-1H5.3v-3h.2A1.7 1.7 0 0 0 7 10a1.7 1.7 0 0 0-.3-1.9l-.1-.1 2.1-2.1.1.1a1.7 1.7 0 0 0 1.9.3 1.7 1.7 0 0 0 1-1.5v-.2h3v.2a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.9-.3l.1-.1 2.1 2.1-.1.1A1.7 1.7 0 0 0 19.4 10a1.7 1.7 0 0 0 1.5 1h.2v3h-.2a1.7 1.7 0 0 0-1.5 1Z"/></svg>',
  organization: '<svg viewBox="0 0 24 24"><circle cx="12" cy="5" r="2"/><circle cx="6" cy="19" r="2"/><circle cx="18" cy="19" r="2"/><path d="M12 7v5M6 17v-2h12v2M6 15v-3h12"/></svg>',
  grades: '<svg viewBox="0 0 24 24"><path d="M4 20V4M4 20h16"/><path d="m7 16 4-4 3 2 5-6"/></svg>',
  funding: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="8"/><path d="M15 9.5c-.5-.8-1.5-1.3-3-1.3-1.7 0-2.8.8-2.8 2 0 3.3 5.7 1.5 5.7 4.4 0 1.2-1.1 2-2.9 2-1.4 0-2.5-.5-3.1-1.3M12 6.5v11"/></svg>',
  mental: '<svg viewBox="0 0 24 24"><path d="M12 20s-7-4.4-7-10a3.7 3.7 0 0 1 6.6-2.3L12 8.3l.4-.6A3.7 3.7 0 0 1 19 10c0 5.6-7 10-7 10Z"/></svg>',
  contact: '<svg viewBox="0 0 24 24"><path d="M6 4h3l1.5 4-2 1.4c1.1 2.3 2.9 4.1 5.2 5.2l1.4-2L19 14v3c0 1.1-.9 2-2 2C10.4 19 5 13.6 5 7c0-1.1.9-2 1-3Z"/></svg>',
  dorm: '<svg viewBox="0 0 24 24"><path d="M4 21V4h16v17M8 8h2M14 8h2M8 12h2M14 12h2M10 21v-5h4v5"/></svg>'
}[name] || '');
const releases = [
  { version: '1.0.18', updatedAt: '2026-09-22 11:00', notes: ['新增邮箱账号登录、注册和邮件重置密码。', '学生资料按账号分别同步到云端；证件照和截图继续仅保存在本机。'] },
  { version: '1.0.17', updatedAt: '2026-09-22 10:15', notes: ['宿舍分布图与 Excel 改为横向房间号、纵向楼层，楼层从高到低排列。', '宿舍分布图左右滑动时，楼层列保持固定。'] },
  { version: '1.0.16', updatedAt: '2026-09-22 10:00', notes: ['学生详情中的敏感信息改为一键显示或隐藏。', '学生搜索会在输入时即时筛选并提示无匹配结果。'] },
  { version: '1.0.15', updatedAt: '2026-09-22 09:15', notes: ['宿舍分布支持按楼号、楼层、房间号生成分布图和 Excel。'] },
  { version: '1.0.14', updatedAt: '2026-09-22 09:00', notes: ['本人电话、家长电话和身份证号支持独立显示或隐藏。', '工作模块新增按宿舍聚合的宿舍分布。'] },
  { version: '1.0.13', updatedAt: '2026-09-21 17:22', notes: ['待办模块支持直接新增待办，并显示今日和逾期数量。'] },
  { version: '1.0.12', updatedAt: '2026-09-21 17:18', notes: ['学生模块新增班级切换，可只查看指定班级学生。'] },
  { version: '1.0.11', updatedAt: '2026-09-21 17:15', notes: ['学生列表按班级分组、组内学号排序，并新增班级职务。'] },
  { version: '1.0.10', updatedAt: '2026-09-21 17:12', notes: ['底部五个导航和工作模块图标统一为扁平线框设计。'] },
  { version: '1.0.9', updatedAt: '2026-09-21 17:05', notes: ['修复 iPhone 学生搜索框输入文字时被页面重绘中断的问题。'] },
  { version: '1.0.8', updatedAt: '2026-09-21 17:00', notes: ['界面调整为更接近 iPhone 原生的分组列表和联系人详情样式。', '首页、学生、工作、待办和设置统一视觉层级。'] },
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
async function persist() { cloudSyncStatus = '同步中…'; try { await saveCloudState(state); cloudSyncStatus = '已同步'; cloudSyncAt = now(); } catch (error) { cloudSyncStatus = '同步失败'; throw error; } render(); }
async function syncCloudNow() { cloudSyncStatus = '同步中…'; render(); try { await saveCloudState(state); cloudSyncStatus = '已同步'; cloudSyncAt = now(); notify('已同步到云端'); } catch (error) { cloudSyncStatus = '同步失败'; notify(error.message || '同步失败，请检查网络后重试', true); } render(); }
function setView(next) { view = next; if (next !== 'work') workModule = ''; render(); window.scrollTo(0, 0); }

async function gate(mode = 'signin') {
  const saved = await restoreAccount();
  if (saved) { localMigrationAvailable = await hasSecurity(); if (!localMigrationAvailable) await openLocalFiles(); state = await loadCloudState(emptyState); cloudSyncStatus = '已同步'; cloudSyncAt = now(); selectedTerm = terms(state)[0] || ''; render(); return; }
  const creating = mode === 'register';
  app.innerHTML = `<div class="lock-page"><div class="lock-card"><div class="lock-mark">档</div><h1>${creating ? '创建账号' : '欢迎回来'}</h1><p>${creating ? '使用邮箱创建个人工作台。学生资料只会同步到你的账号。' : '使用邮箱登录，查看属于你的学生资料。'}</p><form id="gate-form">${field('邮箱','email','','email','required autocomplete="email" inputmode="email"')}${field('密码','password','','password','required minlength="6" autocomplete="current-password"')}${creating ? field('确认密码','confirm','','password','required minlength="6" autocomplete="new-password"') : ''}<button class="btn" type="submit">${creating ? '创建并进入' : '登录'}</button></form><div class="gate-links">${creating ? '<button type="button" id="show-signin">已有账号，去登录</button>' : '<button type="button" id="show-register">创建新账号</button><button type="button" id="reset-password">忘记密码</button>'}</div><div class="lock-foot">附件仅保存在当前设备，不会上传云端。</div></div></div>`;
  app.querySelector('#show-signin')?.addEventListener('click', () => gate('signin'));
  app.querySelector('#show-register')?.addEventListener('click', () => gate('register'));
  app.querySelector('#reset-password')?.addEventListener('click', () => resetPasswordForm());
  app.querySelector('#gate-form').addEventListener('submit', async event => {
    event.preventDefault(); const form = new FormData(event.currentTarget), email = clean(form.get('email')), password = String(form.get('password'));
    try {
      if (creating) { if (password !== String(form.get('confirm'))) throw new Error('两次输入的密码不一致'); await register(email, password); }
      else await signIn(email, password);
      localMigrationAvailable = await hasSecurity(); if (!localMigrationAvailable) await openLocalFiles(); state = await loadCloudState(emptyState); cloudSyncStatus = '已同步'; cloudSyncAt = now(); selectedTerm = terms(state)[0] || ''; render();
    } catch (error) { notify(error.message || String(error), true); }
  });
}

function shell(body) {
  app.innerHTML = `<div class="shell"><header class="topbar"><div class="toprow"><div class="brand">辅导员工作台</div><button class="icon-btn" data-action="lock" aria-label="锁定">⌁</button></div></header><main class="content">${body}</main><nav class="tabbar five-tabs" aria-label="主导航">${[['home','首页'],['students','学生'],['work','工作'],['todo','待办'],['settings','设置']].map(([key, label]) => `<button data-view="${key === 'todo' ? 'alerts' : key === 'settings' ? 'data' : key}" class="${(key === 'todo' ? 'alerts' : key === 'settings' ? 'data' : key) === view ? 'active' : ''}"><span class="tab-icon">${flatIcon(key)}</span><span>${label}</span></button>`).join('')}</nav></div>`;
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
  const query = search.toLowerCase(), classes = [...new Set(Object.values(state.students).map(s => s.className || '未分班'))].sort((a,b) => a.localeCompare(b,'zh-CN')), filtered = Object.values(state.students).filter(s => (!classFilter || (s.className || '未分班') === classFilter) && [s.id,s.name,s.className,s.classRole,s.major,s.dorm].some(v => clean(v).toLowerCase().includes(query))).sort((a,b) => clean(a.className).localeCompare(clean(b.className),'zh-CN') || clean(a.id).localeCompare(clean(b.id),'zh-CN'));
  const groups = filtered.reduce((result, s) => { const key = s.className || '未分班'; (result[key] ||= []).push(s); return result; }, {});
  return `<div class="toolbar"><h2>学生 <span id="student-count" class="muted tiny">${filtered.length}</span></h2><div class="toolbar-actions"><button class="btn small ghost" data-action="class-filter">${esc(classFilter || '全部班级')}</button><button class="btn small" data-action="add-student">＋ 新增</button></div></div><div class="search-wrap"><input class="search" id="student-search" type="search" placeholder="搜索姓名、学号、班级、职务、专业" value="${esc(search)}"></div><div id="student-search-empty" class="search-empty" hidden>没有匹配的学生</div>${filtered.length ? Object.entries(groups).map(([group, rows]) => `<section class="student-group"><h3>${esc(group)} <small>${rows.length}</small></h3><div class="panel list">${rows.map(s => `<button class="list-row student-row" data-student="${esc(s.id)}" data-search="${esc([s.id,s.name,s.className,s.classRole,s.major,s.dorm].map(clean).join(' ').toLowerCase())}"><div class="avatar">${esc(studentInitial(s))}</div><div class="row-main"><strong>${esc(displayName(s.name))}</strong><small>${esc(s.className || '未分班')} · ${esc(s.id)}</small></div>${s.classRole ? `<span class="role-tag">${esc(s.classRole)}</span>` : ''}<span class="chevron">›</span></button>`).join('')}</div></section>`).join('') : `<div class="panel">${empty(search ? '没有匹配的学生' : '还没有学生档案')}</div>`}`;
}
function classFilterForm() { const classes = [...new Set(Object.values(state.students).map(s => s.className || '未分班'))].sort((a,b) => a.localeCompare(b,'zh-CN')); showModal('切换班级', choose('班级','className',['全部班级', ...classes], classFilter || '全部班级'), '确定', async form => { classFilter = form.get('className') === '全部班级' ? '' : clean(form.get('className')); search = ''; render(); }); }
const dialNumber = value => clean(value).replace(/[^\d+]/g, '');
const profileMissing = student => ['姓名','班级','专业','宿舍','本人电话','家长电话','家庭住址','身份证号','证件照'].filter(label => ({ 姓名: student.name, 班级: student.className, 专业: student.major, 宿舍: student.dorm, 本人电话: student.phone, 家长电话: student.parentPhone, 家庭住址: student.address, 身份证号: student.idNumber, 证件照: student.photoId })[label] ? false : true);
const maskPhone = value => value ? `${value.slice(0,3)}****${value.slice(-4)}` : '—';
const maskAddress = value => value ? `${value.slice(0, Math.min(6, value.length))}****` : '—';
function infoItem(label, value, sensitiveKey = '') {
  const visible = !sensitiveKey || sensitiveVisible[sensitiveKey], phone = ['phone', 'parentPhone'].includes(sensitiveKey) && dialNumber(value);
  const masked = sensitiveKey === 'idNumber' ? maskId(value) : sensitiveKey === 'address' ? maskAddress(value) : maskPhone(value);
  const content = visible && phone ? `<a class="phone-link" href="tel:${esc(phone)}">${esc(value)}</a>` : esc(visible ? value || '—' : masked);
  return `<div class="info-item"><small>${esc(label)}</small><strong>${content}</strong></div>`;
}
function detailView() {
  const s = student(); const grades = state.grades.filter(g => g.studentId === s.id), funding = state.funding.filter(f => f.studentId === s.id), records = state.records.filter(r => r.studentId === s.id).sort((a,b) => (b.date || '').localeCompare(a.date || ''));
  const tabs = [['info','信息'],['grades','成绩'],['funding','资助'],['records','记录']];
  let body = '';
  if (detailTab === 'info') {
    const fields = [['专业',s.major],['班级',s.className],['姓名',s.name],['学号',s.id],['民族',s.ethnicity],['性别',s.gender],['本人电话',s.phone,'phone'],['家长电话',s.parentPhone,'parentPhone'],['身份证号',s.idNumber,'idNumber'],['出生日期',s.birthDate],['宿舍',s.dorm],['家庭住址',s.address,'address']];
    const allSensitiveVisible = Object.values(sensitiveVisible).every(Boolean);
    body = `<div class="contact-actions">${sensitiveVisible.phone && s.phone ? `<a href="tel:${esc(dialNumber(s.phone))}">☎ 本人电话</a>` : ''}${sensitiveVisible.parentPhone && s.parentPhone ? `<a href="tel:${esc(dialNumber(s.parentPhone))}">☎ 家长电话</a>` : ''}</div><div class="panel pad"><div class="info-grid">${fields.map(([label,value,key]) => infoItem(label,value,key)).join('')}${Object.entries(s.custom || {}).map(([k,v]) => infoItem(k,v)).join('')}</div></div><div class="btn-row" style="margin-top:12px"><button class="btn secondary" data-action="edit-student">编辑档案</button><button class="btn ghost" data-action="toggle-sensitive">${allSensitiveVisible ? '隐藏全部信息' : '显示全部信息'}</button></div>`;
  } else if (detailTab === 'grades') {
    const stats = gradeStats(state,s.id,selectedTerm);
    body = `<div class="stats"><div class="stat warn"><strong>${stats.semester}</strong><span>本学期挂科</span></div><div class="stat"><strong>${stats.cumulative}</strong><span>历史累计</span></div><div class="stat"><strong>${stats.resolved}</strong><span>后来已通过</span></div></div><div class="section-title"><h2>成绩记录</h2><button class="btn small" data-action="add-grade">＋ 添加</button></div><div class="panel">${grades.length ? grades.sort((a,b) => b.term.localeCompare(a.term)).map(g => `<div class="record"><div class="record-top"><strong>${esc(courseLabel(g))}</strong><span class="badge ${isFailed(g) ? 'red' : 'green'}">${isFailed(g) ? '曾挂科' : '及格'}</span></div><p>${esc(g.term)} · 成绩 ${esc(g.score || '—')} ${g.retakeStatus ? `· ${esc(g.retakeStatus)}` : ''}</p><div class="actions"><button class="btn small ghost" data-action="edit-grade" data-id="${esc(g.id)}">编辑</button></div></div>`).join('') : empty('暂无成绩', '可手工添加或导入 Excel。')}</div>`;
  } else if (detailTab === 'funding') {
    const total = funding.reduce((sum,f) => sum + (Number(f.amount) || 0),0);
    body = `<div class="hero"><div class="kicker">累计发放金额</div><strong>¥ ${total.toLocaleString('zh-CN')}</strong><p>基于已录入的资助记录</p></div><div class="section-title"><h2>资助记录</h2><button class="btn small" data-action="add-funding">＋ 添加</button></div><div class="panel">${funding.length ? funding.map(f => `<div class="record"><div class="record-top"><strong>${esc(f.program || '资助项目')}</strong><span class="badge">¥ ${esc(f.amount || '0')}</span></div><p>${fmtDate(f.paidAt)} · 困难等级 ${esc(f.hardship || '未填')}</p>${f.family ? `<p>家庭情况：${esc(f.family)}</p>` : ''}<div class="actions"><button class="btn small ghost" data-action="edit-funding" data-id="${esc(f.id)}">编辑</button></div></div>`).join('') : empty('暂无资助记录', '可记录困难等级、家庭情况和历次资助。')}</div>`;
  } else {
    body = `<div class="toolbar"><h2>工作记录</h2><button class="btn small" data-action="add-record">＋ 记录</button></div><div class="panel">${records.length ? records.map(r => `<div class="record"><div class="record-top"><strong>${esc(r.type)} · ${fmtDate(r.date)}</strong><span class="badge ${r.done ? 'green' : r.todo ? 'red' : 'gray'}">${r.done ? '已完成' : r.todo ? '待跟进' : '已记录'}</span></div><p>对象：${esc(r.subject || '—')}</p><p>${esc(r.summary || '无摘要')}</p>${r.todo ? `<p>后续：${esc(r.todo)} · 到期 ${fmtDate(r.dueDate)}</p>` : ''}<div class="actions"><button class="btn small ghost" data-action="edit-record" data-id="${esc(r.id)}">编辑</button>${r.todo ? `<button class="btn small secondary" data-action="toggle-done" data-id="${esc(r.id)}">${r.done ? '恢复待办' : '标记完成'}</button>` : ''}${(r.attachments || []).map((id,i) => `<button class="btn small ghost" data-action="open-file" data-id="${esc(id)}">截图 ${i+1}</button>`).join('')}</div></div>`).join('') : empty('暂无工作记录', '谈话和家校联系可在这里记录并附截图。')}</div>`;
  }
  return `<button class="back" data-action="back-students">← 返回学生列表</button><div class="panel detail-head"><div class="avatar" id="detail-avatar">${esc(studentInitial(s))}</div><div><h2>${esc(displayName(s.name))}</h2><p>${esc(s.className || '未分班')} · ${esc(s.gender || '未填性别')}</p><span class="pill">学号 ${esc(s.id)}</span></div></div><div class="segmented">${tabs.map(([key,label]) => `<button data-detail-tab="${key}" class="${detailTab === key ? 'active' : ''}">${label}</button>`).join('')}</div>${body}`;
}
async function loadPhoto(id) {
  const s = state.students[id]; if (!s?.photoId) return;
  try { const file = await getFile(s.photoId); if (!file || selectedId !== id) return; photoUrl = URL.createObjectURL(file.blob); const avatar = document.querySelector('#detail-avatar'); if (avatar) avatar.innerHTML = `<img src="${photoUrl}" alt="${esc(s.name)}证件照">`; }
  catch (error) { notify(error.message,true); }
}
function alertsView() {
  const all = state.records.filter(r => r.todo), overdue = all.filter(r => !r.done && r.dueDate && r.dueDate < today()).length, dueToday = all.filter(r => !r.done && r.dueDate === today()).length, due = all.filter(r => todoFilter === 'done' ? r.done : todoFilter === 'overdue' ? !r.done && r.dueDate && r.dueDate < today() : todoFilter === 'today' ? !r.done && r.dueDate === today() : !r.done).sort((a,b) => (a.dueDate || '9999').localeCompare(b.dueDate || '9999'));
  return `<div class="toolbar"><h2>待办</h2><button class="btn small" data-action="add-todo">＋ 新增</button></div><div class="todo-summary"><span>今日 ${dueToday}</span><span class="${overdue ? 'danger' : ''}">逾期 ${overdue}</span></div><div class="segmented todo-filter">${[['all','待处理'],['today','今日'],['overdue','已逾期'],['done','已完成']].map(([key,label]) => `<button data-action="todo-filter" data-filter="${key}" class="${todoFilter === key ? 'active' : ''}">${label}</button>`).join('')}</div><div class="panel list">${due.length ? due.map(r => `<button class="list-row" data-student="${esc(r.studentId)}"><div class="avatar">✓</div><div class="row-main"><strong>${esc(studentName(r.studentId))} · ${esc(r.todo)}</strong><small>${r.dueDate ? `到期 ${fmtDate(r.dueDate)}` : '未设到期日'} · ${esc(r.type)}</small></div><span class="badge ${r.done ? 'green' : r.dueDate && r.dueDate < today() ? 'red' : ''}">${r.done ? '已完成' : r.priority || (r.dueDate && r.dueDate < today() ? '已逾期' : '待办')}</span></button>`).join('') : empty('暂无待办')}</div>`;
}

const workModules = [
  ['organization', '组织发展'], ['grades', '学业成绩'], ['funding', '资助工作'], ['mental', '心理工作'], ['contact', '家校联系'], ['dorm', '宿舍分布']
];
function workView() {
  if (workModule) return workDetailView();
  const counts = { organization: state.records.filter(r => r.type === '组织发展').length, grades: state.grades.length, funding: state.funding.length, mental: state.records.filter(r => r.type === '心理工作').length, contact: state.records.filter(r => r.type === '家校联系').length, dorm: new Set(Object.values(state.students).map(s => s.dorm).filter(Boolean)).size };
  return `<div class="toolbar"><h2>工作</h2></div><div class="panel work-list">${workModules.map(([key, label]) => `<button class="work-card" data-action="work-module" data-module="${key}"><span>${flatIcon(key)}</span><strong>${label}</strong><small>${counts[key]} 条</small><i>›</i></button>`).join('')}</div>`;
}
function workRecordList(type, emptyTitle) {
  const query = clean(workSearch).toLowerCase(), records = state.records.filter(record => record.type === type && [studentName(record.studentId), state.students[record.studentId]?.className, record.date, record.summary, record.subject].some(value => clean(value).toLowerCase().includes(query))).sort((a,b) => (b.date || '').localeCompare(a.date || ''));
  return `<div class="panel">${records.length ? records.map(record => `<button class="list-row" data-student="${esc(record.studentId)}"><div class="avatar">${type === '心理工作' ? '♡' : type === '组织发展' ? '◇' : '⌁'}</div><div class="row-main"><strong>${esc(studentName(record.studentId))}</strong><small>${fmtDate(record.date)} · ${esc(record.summary || '无摘要')}</small></div><span class="chevron">›</span></button>`).join('') : empty(emptyTitle)}</div>`;
}
function parseDorm(value) {
  const match = clean(value).replace('＃', '#').match(/^(\d+)#(\d)(\d{2})$/);
  return match ? { building: match[1], floor: match[2], room: match[3] } : undefined;
}
function dormDistribution() {
  const buildings = {};
  Object.values(state.students).forEach(student => {
    const dorm = parseDorm(student.dorm); if (!dorm) return;
    const building = buildings[dorm.building] ||= { floors: new Set(), rooms: {} };
    building.floors.add(dorm.floor);
    (building.rooms[dorm.room] ||= {})[dorm.floor] ||= [];
    building.rooms[dorm.room][dorm.floor].push(student);
  });
  const sheets = {}, previews = [];
  Object.entries(buildings).sort(([a],[b]) => Number(a) - Number(b)).forEach(([building, data]) => {
    const floors = [...data.floors].sort((a,b) => Number(b) - Number(a));
    const rooms = Object.keys(data.rooms).sort((a,b) => Number(a) - Number(b));
    const rows = floors.map(floor => [floor, ...rooms.map(room => (data.rooms[room][floor] || []).map(student => `${student.className || '未分班'} ${studentName(student.id)}`).join('\n'))]);
    sheets[`${building}号楼`] = [['楼层 / 房间号', ...rooms], ...rows.map(([floor, ...cells]) => [`${floor}层`, ...cells])];
    previews.push({ building, floors, rooms, rows });
  });
  return { sheets, previews };
}
function dormMapForm() {
  const distribution = dormDistribution();
  const body = distribution.previews.length ? `<div class="dorm-map">${distribution.previews.map(({building,rooms,rows}) => `<section><h3>${esc(building)}号楼</h3><div class="map-scroll"><table><thead><tr><th>楼层</th>${rooms.map(room => `<th>${esc(room)}室</th>`).join('')}</tr></thead><tbody>${rows.map(row => `<tr><th>${esc(row[0])}层</th>${row.slice(1).map(value => `<td>${esc(value).replaceAll('\n','<br>')}</td>`).join('')}</tr>`).join('')}</tbody></table></div></section>`).join('')}</div>` : '<div class="empty"><strong>暂无符合“楼号#楼层房间号”格式的宿舍信息</strong></div>';
  showModal('宿舍分布图', body, '下载 Excel', async () => { if (!Object.keys(distribution.sheets).length) throw new Error('暂无可导出的宿舍数据'); download(createWorkbook(distribution.sheets), `宿舍分布_${stamp()}.xlsx`); });
}
function workDetailView() {
  const module = workModules.find(([key]) => key === workModule);
  if (!module) { workModule = ''; return workView(); }
  const [, title] = module;
  let body = '';
  if (workModule === 'grades') {
    const query = clean(workSearch).toLowerCase(), threshold = Number(state.settings.threshold || 1), warnings = selectedTerm ? Object.values(state.students).map(s => ({ s, count: gradeStats(state,s.id,selectedTerm).semester })).filter(item => item.count >= threshold && [studentName(item.s.id), item.s.className, item.s.major].some(value => clean(value).toLowerCase().includes(query))).sort((a,b) => b.count - a.count) : [];
    body = `<div class="notice">本学期挂科达到 <strong>${threshold} 门</strong>时提醒。</div><div class="section-title"><h2>学业预警</h2><select id="term-select" class="mini-select"><option value="">选择学期</option>${terms(state).map(term => `<option value="${esc(term)}" ${term === selectedTerm ? 'selected' : ''}>${esc(term)}</option>`).join('')}</select></div><div class="panel list">${warnings.length ? warnings.map(({s,count}) => `<button class="list-row" data-student="${esc(s.id)}"><div class="avatar">${esc(studentInitial(s))}</div><div class="row-main"><strong>${esc(studentName(s.id))}</strong><small>${esc(s.className)} · 历史累计 ${gradeStats(state,s.id,selectedTerm).cumulative} 门</small></div><span class="badge red">${count} 门</span></button>`).join('') : empty('暂无学业预警')}</div>`;
  } else if (workModule === 'funding') {
    const query = clean(workSearch).toLowerCase(), rows = state.funding.filter(item => [studentName(item.studentId), state.students[item.studentId]?.className, item.paidAt, item.program].some(value => clean(value).toLowerCase().includes(query))).sort((a,b) => (b.paidAt || '').localeCompare(a.paidAt || ''));
    body = `<div class="panel list">${rows.length ? rows.map(item => `<button class="list-row" data-student="${esc(item.studentId)}"><div class="avatar">¥</div><div class="row-main"><strong>${esc(studentName(item.studentId))} · ${esc(item.program || '资助项目')}</strong><small>${fmtDate(item.paidAt)} · ¥ ${esc(item.amount || '0')}</small></div><span class="chevron">›</span></button>`).join('') : empty('暂无资助记录')}</div>`;
  } else if (workModule === 'dorm') {
    const groups = Object.values(state.students).filter(s => clean(s.dorm)).reduce((result, s) => { (result[s.dorm] ||= []).push(s); return result; }, {});
    body = `<div class="dorm-note">按宿舍聚合；每张卡显示同住学生，便于核对和联络。</div><div class="dorm-grid">${Object.entries(groups).sort(([a],[b]) => a.localeCompare(b,'zh-CN')).map(([dorm, students]) => `<section class="dorm-card"><div><strong>${esc(dorm)}</strong><small>${students.length} 人</small></div><p>${students.sort((a,b) => clean(a.id).localeCompare(clean(b.id),'zh-CN')).map(s => `<button data-student="${esc(s.id)}">${esc(studentName(s.id))}</button>`).join('')}</p></section>`).join('') || empty('暂无宿舍信息')}</div>`;
  } else if (workModule === 'organization') body = workRecordList('组织发展', '暂无组织发展记录');
  else if (workModule === 'mental') body = workRecordList('心理工作', '暂无心理工作记录');
  else body = workRecordList('家校联系', '暂无家校联系记录');
  const action = workModule === 'grades' ? 'work-add-grade' : workModule === 'funding' ? 'work-add-funding' : 'work-add-record';
  return `<button class="back" data-action="back-work">← 返回工作</button><div class="toolbar"><h2>${esc(title)}</h2>${workModule === 'dorm' ? `<button class="btn small" data-action="dorm-map">分布图</button>` : `<button class="btn small" data-action="${action}" data-type="${esc(title)}">＋ 新增</button>`}</div>${workModule === 'dorm' ? '' : `<div class="search-wrap"><input class="search" id="work-search" type="search" placeholder="搜索学生、班级、日期或内容" value="${esc(workSearch)}"></div>`}${body}`;
}
function dataView() {
  const current = releases[0], user = account();
  return `<div class="toolbar"><h2>设置</h2></div><div class="section-title"><h2>账号</h2></div><div class="panel"><div class="file-card"><h3>${esc(user?.email || '未登录')}</h3><p>学生资料仅同步到此账号。</p><div class="btn-row"><button class="btn ghost" data-action="reset-password">重置密码</button><button class="btn secondary" data-action="sign-out">退出登录</button></div></div></div><div class="section-title"><h2>Excel 模板</h2></div><div class="panel"><div class="file-card"><h3>空白固定模板</h3><button class="btn secondary" data-action="template">下载空白模板</button></div><div class="file-card"><h3>导入 Excel</h3><button class="btn secondary" data-action="import-xlsx">选择 Excel 文件</button></div><div class="file-card"><h3>导出当前数据</h3><button class="btn secondary" data-action="export-xlsx">导出 Excel</button></div></div><div class="section-title"><h2>离线附件与备份</h2></div><div class="panel"><div class="file-card"><h3>证件照批量导入</h3><p>证件照和工作截图仅保存于这台设备。</p><button class="btn ghost" data-action="import-photos">选择照片 ZIP</button></div><div class="file-card"><h3>导出加密数据包</h3><p>上次备份：${state.settings.lastExportedAt ? esc(state.settings.lastExportedAt.replace('T',' ').slice(0,16)) : '尚未备份'}</p><button class="btn" data-action="export-package">生成传输包 / 备份</button></div><div class="file-card"><h3>导入加密数据包</h3><button class="btn ghost" data-action="import-package">选择数据包 ZIP</button></div></div><div class="section-title"><h2>版本</h2></div><button class="version-card" data-action="version-history"><span><strong>v${esc(APP_VERSION)}</strong><small>更新于 ${esc(current.updatedAt)}</small></span><span class="chevron">›</span></button>`;
}

function versionHistory() {
  showModal('历史更新记录', `<div class="release-list">${releases.map(release => `<section class="release"><div><strong>v${esc(release.version)}</strong><small>${esc(release.updatedAt)}</small></div><ul>${release.notes.map(note => `<li>${esc(note)}</li>`).join('')}</ul></section>`).join('')}</div>`, '关闭', async () => {});
}
function render() {
  if (!state) return gate();
  const body = view === 'home' ? homeView() : view === 'students' ? studentsView() : view === 'work' ? workView() : view === 'alerts' ? alertsView() : dataView();
  shell(body);
  if (view === 'data') document.querySelector('.content').insertAdjacentHTML('afterbegin', `<div class="section-title"><h2>云端同步</h2></div><div class="panel"><div class="file-card"><h3>${esc(cloudSyncStatus)}</h3><p>最近同步：${cloudSyncAt ? esc(cloudSyncAt.replace('T',' ').slice(0,16)) : '尚未同步'}</p><button class="btn secondary" data-action="sync-cloud" ${cloudSyncStatus === '同步中…' ? 'disabled' : ''}>立即同步</button></div></div>`);
  if (view === 'data' && localMigrationAvailable) document.querySelector('.content').insertAdjacentHTML('afterbegin', '<div class="section-title"><h2>迁移原本机资料</h2></div><div class="panel"><div class="file-card"><h3>原 6 位密码保护的资料</h3><p>将这台设备已有学生资料同步到当前账号；照片和截图继续保存在本机。</p><button class="btn ghost" data-action="migrate-local">开始迁移</button></div></div>');
  const searchInput = document.querySelector('#student-search');
  if (searchInput) {
    const filterStudents = value => {
      search = value;
      const query = search.toLowerCase();
      const rows = [...document.querySelectorAll('.student-row')];
      rows.forEach(row => row.hidden = !row.dataset.search.includes(query));
      document.querySelectorAll('.student-group').forEach(group => group.hidden = ![...group.querySelectorAll('.student-row')].some(row => !row.hidden));
      const matches = rows.filter(row => !row.hidden);
      const count = document.querySelector('#student-count');
      const emptyHint = document.querySelector('#student-search-empty');
      if (count) count.textContent = matches.length;
      if (emptyHint) emptyHint.hidden = !query || matches.length > 0;
    };
    searchInput.addEventListener('input', event => filterStudents(event.target.value));
    searchInput.addEventListener('compositionend', event => filterStudents(event.target.value));
    searchInput.addEventListener('search', event => filterStudents(event.target.value));
  }
  const workInput = document.querySelector('#work-search');
  if (workInput) workInput.addEventListener('input', event => { workSearch = event.target.value; render(); document.querySelector('#work-search')?.focus(); });
  document.querySelector('#term-select')?.addEventListener('change', event => { selectedTerm = event.target.value; render(); });
}

function studentForm(existing) {
  const s = existing || {};
  showModal(existing ? '编辑学生档案' : '新增学生', `${field('学号 *','id',s.id,'text',existing ? 'readonly required' : 'required')}${field('姓名','name',s.name)}${field('班级','className',s.className)}${field('班级职务','classRole',s.classRole)}${field('专业','major',s.major)}${field('宿舍','dorm',s.dorm)}${field('本人电话','phone',s.phone,'tel')}${field('家长电话','parentPhone',s.parentPhone,'tel')}${area('家庭住址','address',s.address)}${field('身份证号','idNumber',s.idNumber)}${choose('性别','gender',['','男','女','其他'],s.gender)}${field('民族','ethnicity',s.ethnicity)}${field('出生日期','birthDate',s.birthDate,'date')}${area('自定义字段（每行：字段名=内容）','custom',Object.entries(s.custom || {}).map(([k,v]) => `${k}=${v}`).join('\n'))}<p class="inline-note">学号用于关联其他数据，必须填写；其余字段可留空。有效的 18 位身份证号可自动生成出生日期。</p>`, '保存档案', async form => {
    const id = clean(form.get('id')), name = clean(form.get('name')); if (!id) throw new Error('请填写学号');
    const custom = {}; for (const line of String(form.get('custom') || '').split(/\r?\n/)) { const at = line.indexOf('='); if (at > 0) custom[clean(line.slice(0,at))] = clean(line.slice(at+1)); }
    const idNumber = clean(form.get('idNumber'));
    state.students[id] = { ...s, id, name, className: clean(form.get('className')), classRole: clean(form.get('classRole')), major: clean(form.get('major')), dorm: clean(form.get('dorm')), phone: clean(form.get('phone')), parentPhone: clean(form.get('parentPhone')), address: clean(form.get('address')), idNumber, gender: clean(form.get('gender')), ethnicity: clean(form.get('ethnicity')), birthDate: clean(form.get('birthDate')) || dateFromId(idNumber), custom, updatedAt: now() };
    const photo = form.get('photo'); if (photo instanceof File && photo.size) { const photoId = uid(), compressed = await normalizedPhoto(photo), ext = compressed.type === 'image/jpeg' ? 'jpg' : photo.name.split('.').pop()?.toLowerCase() || 'jpg'; await putFile({ id: photoId, studentId: id, kind: 'photo', name: fileName(state.students[id], '证件照', today(), 1, ext), type: compressed.type || photo.type || 'image/jpeg', createdAt: now(), blob: compressed }); state.students[id].photoId = photoId; }
    selectedId = id; detailTab = 'info'; view = 'students'; await persist(); notify('档案已保存');
  });
  modalRoot.querySelector('.modal-actions')?.insertAdjacentHTML('beforebegin', field('上传证件照','photo','','file','accept="image/*"'));
}
function pickStudent(next) { showModal('选择学生', choose('学生','studentId',['', ...Object.values(state.students).sort((a,b) => studentName(a.id).localeCompare(studentName(b.id),'zh-CN')).map(s => `${s.id}｜${studentName(s.id)}`)], ''), '下一步', async form => { const id = clean(form.get('studentId')).split('｜')[0]; if (!state.students[id]) throw new Error('请选择学生'); selectedId = id; setTimeout(next, 0); }); }
function todoForm() { pickStudent(() => showModal('新增待办', `${choose('类型','type',['谈话','家校联系','组织发展','心理工作'],'谈话')}${area('待办内容 *','todo')}${choose('优先级','priority',['普通','重要','紧急'],'普通')}${field('到期日期','dueDate',today(),'date','required')}`, '保存待办', async form => { const todo = clean(form.get('todo')); if (!todo) throw new Error('请填写待办内容'); state.records.push({ id: uid(), studentId: selectedId, type: clean(form.get('type')), date: today(), summary: '', todo, priority: clean(form.get('priority')), dueDate: clean(form.get('dueDate')), done: false, attachments: [], updatedAt: now() }); await persist(); notify('待办已添加'); })); }
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
  showModal(existing ? '编辑工作记录' : '添加工作记录', `${choose('类型','type',['谈心谈话','家校联系','学业预警','心理工作','其他'],r.type || (workModule === 'mental' ? '心理工作' : workModule === 'contact' ? '家校联系' : '谈心谈话'))}${field('日期','date',r.date || today(),'date','required')}${field('对象','subject',r.subject,'text','placeholder="例如 学生本人 / 母亲"')}${field('关注等级','attention',r.attention)}${field('来源','source',r.source)}${area('摘要 *','summary',r.summary)}${area('后续待办','todo',r.todo)}${area('处理结论','outcome',r.outcome)}${field('到期日期','dueDate',r.dueDate,'date')}${field('追加截图（可多选）','images','','file','accept="image/*" multiple')}`, '保存记录', async (form,data) => {
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
function resetPasswordForm() {
  showModal('重置密码', `${field('账号邮箱','email',account()?.email || '','','email','required autocomplete="email"')}<p class="inline-note">系统将向该邮箱发送重置链接。打开链接并设置新密码后，再回到这里登录。</p>`, '发送重置邮件', async form => { await sendResetEmail(clean(form.get('email'))); notify('重置邮件已发送，请前往邮箱完成验证'); });
}
function migrateLocalDataForm() {
  showModal('迁移本机资料', `${field('原 6 位本机密码','password','','password','required inputmode="numeric" pattern="[0-9]{6}" maxlength="6"')}<p class="inline-note">会将原有学生资料同步到当前账号；照片和截图仍留在本机。</p>`, '开始迁移', async form => { await unlockSecurity(String(form.get('password'))); state = await loadState(); await persist(); localMigrationAvailable = false; notify('本机资料已同步到当前账号'); });
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
      case 'lock': signOut(); state = undefined; selectedId = ''; gate(); break;
      case 'back-students': selectedId = ''; render(); break;
      case 'back-work': workModule = ''; render(); break;
      case 'work-module': workModule = button.dataset.module; render(); window.scrollTo(0, 0); break;
      case 'dorm-map': dormMapForm(); break;
      case 'todo-filter': todoFilter = button.dataset.filter; render(); break;
      case 'add-todo': todoForm(); break;
      case 'work-add-grade': pickStudent(() => gradeForm()); break;
      case 'work-add-funding': pickStudent(() => fundingForm()); break;
      case 'work-add-record': pickStudent(() => recordForm()); break;
      case 'add-student': studentForm(); break;
      case 'class-filter': classFilterForm(); break;
      case 'edit-student': studentForm(student()); break;
      case 'add-grade': gradeForm(); break;
      case 'edit-grade': gradeForm(state.grades.find(g => g.id === button.dataset.id)); break;
      case 'add-funding': fundingForm(); break;
      case 'edit-funding': fundingForm(state.funding.find(f => f.id === button.dataset.id)); break;
      case 'add-record': recordForm(); break;
      case 'edit-record': recordForm(state.records.find(r => r.id === button.dataset.id)); break;
      case 'photo': photoForm(); break;
      case 'threshold': thresholdForm(); break;
      case 'reset-password': resetPasswordForm(); break;
      case 'migrate-local': migrateLocalDataForm(); break;
      case 'sign-out': signOut(); state = undefined; selectedId = ''; gate(); break;
      case 'sync-cloud': await syncCloudNow(); break;
      case 'version-history': versionHistory(); break;
      case 'toggle-sensitive': { const next = !Object.values(sensitiveVisible).every(Boolean); Object.keys(sensitiveVisible).forEach(key => sensitiveVisible[key] = next); render(); break; }
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
  else if (state && hiddenAt && Date.now() - hiddenAt > 5 * 60 * 1000) { signOut(); state = undefined; selectedId = ''; closeModal(); gate(); }
});
if ('serviceWorker' in navigator && location.protocol !== 'file:') navigator.serviceWorker.register('./sw.js').catch(() => {});
gate().catch(error => { app.innerHTML = `<div class="empty"><strong>应用无法启动</strong><p>${esc(error.message)}</p></div>`; });
