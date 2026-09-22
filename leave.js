const API = 'https://counselor-workbench-d5bo59891098-1493520377.ap-shanghai.app.tcloudbase.com/leave-api';
const formToken = new URLSearchParams(location.search).get('form') || '';
const form = document.querySelector('#leave-form');
const idInput = document.querySelector('#student-id');
const nameInput = document.querySelector('#student-name');
const typeInput = document.querySelector('#leave-type');
const offCampus = document.querySelector('.leave-offcampus');
const course = document.querySelector('.leave-course');
const returnInput = document.querySelector('#return-at');
const message = document.querySelector('#form-message');
const submit = document.querySelector('#submit-leave');
let lookupTimer;

function show(text, isError = false) { message.textContent = text; message.className = `leave-message ${text ? (isError ? 'error' : 'success') : ''}`; }
async function call(path, options = {}) {
  const response = await fetch(`${API}${path}`, { ...options, headers: { 'Content-Type': 'application/json', ...(options.headers || {}) } });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || '服务暂不可用，请稍后重试。');
  return data;
}
async function lookupStudent() {
  const studentId = idInput.value.trim();
  nameInput.value = '';
  if (!studentId) return;
  if (!formToken) { show('请使用辅导员提供的请假二维码进入。', true); return; }
  try {
    const result = await call(`/student?formToken=${encodeURIComponent(formToken)}&studentId=${encodeURIComponent(studentId)}`);
    nameInput.value = result.name || '';
    show(result.name ? `已核对：${result.name}` : '', false);
  } catch (error) { show(error.message, true); }
}
function syncType() {
  const outside = typeInput.value === '离校请假';
  offCampus.hidden = !outside;
  course.hidden = outside;
  returnInput.required = outside;
}
function currentTime() {
  const d = new Date(); d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); return d.toISOString().slice(0,16);
}
idInput.addEventListener('input', () => { clearTimeout(lookupTimer); lookupTimer = setTimeout(lookupStudent, 350); });
idInput.addEventListener('blur', lookupStudent);
typeInput.addEventListener('change', syncType);
form.addEventListener('submit', async event => {
  event.preventDefault();
  if (!formToken) return show('请使用辅导员提供的请假二维码进入。', true);
  if (!nameInput.value) return show('请先填写正确的学号并等待姓名核对。', true);
  const data = new FormData(form);
  submit.disabled = true; show('正在提交…');
  try {
    const result = await call('/submit', { method: 'POST', body: JSON.stringify({ formToken, studentId: data.get('studentId'), leaveType: data.get('leaveType'), startAt: data.get('startAt'), returnAt: data.get('returnAt'), courseInfo: data.get('courseInfo'), destination: data.get('destination'), reason: data.get('reason') }) });
    form.reset(); nameInput.value = ''; document.querySelector('#start-at').value = currentTime(); syncType(); show(result.message || '请假申请已提交。');
  } catch (error) { show(error.message, true); }
  finally { submit.disabled = false; }
});
document.querySelector('#start-at').value = currentTime();
syncType();
if (!formToken) { submit.disabled = true; show('二维码无效，请向辅导员重新获取。', true); }
