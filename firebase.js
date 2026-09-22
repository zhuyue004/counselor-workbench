// Firebase Web configuration. The API key identifies this public web client; it is not an administrator secret.
const config = {
  apiKey: 'AIzaSyDcd7DMNv9UslJp04o2UBEg1WHRAL3Xe4w',
  projectId: 'counselor-workbench'
};

const sessionKey = 'counselor-workbench-cloud-session-v1';
let session;
const authUrl = path => `https://identitytoolkit.googleapis.com/v1/${path}?key=${config.apiKey}`;
const firestoreUrl = uid => `https://firestore.googleapis.com/v1/projects/${config.projectId}/databases/(default)/documents/users/${encodeURIComponent(uid)}/app/state`;

function message(error) {
  const code = error?.error?.message || error?.message || String(error);
  return ({ EMAIL_EXISTS: '该邮箱已注册，请直接登录。', EMAIL_NOT_FOUND: '邮箱或密码不正确。', INVALID_PASSWORD: '邮箱或密码不正确。', INVALID_LOGIN_CREDENTIALS: '邮箱或密码不正确。', INVALID_EMAIL: '请输入有效的邮箱地址。', WEAK_PASSWORD: '密码至少需要 6 位。', TOO_MANY_ATTEMPTS_TRY_LATER: '尝试次数过多，请稍后再试。' })[code] || '操作未完成，请检查网络后重试。';
}
async function request(url, options = {}) {
  const response = await fetch(url, options);
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw data;
  return data;
}
function remember(next) {
  session = next;
  localStorage.setItem(sessionKey, JSON.stringify(next));
  return session;
}
function fromAuth(data) {
  return remember({ uid: data.localId, email: data.email, idToken: data.idToken, refreshToken: data.refreshToken, expiresAt: Date.now() + Number(data.expiresIn || 3600) * 1000 - 60000 });
}
async function refresh() {
  const body = new URLSearchParams({ grant_type: 'refresh_token', refresh_token: session.refreshToken });
  const data = await request(`https://securetoken.googleapis.com/v1/token?key=${config.apiKey}`, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body });
  return remember({ uid: data.user_id, email: session.email, idToken: data.id_token, refreshToken: data.refresh_token, expiresAt: Date.now() + Number(data.expires_in || 3600) * 1000 - 60000 });
}
async function token() {
  if (!session) throw new Error('请先登录');
  if (Date.now() >= session.expiresAt) await refresh();
  return session.idToken;
}
export function account() { return session ? { uid: session.uid, email: session.email } : undefined; }
export async function restoreAccount() {
  try {
    const saved = localStorage.getItem(sessionKey);
    if (!saved) return undefined;
    session = JSON.parse(saved);
    return account();
  } catch { signOut(); return undefined; }
}
export async function signIn(email, password) {
  try { return fromAuth(await request(authUrl('accounts:signInWithPassword'), { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ email, password, returnSecureToken: true }) })); }
  catch (error) { throw new Error(message(error)); }
}
export async function register(email, password) {
  try { return fromAuth(await request(authUrl('accounts:signUp'), { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ email, password, returnSecureToken: true }) })); }
  catch (error) { throw new Error(message(error)); }
}
export async function sendResetEmail(email) {
  try { await request(authUrl('accounts:sendOobCode'), { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ requestType: 'PASSWORD_RESET', email }) }); }
  catch (error) { throw new Error(message(error)); }
}
export function signOut() { session = undefined; localStorage.removeItem(sessionKey); }
export async function loadCloudState(emptyState) {
  const response = await fetch(firestoreUrl((await accountOrThrow()).uid), { headers: { Authorization: `Bearer ${await token()}` } });
  if (response.status === 404) return emptyState();
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error('无法读取云端资料，请检查网络后重试。');
  try { return JSON.parse(data.fields?.payload?.stringValue || ''); }
  catch { throw new Error('云端资料格式异常，未加载任何资料。'); }
}
async function accountOrThrow() { if (!session) throw new Error('请先登录'); return session; }
export async function saveCloudState(state) {
  const user = await accountOrThrow();
  await request(firestoreUrl(user.uid), { method: 'PATCH', headers: { Authorization: `Bearer ${await token()}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields: { payload: { stringValue: JSON.stringify(state) }, updatedAt: { timestampValue: new Date().toISOString() } } }) });
}
