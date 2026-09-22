const envId = 'counselor-workbench-d5bo59891098';
const gateway = `https://${envId}.api.tcloudbasegateway.com`;
const sessionKey = 'counselor-workbench-cloudbase-session-v1';
let session;
function readable(error) { const code = error?.error || error?.code || error?.message || String(error); return ({ invalid_username_or_password: '账号或密码不正确。', user_not_found: '未找到该邮箱对应的账号。', rate_limit_exceeded: '操作过于频繁，请稍后再试。', weak_password: '密码强度不足，请使用 8 位以上的字母、数字和符号组合。', invalid_verification_code: '验证码不正确或已过期。' })[code] || error?.error_description || '操作未完成，请检查网络后重试。'; }
async function request(path, { method = 'POST', body, token, headers = {} } = {}) { const response = await fetch(`${gateway}${path}`, { method, headers: { ...(token ? { Authorization: `Bearer ${token}` } : {}), ...(body ? { 'Content-Type': 'application/json' } : {}), ...headers }, body: body ? JSON.stringify(body) : undefined }); const data = await response.json().catch(() => ({})); if (!response.ok) throw new Error(readable(data)); return data; }
function claims(token) { try { return JSON.parse(atob(token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/'))); } catch { return {}; } }
function remember(data, identifier = '') { const payload = claims(data.access_token); session = { uid: data.sub || payload.sub, email: payload.email || (identifier.includes('@') ? identifier : ''), username: payload.username || identifier, accessToken: data.access_token, expiresAt: Date.now() + Number(data.expires_in || 7200) * 1000 - 60000 }; if (!session.uid) throw new Error('登录状态异常，请重新登录。'); localStorage.setItem(sessionKey, JSON.stringify(session)); return account(); }
export function account() { return session ? { uid: session.uid, email: session.email || session.username, username: session.username || session.email } : undefined; }
export async function restoreAccount() { try { session = JSON.parse(localStorage.getItem(sessionKey) || 'null'); if (!session || !session.uid || Date.now() >= session.expiresAt) { signOut(); return undefined; } return account(); } catch { signOut(); return undefined; } }
export async function signIn(username, password) { return remember(await request('/auth/v1/signin', { body: { username, password } }), username); }
export async function sendRegistrationCode(email) { const result = await request('/auth/v1/verification', { body: { email, target: 'ANY' } }); if (result.is_user) throw new Error('该邮箱已注册，请直接登录。'); return result.verification_id; }
export async function register({ displayName, email, password, verificationId, verificationCode }) { const verified = await request('/auth/v1/verification/verify', { body: { verification_id: verificationId, verification_code: verificationCode } }); return remember(await request('/auth/v1/signup', { body: { name: displayName, email, password, verification_token: verified.verification_token } }), email); }
export async function sendResetCode(email) { return (await request('/auth/v1/verification', { body: { email, target: 'USER' } })).verification_id; }
export async function resetPassword({ email, password, verificationId, verificationCode }) { const verified = await request('/auth/v1/verification/verify', { body: { verification_id: verificationId, verification_code: verificationCode } }); await request('/auth/v1/user/password/reset', { body: { email, new_password: password, verification_token: verified.verification_token } }); }
export function signOut() { session = undefined; localStorage.removeItem(sessionKey); }
function requireSession() { if (!session || Date.now() >= session.expiresAt) { signOut(); throw new Error('登录已过期，请重新登录。'); } return session; }
export async function loadCloudState(emptyState) { const user = requireSession(); const response = await fetch(`${gateway}/v1/rdb/rest/account_states?owner_id=eq.${encodeURIComponent(user.uid)}&select=payload`, { headers: { Authorization: `Bearer ${user.accessToken}` } }); const data = await response.json().catch(() => []); if (!response.ok) throw new Error(readable(data)); return data[0]?.payload || emptyState(); }
export async function saveCloudState(state) { const user = requireSession(); await request('/v1/rdb/rest/account_states', { method: 'POST', token: user.accessToken, headers: { Prefer: 'resolution=merge-duplicates,return=minimal' }, body: { owner_id: user.uid, payload: state, updated_at: new Date().toISOString() } }); }
async function leaveRequest(path, { method = 'GET', body, headers = {} } = {}) {
  const user = requireSession();
  const response = await fetch(`${gateway}/v1/rdb/rest${path}`, { method, headers: { Authorization: `Bearer ${user.accessToken}`, ...(body ? { 'Content-Type': 'application/json' } : {}), ...headers }, body: body ? JSON.stringify(body) : undefined });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(readable(data));
  return data;
}
export async function configureLeaveForm({ formToken, managerToken, roster }) {
  const user = requireSession();
  await leaveRequest('/leave_forms', { method: 'POST', headers: { Prefer: 'resolution=merge-duplicates,return=minimal' }, body: { owner_id: user.uid, form_token: formToken, manager_token: managerToken, roster, updated_at: new Date().toISOString() } });
}
export async function loadLeaveRequests() { const user = requireSession(); return leaveRequest(`/leave_requests?owner_id=eq.${encodeURIComponent(user.uid)}&order=created_at.desc`); }
export async function reviewLeaveRequest(id, status) {
  const patch = { status };
  if (status === '已返校') patch.returned_at = new Date().toISOString(); else patch.reviewed_at = new Date().toISOString();
  await leaveRequest(`/leave_requests?id=eq.${encodeURIComponent(id)}`, { method: 'PATCH', headers: { Prefer: 'return=minimal' }, body: patch });
}
