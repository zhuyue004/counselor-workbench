import { zipSync, unzipSync, strToU8, strFromU8 } from './vendor/fflate.js';
import { APP_VERSION } from './data.js';

const encoder = new TextEncoder();
const decoder = new TextDecoder();
const deriveKey = async (password, salt) => {
  const material = await crypto.subtle.importKey('raw', encoder.encode(password), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations: 250000, hash: 'SHA-256' }, material, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
};
const hex = bytes => [...bytes].map(n => n.toString(16).padStart(2, '0')).join('');
const unhex = value => new Uint8Array(value.match(/.{2}/g).map(part => parseInt(part, 16)));
export const stamp = (date = new Date()) => `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, '0')}${String(date.getDate()).padStart(2, '0')}_${String(date.getHours()).padStart(2, '0')}${String(date.getMinutes()).padStart(2, '0')}${String(date.getSeconds()).padStart(2, '0')}`;
export const packageName = (date = new Date()) => `辅导员工作台_数据包_v${APP_VERSION}_${stamp(date)}.zip`;

export async function makePackage(state, files, password) {
  if (password.length < 8) throw new Error('数据包密码至少需要 8 位');
  const inner = { 'state.json': strToU8(JSON.stringify(state)) };
  const metadata = [];
  for (const file of files) {
    inner[`files/${file.id}`] = new Uint8Array(await file.blob.arrayBuffer());
    metadata.push({ id: file.id, studentId: file.studentId, kind: file.kind, name: file.name, type: file.type, createdAt: file.createdAt });
  }
  inner['files.json'] = strToU8(JSON.stringify(metadata));
  const compressed = zipSync(inner, { level: 6 });
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const encrypted = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, await deriveKey(password, salt), compressed));
  const manifest = { format: 'counselor-workbench-encrypted', formatVersion: 1, appVersion: APP_VERSION, exportedAt: new Date().toISOString(), salt: hex(salt), iv: hex(iv) };
  return new Blob([zipSync({ 'manifest.json': strToU8(JSON.stringify(manifest)), 'payload.enc': encrypted }, { level: 0 })], { type: 'application/zip' });
}

export async function readPackage(blob, password) {
  const outer = unzipSync(new Uint8Array(await blob.arrayBuffer()));
  if (!outer['manifest.json'] || !outer['payload.enc']) throw new Error('不是本应用的数据包');
  const manifest = JSON.parse(strFromU8(outer['manifest.json']));
  if (manifest.format !== 'counselor-workbench-encrypted' || manifest.formatVersion !== 1) throw new Error('数据包版本不受支持');
  let bytes;
  try { bytes = new Uint8Array(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: unhex(manifest.iv) }, await deriveKey(password, unhex(manifest.salt)), outer['payload.enc'])); }
  catch { throw new Error('密码不正确，或数据包已损坏'); }
  const inner = unzipSync(bytes);
  const state = JSON.parse(decoder.decode(inner['state.json']));
  if (state.schema !== 1) throw new Error('数据结构版本不受支持');
  const files = JSON.parse(decoder.decode(inner['files.json'])).map(meta => {
    const data = inner[`files/${meta.id}`];
    if (!data) throw new Error(`数据包缺少附件 ${meta.id}`);
    return { ...meta, blob: new Blob([data], { type: meta.type || 'application/octet-stream' }) };
  });
  return { state, files, manifest };
}

export function mergePackage(current, incoming) {
  const merged = structuredClone(current);
  let added = 0, updated = 0;
  for (const [id, student] of Object.entries(incoming.students || {})) {
    const old = merged.students[id];
    if (!old) { merged.students[id] = student; added++; }
    else if (student.updatedAt > old.updatedAt) { merged.students[id] = student; updated++; }
  }
  for (const key of ['grades', 'funding', 'records']) {
    const index = new Map(merged[key].map((item, i) => [item.id, i]));
    for (const item of incoming[key] || []) {
      const at = index.get(item.id);
      if (at === undefined) { merged[key].push(item); added++; }
      else if (item.updatedAt > merged[key][at].updatedAt) { merged[key][at] = item; updated++; }
    }
  }
  return { merged, added, updated };
}
