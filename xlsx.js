import { zipSync, unzipSync, strToU8, strFromU8 } from './vendor/fflate.js';

export const STUDENT_TEMPLATE_COLUMNS = ['序号', '专业', '班级', '学号', '姓名', '班级职务', '性别', '民族', '身份证号', '宿舍', '本人电话', '家长电话', '家庭住址'];
export const STUDENT_EXPORT_COLUMNS = [...STUDENT_TEMPLATE_COLUMNS, '出生日期'];
export const STUDENT_KNOWN_COLUMNS = new Set(STUDENT_EXPORT_COLUMNS);

export const COLUMNS = {
  学生: STUDENT_TEMPLATE_COLUMNS,
  成绩: ['记录ID', '学号', '学期', '课程代码', '课程名称', '成绩', '是否挂科', '补考或重修状态', '备注'],
  资助: ['记录ID', '学号', '困难等级', '家庭情况', '资助项目', '发放日期', '发放金额', '备注'],
  工作记录: ['记录ID', '学号', '类型', '日期', '对象', '关注等级', '来源', '摘要', '后续待办', '处理结论', '到期日期', '完成状态']
};

const xml = value => String(value ?? '').replace(/[&<>"']/g, char => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&apos;'}[char]));
const parseXml = value => {
  const doc = new DOMParser().parseFromString(value, 'application/xml');
  if (doc.querySelector('parsererror')) throw new Error('Excel 文件中的 XML 无法读取');
  return doc;
};
const colRef = index => {
  let result = '';
  for (let n = index + 1; n; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + ((n - 1) % 26)) + result;
  return result;
};
const colIndex = ref => [...ref.match(/^[A-Z]+/)?.[0] ?? ''].reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0) - 1;

export function createWorkbook(sheets) {
  const entries = Object.entries(sheets);
  const files = {
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${entries.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`),
    '_rels/.rels': strToU8('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${entries.map(([name], i) => `<sheet name="${xml(name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${entries.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`),
    'xl/styles.xml': strToU8('<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
  };
  entries.forEach(([name, rows], i) => {
    const body = rows.map((row, ri) => `<row r="${ri + 1}">${row.map((value, ci) => `<c r="${colRef(ci)}${ri + 1}" t="inlineStr"><is><t xml:space="preserve">${xml(value)}</t></is></c>`).join('')}</row>`).join('');
    const textCols = name === '学生' ? [4, 9] : [1, 2];
    const cols = `<cols>${textCols.map(n => `<col min="${n}" max="${n}" width="18" customWidth="1" style="1"/>`).join('')}</cols>`;
    files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(`<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${cols}<sheetData>${body}</sheetData></worksheet>`);
  });
  return new Blob([zipSync(files, { level: 6 })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

export function blankTemplate() {
  return createWorkbook(Object.fromEntries(Object.entries(COLUMNS).map(([name, headers]) => [name, [headers]])));
}

function normalizedPath(target) {
  const path = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
  const stack = [];
  path.split('/').forEach(part => { if (part === '..') stack.pop(); else if (part && part !== '.') stack.push(part); });
  return stack.join('/');
}

export async function readWorkbook(file) {
  const zip = unzipSync(new Uint8Array(await file.arrayBuffer()));
  const read = path => {
    if (!zip[path]) throw new Error(`Excel 缺少 ${path}`);
    return strFromU8(zip[path]);
  };
  const workbook = parseXml(read('xl/workbook.xml'));
  const rels = parseXml(read('xl/_rels/workbook.xml.rels'));
  const relation = Object.fromEntries([...rels.getElementsByTagName('Relationship')].map(el => [el.getAttribute('Id'), el.getAttribute('Target')]));
  const shared = zip['xl/sharedStrings.xml'] ? [...parseXml(read('xl/sharedStrings.xml')).getElementsByTagName('si')].map(el => [...el.getElementsByTagName('t')].map(t => t.textContent).join('')) : [];
  const result = {};
  for (const sheet of workbook.getElementsByTagName('sheet')) {
    const name = sheet.getAttribute('name');
    const id = sheet.getAttribute('r:id');
    const target = relation[id];
    if (!target) continue;
    const doc = parseXml(read(normalizedPath(target)));
    const rows = [];
    for (const row of doc.getElementsByTagName('row')) {
      const cells = [];
      for (const cell of row.getElementsByTagName('c')) {
        const index = colIndex(cell.getAttribute('r') || 'A1');
        if (index < 0 || index > 500) continue;
        const kind = cell.getAttribute('t');
        const value = kind === 'inlineStr' ? [...cell.getElementsByTagName('t')].map(t => t.textContent).join('') : cell.getElementsByTagName('v')[0]?.textContent ?? '';
        cells[index] = kind === 's' ? shared[Number(value)] ?? '' : value;
      }
      rows.push(cells.map(v => v ?? ''));
    }
    result[name] = rows;
  }
  return result;
}
