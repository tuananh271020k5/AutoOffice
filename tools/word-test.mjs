import fs from 'node:fs';
import path from 'node:path';
import http from 'node:http';
import https from 'node:https';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const outputDir = path.join(root, '.autooffice-test');
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));

async function main() {
  const config = JSON.parse(fs.readFileSync(path.join(outputDir, 'connection.json'), 'utf8'));
  const base = new URL(config.baseUrl);
  if (base.hostname !== 'localhost' || !['http:', 'https:'].includes(base.protocol) ||
      !/^[a-f0-9]{64}$/.test(config.token)) throw new Error('Invalid local bridge configuration.');
  const request = (route, body) => new Promise((resolve, reject) => {
    const transport = base.protocol === 'https:' ? https : http;
    const data = body === undefined ? undefined : JSON.stringify(body);
    const req = transport.request(new URL(`/__autooffice_test${route}`, base), {
      method: data === undefined ? 'GET' : 'POST',
      // Only this localhost dev certificate is accepted without CA verification.
      ...(base.protocol === 'https:' ? { rejectUnauthorized: false } : {}),
      headers: { Authorization: `Bearer ${config.token}`,
        ...(data === undefined ? {} : { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(data) }) },
      timeout: 10_000,
    }, res => {
      let text = '';
      res.setEncoding('utf8');
      res.on('data', chunk => { text += chunk; });
      res.on('end', () => {
        try {
          const value = JSON.parse(text);
          if (res.statusCode >= 400) reject(new Error(`HTTP ${res.statusCode}: ${value.error}`));
          else resolve(value);
        } catch (error) { reject(error); }
      });
      res.on('error', reject);
    });
    req.on('timeout', () => req.destroy(new Error('Local bridge request timed out')));
    req.on('error', reject);
    req.end(data);
  });
  const execute = async (sessionId, action, args = {}) => {
    // POST once: never repeat a mutation on timeout or transport failure.
    const command = await request('/commands', { sessionId, action, args });
    console.error(`Command ${command.id}: ${action}`);
    const start = Date.now();
    while (Date.now() - start < 240_000) {
      await sleep(800);
      const current = await request(`/commands/${command.id}`);
      if (current.state === 'done' || current.state === 'error') {
        fs.mkdirSync(outputDir, { recursive: true });
        fs.writeFileSync(path.join(outputDir, `${command.id}.json`), JSON.stringify(current, null, 2));
        if (current.state === 'error') throw new Error(JSON.stringify(current.result));
        return current.result;
      }
    }
    throw new Error(`Command ${command.id} still pending. It may still execute; inspect its status before doing anything else. If approval is enabled, check the Word task pane.`);
  };
  const [action = 'sessions', sessionId, ...rest] = process.argv.slice(2);
  if (action === 'diagnostics') { console.log(JSON.stringify(await request('/diagnostics'), null, 2)); return; }
  if (action === 'sessions') { console.log(JSON.stringify(await request('/sessions'), null, 2)); return; }
  if (action === 'result' && sessionId) {
    console.log(JSON.stringify(await request(`/commands/${encodeURIComponent(sessionId)}`), null, 2)); return;
  }
  if (!sessionId) throw new Error('Usage: npm run test:word -- sessions | status|snapshot|suite <sessionId> | run|select <sessionId> <text> | fixture <sessionId> <JSON array> | result <commandId>');
  if (action === 'preview-test') {
    const decision = rest[0];
    if (!['accept', 'reject'].includes(decision)) throw new Error('preview-test requires accept or reject.');
    const status = await execute(sessionId, 'status');
    if (!status.previewWordChanges) throw new Error('Enable Word preview/recovery before testing this feature.');
    await execute(sessionId, 'fixture', { paragraphs: ['AUTO-OFFICE-PREVIEW-BEFORE', 'AUTO-OFFICE-KEEP'] });
    await execute(sessionId, 'select', { text: 'AUTO-OFFICE-PREVIEW-BEFORE' });
    console.error(`Waiting for ${decision} in the Word preview panel. The CLI does not approve on your behalf.`);
    const result = await execute(sessionId, 'run', {
      prompt: 'Chỉ thay phần đang bôi chọn bằng đúng AUTO-OFFICE-PREVIEW-AFTER. Giữ nguyên phần còn lại. Nếu tôi từ chối thì dừng chỉnh sửa.',
    });
    const expected = decision === 'accept' ? 'AUTO-OFFICE-PREVIEW-AFTER\nAUTO-OFFICE-KEEP'
      : 'AUTO-OFFICE-PREVIEW-BEFORE\nAUTO-OFFICE-KEEP';
    const passed = result.after.text.replace(/\r\n?/g, '\n').trim() === expected &&
      (decision !== 'reject' || result.turn.toolResults.some(tool => tool.status === 'rejected'));
    const reportFile = path.join(outputDir, `preview-${decision}.json`);
    fs.writeFileSync(reportFile, JSON.stringify({ decision, passed, result }, null, 2));
    console.log(JSON.stringify({ decision, passed, reportFile }, null, 2));
    if (!passed) process.exitCode = 1;
    return;
  }
  if (action === 'suite') {
    const status = await execute(sessionId, 'status');
    if (!status.hasApiKey) throw new Error('AI provider is not configured in the add-in. Enter your key in AutoOffice settings, never in this CLI.');
    if (status.previewWordChanges) throw new Error('Word preview/recovery is enabled. Use individual run commands and approve/reject in Word, or explicitly disable preview in Settings before running the unattended suite.');
    if (!status.autoApprove) throw new Error('The unattended suite requires AutoOffice auto-approval enabled. Individual run commands still support manual approval in Word.');
    const cases = [];
    const clean = text => String(text).replace(/\r\n?/g, '\n').trim();
    const checkTurn = result => !result.turn?.errors?.length &&
      !result.turn?.toolResults?.some(t => ['error', 'rejected'].includes(t.status));
    const test = async (name, paragraphs, prompt, assert, select) => {
      await execute(sessionId, 'fixture', { paragraphs });
      if (select) await execute(sessionId, 'select', { text: select });
      const result = await execute(sessionId, 'run', { prompt });
      const assertions = assert(result);
      cases.push({ name, passed: checkTurn(result) && Object.values(assertions).every(Boolean), assertions, result });
      console.error(`${cases.at(-1).passed ? 'PASS' : 'FAIL'}: ${name}`);
    };
    try {
      await test('Replace selected text and preserve surrounding text',
        ['GIỮ NGUYÊN ĐẦU.', 'Nội dung cần sửa.', 'GIỮ NGUYÊN CUỐI.'],
        'Chỉ thay đúng phần tôi đang bôi chọn bằng chính xác: Nội dung đã sửa. Giữ nguyên tất cả nội dung và định dạng khác.',
        r => ({ exactText: clean(r.after.text) === 'GIỮ NGUYÊN ĐẦU.\nNội dung đã sửa.\nGIỮ NGUYÊN CUỐI.',
          othersUnchanged: JSON.stringify(r.before.paragraphs.filter(p => p.text !== 'Nội dung cần sửa.')) ===
            JSON.stringify(r.after.paragraphs.filter(p => p.text !== 'Nội dung đã sửa.')) }),
        'Nội dung cần sửa.');
      await test('Summarize without modifying the document',
        ['Cuộc họp bắt đầu lúc 9 giờ.', 'Nhóm thống nhất hoàn thành báo cáo vào thứ Sáu.'],
        'Tóm tắt tài liệu này trong khung chat bằng một câu. Không chỉnh sửa tài liệu.',
        r => ({ unchanged: JSON.stringify(r.before) === JSON.stringify(r.after),
          answered: r.turn.modelMessages.some(m => m.role === 'assistant' &&
            (typeof m.content === 'string' ? m.content.trim() : m.content.some(p => p.type === 'text' && p.text.trim()))) }));
      await test('Apply bold only to the selected paragraph',
        ['ĐOẠN ĐẦU GIỮ NGUYÊN.', 'Định dạng đoạn này.', 'ĐOẠN CUỐI GIỮ NGUYÊN.'],
        'Chỉ in đậm phần tôi đang bôi chọn. Không thay đổi chữ và không định dạng các đoạn khác.',
        r => ({ textUnchanged: r.before.text === r.after.text,
          selectedBold: r.after.paragraphs.filter(p => p.text === 'Định dạng đoạn này.').length === 1 &&
            r.after.paragraphs.find(p => p.text === 'Định dạng đoạn này.').font.bold === true,
          othersUnchanged: JSON.stringify(r.before.paragraphs.filter(p => p.text !== 'Định dạng đoạn này.')) ===
            JSON.stringify(r.after.paragraphs.filter(p => p.text !== 'Định dạng đoạn này.')) }),
        'Định dạng đoạn này.');
      await test('Insert a table with exact cell values', ['DÒNG GỐC GIỮ NGUYÊN.'],
        'Thêm vào cuối tài liệu một bảng 3 hàng, 2 cột. Hàng 1: Tên | Số lượng. Hàng 2: Bút | 2. Hàng 3: Vở | 3. Giữ nguyên dòng gốc. Không thêm nội dung khác.',
        r => ({ exactTable: r.after.tables.length === 1 &&
          JSON.stringify(r.after.tables[0].values.map(row => row.map(clean))) ===
            JSON.stringify([['Tên', 'Số lượng'], ['Bút', '2'], ['Vở', '3']]),
          originalPreserved: r.after.paragraphs.some(p => p.text === 'DÒNG GỐC GIỮ NGUYÊN.') }));
    } finally {
      const report = { time: new Date().toISOString(), sessionId, provider: status.provider,
        model: status.model, passed: cases.filter(c => c.passed).length, total: cases.length, cases };
      const reportFile = path.join(outputDir, 'report.json');
      fs.writeFileSync(reportFile, JSON.stringify(report, null, 2));
      console.log(JSON.stringify({ passed: report.passed, total: report.total, reportFile }, null, 2));
    }
    if (cases.some(c => !c.passed)) process.exitCode = 1;
    return;
  }
  const text = rest.join(' ');
  const args = action === 'fixture' ? { paragraphs: JSON.parse(text) }
    : action === 'run' ? { prompt: text } : action === 'select' ? { text } : {};
  console.log(JSON.stringify(await execute(sessionId, action, args), null, 2));
}

main().catch(error => { console.error(error.message); process.exitCode = 1; });
