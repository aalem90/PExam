/* Dasbor dosen */
const A = { t: sessionStorage.getItem('adm_t') || '', courses: [], cid: localStorage.getItem('adm_cid') || '', tab: 'mk', fsImgs: [] };
const TABS = [['mk', 'Mata Kuliah'], ['ps', 'Peserta'], ['bk', 'Bank & Generator'], ['qc', 'QC'], ['bb', 'Bobot'], ['fs', 'Formula'], null, ['mon', 'Monitor'], ['vl', 'Pelanggaran'], null, ['gr', 'Koreksi'], ['ch', 'Kecurangan'], ['fn', 'Nilai Akhir']];
const LET = ['A', 'B', 'C', 'D', 'E'];
const cur = () => A.courses.find(c => String(c.id) === String(A.cid));
const yes = v => v === true || /^(true|ya|y|1)$/i.test(String(v).trim());

async function call(fn, ...a) {
  try { return await api(fn, A.t, ...a); }
  catch (e) { if (/LOGIN_DOSEN_HABIS/.test(e.message)) { sessionStorage.removeItem('adm_t'); A.t = ''; showPin('Sesi login habis. Masukkan PIN lagi.'); } throw e; }
}
function toast(msg, kind) {
  const el = $('#toast'); el.innerHTML = `<div class="msg ${kind || 'ok'}">${esc(msg)}</div>`;
  clearTimeout(A.tt); if (kind !== 'err') A.tt = setTimeout(() => el.innerHTML = '', 6000); if (kind === 'err') scrollTo(0, 0);
}
async function run(btn, f) {
  const b = typeof btn === 'string' ? $(btn) : btn, txt = b && b.textContent;
  if (b) { b.disabled = true; b.textContent = 'Memproses…'; }
  try { await f(); } catch (e) { if (!/LOGIN_DOSEN_HABIS/.test(e.message)) toast(e.message, 'err'); }
  finally { if (b) { b.disabled = false; b.textContent = txt; } }
}
function download(name, data, type) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(data instanceof Blob ? data : new Blob([data], { type: type || 'text/plain' })); a.download = name;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
}
const csvCell = v => { v = String(v === undefined || v === null ? '' : v); if (/^[=+\-@]/.test(v)) v = "'" + v; return /[";\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
// CSV untuk Excel berbahasa Indonesia: pemisah titik koma, desimal koma
const toCsv = rows => '\ufeff' + rows.map(r => r.map(v => csvCell(typeof v === 'number' ? String(v).replace('.', ',') : v)).join(';')).join('\r\n');
async function copyText(s) {
  try { await navigator.clipboard.writeText(s); } catch (e) { const t = document.createElement('textarea'); t.value = s; document.body.appendChild(t); t.select(); document.execCommand('copy'); t.remove(); }
  toast('Prompt disalin (' + s.length.toLocaleString('id-ID') + ' karakter). Tempel di chat AI Anda.');
}
function openMod(title, html) { $('#modT').textContent = title; $('#modB').innerHTML = html; $('#mod').classList.add('on'); }
function closeMod() { $('#mod').classList.remove('on'); }
const stBadge = s => `<span class="badge ${({ AKTIF: 'b-green', TERKUNCI: 'b-red', SELESAI: 'b-blue', WAKTU_HABIS: 'b-amber', IMPOR: '' })[s] || ''}">${esc(String(s).replace('_', ' '))}</span>`;
const riskBadge = r => `<span class="badge ${({ High: 'b-red', Medium: 'b-amber', Low: 'b-blue' })[r]}">${r}</span>`;
const durTxt = s => { s = Number(s); if (s === 0) return '0 dtk'; if (!s) return ''; return s < 60 ? s + ' dtk' : Math.floor(s / 60) + ' mnt ' + String(s % 60).padStart(2, '0') + ' dtk'; };
const katBadge = k => `<span class="badge ${({ 'Very Close': 'b-green', 'Close': 'b-green', 'Partially Correct': 'b-amber', 'Far': 'b-red', 'Incorrect': 'b-red' })[k] || ''}">${esc(k || '–')}</span>`;

/* ---------- baca file ---------- */
const readBuf = f => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsArrayBuffer(f); });
async function sheetRows(file, re) {
  const wb = XLSX.read(await readBuf(file), { type: 'array' }), name = wb.SheetNames.find(n => re && re.test(n)) || wb.SheetNames[0];
  return XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '', raw: false });
}
/* ---------- pembaca bank soal dari Word / teks ---------- */
const RX = {
  q: /^(?:no(?:mor)?\.?\s*|soal\s*)?(\d{1,3})\s*[.):]\s*(.*)$/i,
  o: /^(\*?)\s*(?:\(([A-Ea-e])\)|([A-Ea-e])\s*[.)])\s*(.*?)\s*(\*?)$/,
  key: /^(kunci(?:\s*jawaban)?|jawaban(?:\s*benar)?|jwb|answer|key)\s*[:=]\s*(.+)$/i,
  keyHead: /^kunci\s*jawaban\s*:?\s*$/i,
  mlab: /^(bobot|poin|parafrase|gambar)\s*[:=]\s*(.*)$/i,
  secPG: /^(?:bagian\s*[a-z0-9]+\s*[:.\-–—]?\s*)?(?:[a-e1-9][.)]\s*)?(?:soal\s*)?(?:pilihan\s*ganda|multiple\s*choice|pg)\b/i,
  secES: /^(?:bagian\s*[a-z0-9]+\s*[:.\-–—]?\s*)?(?:[a-e1-9][.)]\s*)?(?:soal\s*)?(?:analitik|analisis|hitungan)?\s*[\/&,]?\s*(?:essay|esai|uraian|analitik|analisis)\b/i,
  lab: /^(jawaban\s*inti|jawaban|kunci\s*jawaban|kunci|answer|keyword|keywords|kata\s*kunci|rubrik|rubric|variasi|bobot|poin|skor|parafrase|jebakan|gambar)\s*[:=]\s*(.*)$/i
};
const LABMAP = { 'jawaban inti': 'jawaban_inti', jawaban: 'jawaban_inti', 'kunci jawaban': 'jawaban_inti', kunci: 'jawaban_inti', answer: 'jawaban_inti', keyword: 'keyword', keywords: 'keyword', 'kata kunci': 'keyword', rubrik: 'rubrik', rubric: 'rubrik', poin: 'bobot', skor: 'bobot' };
/** Pasangan kunci "1. A, 2-C, 3) b" dari daftar kunci di akhir dokumen. */
function keyPairs(t) { const out = {}; let m; const re = /(\d{1,3})\s*[.):\-]?\s*([A-Ea-e])(?![A-Za-z])/g; while ((m = re.exec(t))) out[Number(m[1])] = m[2].toUpperCase(); return out; }
/** items: [{t, li, depth, ol, bold}] (dari Word) atau baris teks biasa → baris bank soal. */
const endsQ = t => /[?:…]\s*\)?$|\.\.\.\s*$/.test(t);
/** Jumlah pilihan per soal pada daftar datar (soal & pilihan satu daftar): dari jarak antar-kalimat bertanda soal. */
function inferK(items) {
  const idx = []; items.forEach((it, i) => { if (it.li && !RX.q.test(it.t || '') && !RX.o.test(it.t || '') && endsQ(String(it.t || '').trim())) idx.push(i); });
  const c = {}; for (let i = 1; i < idx.length; i++) { const g = idx[i] - idx[i - 1] - 1; if (g >= 2 && g <= 5) c[g] = (c[g] || 0) + 1; }
  const k = Object.keys(c).sort((a, b) => c[b] - c[a])[0]; return k ? Number(k) : 5;
}
/** items: [{t, li, depth, ol, bold}] (dari Word) atau baris teks biasa → baris bank soal. */
function itemsToMcq(items) {
  const out = []; let q = null, lastOl = null, keyMode = false, keyText = '';
  const K = inferK(items);
  const close = () => { if (q && !q.kunci) { const b = q.opts.filter(o => o.bold || o.star); if (b.length === 1) q.kunci = b[0].L; } };
  const newQ = (t, n) => { close(); q = { n: n || out.length + 1, soal: t, opts: [], kunci: '' }; out.push(q); };
  const addOpt = (t, bold, star, L) => { if (q.opts.length >= 5) return; q.opts.push({ L: L ? L.toUpperCase() : 'ABCDE'[q.opts.length], t, bold, star }); };
  items.forEach(it => {
    const t = String(it.t || '').replace(/\s+/g, ' ').trim(); if (!t) return; let m;
    if (keyMode) { keyText += ' ' + t; return; }
    if (RX.keyHead.test(t)) { keyMode = true; return; }
    if (q && (m = RX.mlab.exec(t))) { const k = m[1].toLowerCase() === 'poin' ? 'bobot' : m[1].toLowerCase(); q[k] = m[2].trim(); return; }
    if ((m = RX.key.exec(t))) { const kp = keyPairs(m[2]); if (Object.keys(kp).length >= 2) { keyText += ' ' + m[2]; return; } if (q) q.kunci = m[2].trim(); return; }
    if ((m = RX.q.exec(t)) && m[2]) { newQ(m[2], Number(m[1])); lastOl = it.ol; return; }
    if (q && (m = RX.o.exec(t)) && m[4]) { addOpt(m[4], it.bold, !!(m[1] || m[5]), m[2] || m[3]); lastOl = it.ol; return; }
    if (it.li) {
      if (it.depth >= 1 && q) addOpt(t, it.bold, false);
      else if (!q || q.kunci || q.opts.length >= K || (it.ol !== lastOl && q.opts.length > 0) || (q.opts.length >= 2 && endsQ(t) && !it.bold)) newQ(t);
      else addOpt(t, it.bold, false);
      lastOl = it.ol; return;
    }
    if (!q) return;                                         // judul/keterangan sebelum soal pertama
    if (!q.opts.length && !endsQ(q.soal) && !q.kunci) q.soal += '\n' + t;   // soal lebih dari satu baris
    else if (!q.kunci && q.opts.length < 5) addOpt(t, it.bold, false);
  });
  close();
  const kp = keyPairs(keyText); out.forEach(x => { if (!x.kunci && kp[x.n]) x.kunci = kp[x.n]; });
  return out.map((x, i) => { const r = { id: 'M' + (i + 1), soal: x.soal, bobot: x.bobot || '', parafrase: x.parafrase || '', gambar: x.gambar || '', kunci: x.kunci.replace(/^([A-Ea-e])\s*[.)]?$/, (_, l) => l.toUpperCase()) }; x.opts.forEach(o => r[o.L] = o.t); return r; });
}
function itemsToEssay(items, inSection) {
  const out = []; let q = null, last = 'soal', gap = false;
  const nq = t => { q = { id: 'E' + (out.length + 1), soal: t, _ol: null }; out.push(q); last = 'soal'; };
  items.forEach(it => {
    const t = String(it.t || '').trim(); if (!t) { gap = true; return; } let m;
    const g = gap; gap = false;
    if ((m = RX.lab.exec(t)) && q) { const k = m[1].toLowerCase().replace(/\s+/g, ' '); last = LABMAP[k] || k; q[last] = m[2].trim(); return; }
    if ((m = RX.q.exec(t)) && m[2]) { nq(m[2]); q._ol = it.ol; return; }
    if (it.li && it.depth === 0 && (!q || last !== 'soal' || it.ol === q._ol)) { nq(t); q._ol = it.ol; return; }
    if (!q) { if (inSection) nq(t); return; }
    if (inSection && g && last === 'soal') { nq(t); return; }
    q[last] = (q[last] ? q[last] + (last === 'soal' ? '\n' : ' ') : '') + t;
  });
  return out.map(x => { delete x._ol; return x; });
}
/** Satu dokumen berisi "Bagian A – Pilihan Ganda" dan "Bagian B – Essay": ambil bagian yang sesuai. */
function section(items, kind) {
  const head = it => { const t = String(it.t || '').trim(); return t.length <= 70 && !RX.q.test(t) ? (RX.secES.test(t) ? 'ESSAY' : RX.secPG.test(t) ? 'MCQ' : '') : ''; };
  const H = items.map(head); if (!H.some(Boolean)) return items;
  const out = []; let cur = '';
  items.forEach((it, i) => { if (H[i]) { cur = H[i]; return; } if (cur === kind || (!cur && !H.includes(kind))) out.push(it); });
  out.found = H.includes(kind); return out;
}
function parseText(text, kind) {
  const items = section(String(text).split(/\r?\n/).map(t => ({ t })), kind);
  return kind === 'ESSAY' ? itemsToEssay(items, items.found) : itemsToMcq(items);
}
/** Word (.docx): baca struktur — tabel soal, daftar bernomor otomatis (bertingkat/datar), teks tebal/garis bawah sebagai kunci. */
async function wordRows(buf, kind) {
  const html = (await mammoth.convertToHtml({ arrayBuffer: buf }, { styleMap: ['u => u'] })).value, d = document.createElement('div'); d.innerHTML = html;
  const SUP = { 0: '⁰', 1: '¹', 2: '²', 3: '³', 4: '⁴', 5: '⁵', 6: '⁶', 7: '⁷', 8: '⁸', 9: '⁹', '-': '⁻', '+': '⁺', o: '°', 'º': '°' }, SUB = { 0: '₀', 1: '₁', 2: '₂', 3: '₃', 4: '₄', 5: '₅', 6: '₆', 7: '₇', 8: '₈', 9: '₉' };
  d.querySelectorAll('sup').forEach(x => { const t = x.textContent; x.replaceWith([...t].every(c => SUP[c]) ? [...t].map(c => SUP[c]).join('') : '^' + t); });
  d.querySelectorAll('sub').forEach(x => { const t = x.textContent; x.replaceWith([...t].every(c => SUB[c]) ? [...t].map(c => SUB[c]).join('') : '_' + t); });
  for (const tb of d.querySelectorAll('table')) {
    const rows = [...tb.rows].map(r => [...r.cells].map(c => c.textContent.replace(/\s+/g, ' ').trim()));
    const hi = rows.findIndex(r => r.some(c => /^(soal|pertanyaan|question|butir soal)$/i.test(c)));
    if (hi >= 0 && rows.length > hi + 1) { const H = rows[hi].map((h, i) => h || 'kolom' + i); return rows.slice(hi + 1).filter(r => r.some(Boolean)).map(r => Object.fromEntries(H.map((h, i) => [h, r[i] || '']))); }
  }
  const items = []; let olId = 0;
  const bold = el => { const all = el.textContent.replace(/\s/g, ''); if (!all) return false; const b = [...el.querySelectorAll('strong,b,u')].map(x => x.textContent).join('').replace(/\s/g, ''); return b.length >= all.length * 0.8; };
  const walk = (nodes, depth) => [...nodes].forEach(n => {
    const tag = n.tagName;
    if (tag === 'OL' || tag === 'UL') { const id = ++olId; [...n.children].forEach(li => { const c = li.cloneNode(true); c.querySelectorAll('ol,ul').forEach(x => x.remove()); items.push({ t: c.textContent, li: true, depth, ol: id, bold: bold(c) }); walk(li.querySelectorAll(':scope > ol, :scope > ul'), depth + 1); }); }
    else if (tag === 'TABLE') [...n.rows].forEach(r => items.push({ t: [...r.cells].map(c => c.textContent.trim()).join(' ') }));
    else n.textContent.split(/\n/).forEach(t => items.push({ t, bold: bold(n) }));
  });
  walk(d.children, 0);
  const sec = section(items, kind);
  return kind === 'ESSAY' ? itemsToEssay(sec, sec.found) : itemsToMcq(sec);
}
async function bankRows(file, text, kind) {
  if (file) {
    const n = file.name.toLowerCase();
    if (/\.(xlsx|xls|csv)$/.test(n)) return sheetRows(file, kind === 'ESSAY' ? /essay|esai|uraian/i : /mcq|pg|pilihan|ganda/i);
    if (/\.docx$/.test(n)) return wordRows(await readBuf(file), kind);
    return parseText(await file.text(), kind);
  }
  return parseText(text, kind);
}

/* ---------- login & kerangka ---------- */
function showPin(msg) { $('#vApp').style.display = 'none'; $('#vPin').style.display = ''; $('#pinMsg').innerHTML = msg ? `<div class="msg err">${esc(msg)}</div>` : ''; $('#pin').value = ''; $('#pin').focus(); }
$('#bPin').onclick = () => run('#bPin', async () => {
  try { const r = await api('adminLogin', $('#pin').value.trim(), deviceId()); A.t = r.t; sessionStorage.setItem('adm_t', A.t); await boot(); }
  catch (e) { $('#pinMsg').innerHTML = `<div class="msg err">${esc(e.message)}</div>`; }
});
$('#pin').onkeydown = e => { if (e.key === 'Enter') $('#bPin').click(); };
$('#bOut').onclick = () => { sessionStorage.removeItem('adm_t'); A.t = ''; showPin(); };
async function boot() {
  const d = await call('adminInit'); A.courses = d.courses;
  if (!cur()) A.cid = A.courses.length ? String(A.courses[0].id) : '';
  $('#vPin').style.display = 'none'; $('#vApp').style.display = '';
  $('#tabs').innerHTML = TABS.map(t => t ? `<button data-t="${t[0]}">${t[1]}</button>` : '<span class="sep"></span>').join('');
  fillSel(); go(A.courses.length ? A.tab : 'mk');
}
function fillSel() {
  $('#selMk').innerHTML = A.courses.length ? A.courses.map(c => `<option value="${esc(c.id)}">${esc(c.nama)}</option>`).join('') : '<option value="">(belum ada mata kuliah)</option>';
  $('#selMk').value = A.cid; const c = cur();
  $('#mkMeta').innerHTML = c ? `${esc(c.durasi_menit)} menit • ${String(c.mode_jawaban).toUpperCase() === 'ONLINE' ? 'jawab online' : 'jawab di kertas'} • ${yes(c.aktif) ? '<span class="badge b-green">aktif</span>' : '<span class="badge">nonaktif</span>'}` : '';
}
async function reloadCourses() { A.courses = (await call('adminInit')).courses; fillSel(); }
$('#selMk').onchange = () => { A.cid = $('#selMk').value; localStorage.setItem('adm_cid', A.cid); A.pk = A.res = A.ch = A.mon = null; fillSel(); go(A.tab); };
$('#tabs').onclick = e => { const b = e.target.closest('button'); if (b) go(b.dataset.t); };
function go(tab) {
  A.tab = tab; $$('#tabs button').forEach(b => b.classList.toggle('on', b.dataset.t === tab)); $$('section.tab').forEach(s => s.classList.toggle('on', s.id === 't-' + tab));
  clearInterval(A.monT); $('#toast').innerHTML = '';
  if (tab !== 'mk' && !cur()) { toast('Buat mata kuliah terlebih dahulu.', 'warn'); return go('mk'); }
  run(null, LOAD[tab]);
}
const LOAD = {};

/* ---------- 1. mata kuliah ---------- */
const dtIn = v => { v = String(v || '').trim().replace(' ', 'T'); return /^\d{4}-\d\d-\d\dT\d\d:\d\d/.test(v) ? v.slice(0, 16) : ''; };
LOAD.mk = async () => { A.newMk = !cur(); fillMk(); };
function fillMk() {
  const c = A.newMk ? null : cur();
  $('#mkFormTitle').textContent = c ? 'Ubah: ' + c.nama : 'Mata kuliah baru';
  $('#mkNama').value = c ? c.nama : ''; $('#mkDur').value = c ? c.durasi_menit : 55; $('#mkMaks').value = c ? (c.maks_pelanggaran || 0) : 0;
  $('#mkMode').value = c && String(c.mode_jawaban).toUpperCase() === 'ONLINE' ? 'ONLINE' : 'KERTAS';
  $('#mkBuka').value = c ? dtIn(c.buka) : ''; $('#mkTutup').value = c ? dtIn(c.tutup) : ''; $('#mkAktif').checked = c ? yes(c.aktif) : false; $('#bMkDel').style.display = c ? '' : 'none';
  $('#mkList').innerHTML = '<tr><th>Mata kuliah</th><th class="r">Durasi</th><th>Tipe</th><th>Status</th></tr>' + (A.courses.map(x => `<tr class="click" data-id="${esc(x.id)}"><td><b>${esc(x.nama)}</b><div class="sm mut">${esc(x.id)}</div></td><td class="r num">${esc(x.durasi_menit)} mnt</td><td class="num">${esc(x.jumlah_tipe)}</td><td>${yes(x.aktif) ? '<span class="badge b-green">aktif</span>' : '<span class="badge">nonaktif</span>'} <span class="badge">${String(x.mode_jawaban).toUpperCase() === 'ONLINE' ? 'online' : 'kertas'}</span></td></tr>`).join('') || '<tr><td colspan="4" class="mut">Belum ada.</td></tr>');
}
$('#mkList').onclick = e => { const r = e.target.closest('tr[data-id]'); if (r) { A.cid = r.dataset.id; localStorage.setItem('adm_cid', A.cid); A.newMk = false; A.pk = A.res = A.ch = null; fillSel(); fillMk(); } };
$('#bMkNew').onclick = () => { A.newMk = true; fillMk(); $('#mkNama').focus(); };
$('#bMkSave').onclick = () => run('#bMkSave', async () => {
  const r = await call('adminSaveCourse', { id: A.newMk ? '' : A.cid, nama: $('#mkNama').value.trim(), durasi_menit: Number($('#mkDur').value), maks_pelanggaran: Number($('#mkMaks').value) || 0, buka: $('#mkBuka').value, tutup: $('#mkTutup').value, aktif: $('#mkAktif').checked, mode_jawaban: $('#mkMode').value });
  A.cid = r.id; localStorage.setItem('adm_cid', A.cid); A.newMk = false; await reloadCourses(); fillMk(); toast('Mata kuliah tersimpan.');
});
$('#bMkDel').onclick = () => { const c = cur(); if (c && confirm('Hapus mata kuliah "' + c.nama + '" beserta bank soal, paket, dan formula sheet-nya?')) run('#bMkDel', async () => { await call('adminDeleteCourse', A.cid); A.cid = ''; await reloadCourses(); if (!cur()) A.cid = A.courses.length ? String(A.courses[0].id) : ''; fillSel(); A.newMk = !cur(); fillMk(); toast('Mata kuliah dihapus.'); }); };

/* ---------- 2. peserta ---------- */
LOAD.ps = async () => {
  A.roster = await call('adminRoster'); let T = {};
  try { const L = await call('adminTipeList', A.cid); L.list.forEach(x => T[x.nim] = x.tipe); } catch (e) { T = null; }
  const mine = r => !String(r.course || '').trim() || String(r.course).split(/[,;\s]+/).includes(String(A.cid));
  $('#psTbl').innerHTML = `<tr><th class="r">No</th><th>NIM</th><th>Nama</th><th>Kelas</th><th>Password</th>${T ? '<th>Tipe</th>' : ''}<th>Dibatasi ke</th></tr>` + (A.roster.map((r, i) => `<tr${mine(r) ? '' : ' style="opacity:.5"'}><td class="r num">${i + 1}</td><td class="mono">${esc(r.nim)}</td><td>${esc(r.nama)}</td><td>${esc(r.kelas)}</td><td class="mono"><b>${esc(r.password)}</b></td>${T ? `<td class="num">${T[String(r.nim)] || '–'}</td>` : ''}<td class="mut">${esc(r.course) || 'semua'}</td></tr>`).join('') || '<tr><td colspan="7" class="mut">Belum ada peserta.</td></tr>');
};
$('#bPsImp').onclick = () => run('#bPsImp', async () => {
  const f = $('#psFile').files[0]; let rows;
  if (f) rows = await rosterRows(f);
  else rows = $('#psPaste').value.split(/\r?\n/).map(l => l.split(/\t|;|,/).map(x => x.trim())).filter(a => a.length >= 2 && a[0]).map(a => ({ nim: a[0], nama: a[1], kelas: a[2] || '', password: a[3] || '' })).filter(r => !/^nim$/i.test(r.nim));
  if (!rows.length) throw new Error('Tidak ada baris peserta yang terbaca.');
  const r = await call('adminImportRoster', rows, $('#psRep').checked); $('#psFile').value = ''; $('#psPaste').value = '';
  await LOAD.ps(); toast(`${r.tambah} peserta baru, ${r.perbarui} diperbarui. Total ${r.total}.`);
});
$('#bPsGen').onclick = () => run('#bPsGen', async () => { const r = await call('adminGenPasswords', false); await LOAD.ps(); toast(r.dibuat + ' password dibuat.'); });
$('#bPsCsv').onclick = () => download('password_ujian.csv', toCsv([['NIM', 'Nama', 'Kelas', 'Password']].concat((A.roster || []).map(r => [r.nim, r.nama, r.kelas, r.password]))), 'text/csv');

/* ---------- 3. bank & generator ---------- */
LOAD.bk = async () => {
  const d = await call('adminGetBank', A.cid), c = d.course; A.bank = d;
  $('#bkMcqInfo').textContent = d.mcq.length ? d.mcq.length + ' soal di bank.' : 'Bank masih kosong.'; $('#bkEsInfo').textContent = d.essay.length ? d.essay.length + ' soal di bank.' : 'Bank masih kosong.';
  $('#gTipe').value = c.jumlah_tipe || 4; $('#gNm').value = c.n_mcq || 0; $('#gNe').value = c.n_essay || 0; $('#gPara').value = String(c.parafrase || 'KAMUS').toUpperCase();
  $('#gAs').checked = yes(c.acak_soal); $('#gAo').checked = yes(c.acak_opsi); $('#gTrap').checked = yes(c.jebakan_aktif); $('#gTrapMode').value = c.jebakan_mode || 'KEDUANYA';
  $('#gTrapTpl').value = c.jebakan_template || ''; $('#gTrapKata').value = c.jebakan_kata || ''; $('#trapBox').style.display = $('#gTrap').checked ? '' : 'none';
  renderEsList(d.essay);
  let m = null; try { m = JSON.parse(String(c.essay_tipe || '').trim() || 'null'); } catch (e) { m = null; }
  A.esMap = m && Object.keys(m).length ? m : null;
  $$('input[name=esMode]').forEach(r => r.checked = r.value === (A.esMap ? 'manual' : 'auto')); renderEsMap();
};
/* ---- pembagian essay per tipe ---- */
const esActive = () => ((A.bank && A.bank.essay) || []).filter(q => String(q.aktif) === '' || yes(q.aktif));
function esMapGet() { const m = {}; const n = Number($('#gTipe').value) || 1; for (let k = 1; k <= n; k++) m[k] = []; $$('#esMapTbl input[data-k]:checked').forEach(x => m[x.dataset.k].push(x.dataset.e)); return m; }
function renderEsMap() {
  const manual = ($('input[name=esMode]:checked') || {}).value === 'manual', es = esActive(), n = Number($('#gTipe').value) || 1;
  $('#esMapBox').style.display = manual ? '' : 'none'; $('#gNe').disabled = manual; if (!manual) return;
  const cur = A.esMap || {}, has = (k, id) => (cur[k] || cur[String(k)] || []).map(String).includes(String(id));
  const cnt = k => es.filter(q => has(k, q.id)).length;
  if (!es.length) { $('#esMapTbl').innerHTML = '<tr><td class="mut">Belum ada soal essay di bank.</td></tr>'; return; }
  $('#esMapTbl').innerHTML = `<tr><th>Soal essay</th><th class="r">Poin</th>${Array.from({ length: n }, (_, i) => `<th style="text-align:center">Tipe ${i + 1}</th>`).join('')}</tr>` +
    es.map(q => `<tr><td><b class="mono">${esc(q.id)}</b> <span class="sm">${esc(String(q.soal).slice(0, 90))}</span></td><td class="r num">${fmtNum(q.bobot === '' ? 1 : parseNum(q.bobot))}</td>${Array.from({ length: n }, (_, i) => `<td style="text-align:center"><input type="checkbox" data-k="${i + 1}" data-e="${esc(q.id)}" ${has(i + 1, q.id) ? 'checked' : ''} aria-label="${esc(q.id)} untuk tipe ${i + 1}"></td>`).join('')}</tr>`).join('') +
    `<tr><td class="sm mut" colspan="2">Jumlah essay per tipe</td>${Array.from({ length: n }, (_, i) => { const c = cnt(i + 1); return `<td style="text-align:center"><span class="badge ${c ? 'b-blue' : 'b-red'}">${c}</span></td>`; }).join('')}</tr>`;
}
const esMapSet = m => { A.esMap = m; renderEsMap(); };
$('#esMapTbl').onchange = () => { A.esMap = esMapGet(); renderEsMap(); };
$$('input[name=esMode]').forEach(r => r.onchange = () => { if (r.value === 'manual' && !A.esMap) { const m = {}, n = Number($('#gTipe').value) || 1; for (let k = 1; k <= n; k++) m[k] = esActive().map(q => String(q.id)); A.esMap = m; } renderEsMap(); });
$('#gTipe').addEventListener('change', renderEsMap);
$('#bEsMapAll').onclick = () => { const m = {}, n = Number($('#gTipe').value) || 1; for (let k = 1; k <= n; k++) m[k] = esActive().map(q => String(q.id)); esMapSet(m); };
$('#bEsMapNone').onclick = () => { const m = {}, n = Number($('#gTipe').value) || 1; for (let k = 1; k <= n; k++) m[k] = []; esMapSet(m); };
$('#bEsMapRR').onclick = () => { const m = {}, n = Number($('#gTipe').value) || 1; for (let k = 1; k <= n; k++) m[k] = []; esActive().forEach((q, i) => m[(i % n) + 1].push(String(q.id))); esMapSet(m); };
function renderEsList(list) {
  $('#esWrap').style.display = list.length ? '' : 'none';
  $('#esTbl').innerHTML = '<tr><th>ID</th><th>Soal essay</th><th class="r">Poin</th><th></th></tr>' + list.map(q => `<tr><td class="mono">${esc(q.id)}</td><td>${esc(String(q.soal).slice(0, 120))}${String(q.jawaban_inti).trim() ? '' : ' <span class="badge b-amber">belum ada jawaban inti</span>'}</td><td class="r num">${fmtNum(q.bobot === '' ? 1 : parseNum(q.bobot))}</td><td><div class="row" style="gap:4px"><button class="btn sec xs" data-esed="${esc(q.id)}">Edit</button><button class="btn sec xs" data-esdel="${esc(q.id)}">Hapus</button></div></td></tr>`).join('');
}
function essayForm(q) {
  q = q || {};
  openMod(q.id ? 'Edit soal essay ' + q.id : 'Tulis soal essay baru', `
    <label>Soal <span class="sm mut" style="font-weight:400">— rumus pakai $…$, mis. $\phi = V_p / V_b$</span></label><textarea id="eqSoal" style="min-height:110px">${esc(q.soal)}</textarea>
    <div class="grid" style="grid-template-columns:140px 1fr"><div><label>Poin soal ini</label><input type="text" inputmode="decimal" id="eqPoin" value="${esc(q.bobot === undefined || q.bobot === '' ? '' : fmtNum(parseNum(q.bobot)))}" placeholder="mis. 10"></div>
      <div><label>Jawaban inti</label><textarea id="eqInti" style="min-height:70px">${esc(q.jawaban_inti)}</textarea></div></div>
    <label>Keyword <span class="sm mut" style="font-weight:400">— konsep kunci dipisah titik koma; padanan dengan |; bobot konsep dengan *2. Contoh: porositas; 28,8*2; volume pori|Vp</span></label><input type="text" id="eqKw" value="${esc(q.keyword)}">
    <label>Rubrik (opsional)</label><input type="text" id="eqRub" value="${esc(q.rubrik)}" placeholder="mis. 30% rumus; 30% substitusi; 40% hasil akhir">
    <label>Variasi kalimat soal per tipe (opsional, pisahkan dengan ||)</label><input type="text" id="eqVar" value="${esc(q.variasi)}">
    <div class="msg info sm">Poin yang diisi di sini membuat poin essay diatur <b>per soal</b>. Pastikan total PG + essay tetap 100 di tab <b>Bobot</b>.</div>
    <div class="row"><button class="btn" id="eqSave">Simpan soal</button></div>`);
  $('#eqSave').onclick = () => run('#eqSave', async () => {
    const poin = $('#eqPoin').value.trim(); if (poin && (isNaN(parseNum(poin)) || parseNum(poin) < 0)) throw new Error('Poin tidak valid. Contoh: 10 atau 7,5');
    const r = await call('adminSaveBankItem', A.cid, 'ESSAY', { id: q.id || '', soal: $('#eqSoal').value, jawaban_inti: $('#eqInti').value, keyword: $('#eqKw').value, rubrik: $('#eqRub').value, variasi: $('#eqVar').value, bobot: poin ? parseNum(poin) : undefined });
    closeMod(); A.pk = null; await reloadCourses(); await LOAD.bk();
    const tot = r.total, ok = Math.abs(tot.total - 100) < 0.01;
    toast(`Soal ${r.id} tersimpan.` + (r.diPaket ? '' : ' Klik "Simpan & buat tipe ujian" agar soal ini masuk ke tipe ujian.') + ` Total poin saat ini ${fmtNum(tot.total)}` + (ok ? ' ✓' : ' — sesuaikan di tab Bobot agar 100.'), ok ? 'ok' : 'warn');
  });
}
$('#bEsAdd').onclick = () => essayForm(null);
$('#esTbl').onclick = e => {
  const ed = e.target.closest('[data-esed]'), del = e.target.closest('[data-esdel]');
  if (ed) essayForm((A.bank.essay || []).find(q => String(q.id) === ed.dataset.esed));
  if (del && confirm('Hapus soal essay ' + del.dataset.esdel + ' dari bank?')) run(del, async () => { const r = await call('adminDeleteBankItem', A.cid, 'ESSAY', del.dataset.esdel); A.pk = null; await LOAD.bk(); toast('Soal dihapus. Buat ulang tipe ujian agar perubahan berlaku. Total poin saat ini ' + fmtNum(r.total.total) + '.', 'warn'); });
};
$('#gTrap').onchange = () => $('#trapBox').style.display = $('#gTrap').checked ? '' : 'none';
async function impBank(kind, fileSel, textSel, repSel, btn) {
  await run(btn, async () => {
    const rows = await bankRows($(fileSel).files[0], $(textSel).value, kind);
    if (!rows.length) throw new Error('Tidak ada soal yang terbaca. Periksa format file/teks.');
    const r = await call('adminImportBank', A.cid, kind, rows, $(repSel).checked); $(fileSel).value = ''; $(textSel).value = ''; await LOAD.bk();
    toast(`${r.masuk} soal ${kind === 'ESSAY' ? 'essay' : 'pilihan ganda'} masuk (total ${r.total}).` + (r.peringatan.length ? ' Perlu diperbaiki: ' + r.peringatan.join('; ') : ''), r.peringatan.length ? 'warn' : 'ok');
  });
}
$('#bBkMcq').onclick = () => impBank('MCQ', '#bkMcqFile', '#bkMcqText', '#bkMcqRep', '#bBkMcq');
$('#bBkEs').onclick = () => impBank('ESSAY', '#bkEsFile', '#bkEsText', '#bkEsRep', '#bBkEs');
$('#bTpl').onclick = () => { location.href = 'template_bank_soal.xlsx'; };
$('#bTplW').onclick = () => { location.href = 'template_bank_soal.docx'; };
$('#bGen').onclick = () => run('#bGen', async () => {
  const c = cur(), manual = ($('input[name=esMode]:checked') || {}).value === 'manual';
  let essay_tipe = '';
  if (manual) {
    const m = esMapGet(), kosong = Object.keys(m).filter(k => !m[k].length);
    if (esActive().length && kosong.length) throw new Error('Tipe ' + kosong.join(', ') + ' belum diberi soal essay. Centang minimal satu soal essay untuk tiap tipe.');
    essay_tipe = JSON.stringify(m);
  }
  await call('adminSaveCourse', { id: A.cid, nama: c.nama, durasi_menit: c.durasi_menit, jumlah_tipe: Number($('#gTipe').value), n_mcq: Number($('#gNm').value) || 0, n_essay: Number($('#gNe').value) || 0, parafrase: $('#gPara').value,
    acak_soal: $('#gAs').checked, acak_opsi: $('#gAo').checked, jebakan_aktif: $('#gTrap').checked, jebakan_mode: $('#gTrapMode').value, jebakan_template: $('#gTrapTpl').value.trim(), jebakan_kata: $('#gTrapKata').value.trim(), essay_tipe });
  try { A.pk = await call('adminGenerate', A.cid, false); }
  catch (e) { if (!/SUDAH_ADA_SESI/.test(e.message)) throw e; if (!confirm('Sudah ada mahasiswa yang mulai/selesai ujian ini. Membuat ulang akan mengubah soal dan kunci mereka.\n\nTetap buat ulang?')) return; A.pk = await call('adminGenerate', A.cid, true); }
  await reloadCourses(); const er = A.pk.qc.filter(q => q.level === 'ERROR').length; go('qc');
  setTimeout(() => toast(`Tipe ujian dibuat. ${er ? er + ' masalah perlu diperbaiki.' : 'QC otomatis tidak menemukan kesalahan — periksa parafrase yang ditandai.'}`, er ? 'warn' : 'ok'), 50);
});
$('#bParaCopy').onclick = () => run('#bParaCopy', async () => copyText((await call('adminParaPrompt', A.cid)).prompt));
$('#bParaApply').onclick = () => run('#bParaApply', async () => { const r = await call('adminApplyPara', A.cid, $('#paraJson').value); $('#paraJson').value = ''; toast(`Usulan parafrase disimpan untuk ${r.soal} soal (${r.ditolak} usulan ditolak karena frasanya tidak ada di soal). Klik "Simpan & buat tipe ujian" untuk memakainya.`); });

/* ---------- 4. QC ---------- */
LOAD.qc = async () => { if (!A.pk) A.pk = await call('adminGetPaket', A.cid); A.qcT = A.qcT || 1; renderQc(); };
$('#bQcReload').onclick = () => run('#bQcReload', async () => { A.pk = await call('adminGetPaket', A.cid); renderQc(); });
$('#qcOnly').onchange = renderQc;
function markPara(p) {
  let t = String(p.soal), log = []; try { log = JSON.parse(p.parafrase || '[]'); } catch (e) { }
  if (Array.isArray(log)) log.forEach(x => { const i = t.indexOf(x.jadi); if (i >= 0) t = t.slice(0, i) + '\u0002' + x.jadi + '\u0003' + t.slice(i + x.jadi.length); });
  return rich(t).replace(/\u0002/g, '<mark title="hasil parafrase">').replace(/\u0003/g, '</mark>');
}
function renderQc() {
  const P = A.pk.paket, tipes = [...new Set(P.map(p => Number(p.tipe)))].sort();
  if (!tipes.length) { $('#qcTipe').innerHTML = ''; $('#qcSum').innerHTML = ''; $('#qcList').innerHTML = '<div class="card mut">Belum ada tipe ujian. Buat dulu di tab Bank &amp; Generator.</div>'; return; }
  if (!tipes.includes(A.qcT)) A.qcT = tipes[0];
  const cnt = l => A.pk.qc.filter(q => q.level === l).length;
  $('#qcTipe').innerHTML = tipes.map(k => { const e = A.pk.qc.filter(q => q.tipe === k && q.level === 'ERROR').length; return `<button class="chip ${k === A.qcT ? 'on' : ''}" data-k="${k}">Tipe ${k}${e ? ' ⚠' + e : ''}</button>`; }).join('');
  $('#qcSum').innerHTML = `<span class="badge ${cnt('ERROR') ? 'b-red' : 'b-green'}">${cnt('ERROR')} error</span> <span class="badge b-amber">${cnt('CEK')} perlu dicek</span> <span class="badge">${cnt('INFO')} info</span>`;
  const M = Object.fromEntries(A.pk.mcq.map(q => [String(q.id), q])), E = Object.fromEntries(A.pk.essay.map(q => [String(q.id), q])), only = $('#qcOnly').checked;
  const glob = A.pk.qc.filter(q => q.no === 0).map(q => `<div class="msg err">${esc(q.pesan)}</div>`).join('');
  $('#qcList').innerHTML = glob + P.filter(p => Number(p.tipe) === A.qcT).sort((a, b) => a.no - b.no).map(p => {
    const iss = A.pk.qc.filter(q => q.tipe === Number(p.tipe) && q.no === Number(p.no)), err = iss.some(q => q.level === 'ERROR');
    if (only && !iss.some(q => q.level !== 'INFO')) return '';
    const b = (p.jenis === 'MCQ' ? M : E)[String(p.bank_id)] || {}, asli = p.jenis === 'MCQ' && b.soal && String(b.soal) !== String(p.soal) ? `<div class="sm mut" style="margin:4px 0">Asli di bank: ${rich(b.soal)}</div>` : '';
    let body = '';
    if (p.jenis === 'MCQ') { let o = []; try { o = JSON.parse(p.opsi_json || '[]'); } catch (e) { } body = o.map(x => `<div class="o ${x.l === p.kunci ? 'k' : ''}"><span style="width:22px">${x.l}.</span><span class="grow">${rich(x.t)}</span><span class="sm mut">asal ${x.asal}</span>${x.l === p.kunci ? '<span class="badge b-green">kunci</span>' : ''}</div>`).join(''); }
    else body = `<div class="sm"><b>Jawaban inti:</b> ${esc(b.jawaban_inti || '–')}</div><div class="sm"><b>Keyword:</b> ${esc(b.keyword || '–')}</div><div class="sm"><b>Rubrik:</b> ${esc(b.rubrik || '–')}</div>` + (p.jebakan ? `<div class="msg warn sm"><b>Jebakan AI</b> (istilah: <b>${esc(p.jebakan_kata)}</b>): ${esc(p.jebakan)}</div>` : '');
    return `<div class="qc ${err ? 'err' : ''}"><div class="row"><span class="badge b-blue">No ${p.no}</span><span class="badge">${p.jenis === 'MCQ' ? 'PG' : 'Essay'}</span><span class="badge">bank ${esc(p.bank_id)}</span><span class="badge b-amber">${fmtNum(A.pk.bobot[p.tipe + ':' + p.no] ?? '–')} poin</span><button class="btn sec xs right" data-edit="${p.no}">Edit</button></div>
      <div style="margin:8px 0;font-size:15px">${markPara(p)}</div>${asli}${body}${iss.map(q => `<div class="sm" style="margin-top:4px"><span class="badge ${q.level === 'ERROR' ? 'b-red' : q.level === 'CEK' ? 'b-amber' : ''}">${q.level}</span> ${esc(q.pesan)}</div>`).join('')}</div>`;
  }).join('');
}
$('#qcTipe').onclick = e => { const b = e.target.closest('.chip'); if (b) { A.qcT = Number(b.dataset.k); renderQc(); } };
$('#qcList').onclick = e => { const b = e.target.closest('[data-edit]'); if (b) editQ(Number(b.dataset.edit)); };
function editQ(no) {
  const p = A.pk.paket.find(x => Number(x.tipe) === A.qcT && Number(x.no) === no); let h = `<label>Teks soal (tipe ${A.qcT}, no ${no})</label><textarea id="eSoal">${esc(p.soal)}</textarea>`;
  if (p.jenis === 'MCQ') {
    const o = JSON.parse(p.opsi_json || '[]');
    h += '<label>Pilihan jawaban — pilih bulatan untuk kunci</label>' + o.map(x => `<div class="row" style="margin:6px 0"><input type="radio" name="eK" value="${x.l}" ${x.l === p.kunci ? 'checked' : ''} aria-label="Kunci ${x.l}"><b style="width:18px">${x.l}</b><input type="text" class="grow eO" data-asal="${x.asal}" value="${esc(x.t)}"></div>`).join('') +
      `<label class="chk"><input type="checkbox" id="eAll"> <span>Kunci di <b>bank soal</b> memang salah → perbaiki kunci di bank dan di <b>semua tipe</b> sekaligus</span></label>`;
  } else h += `<label>Kalimat jebakan AI (kosongkan untuk tanpa jebakan)</label><textarea id="eTrap" style="min-height:60px">${esc(p.jebakan)}</textarea><label>Istilah unik yang dicari di jawaban</label><input type="text" id="eKata" value="${esc(p.jebakan_kata)}">`;
  openMod('Edit soal', h + '<div class="row" style="margin-top:12px"><button class="btn" id="eSave">Simpan</button><span class="sm mut">Perubahan teks hanya berlaku untuk tipe ini.</span></div>');
  $('#eSave').onclick = () => run('#eSave', async () => {
    const f = { soal: $('#eSoal').value };
    if (p.jenis === 'MCQ') {
      const k = ($('input[name=eK]:checked') || {}).value; if (!k) throw new Error('Pilih kunci jawaban.');
      f.opsi = $$('.eO').map(i => ({ t: i.value })); f.kunci = k;
      const asal = $$('.eO')[LET.indexOf(k)].dataset.asal, teksUbah = f.soal !== String(p.soal) || JSON.stringify(f.opsi.map(x => x.t)) !== JSON.stringify(JSON.parse(p.opsi_json).map(x => x.t));
      if ($('#eAll').checked) { await call('adminFixKey', A.cid, p.bank_id, asal); if (teksUbah) await call('adminSavePaketRow', A.cid, A.qcT, no, f); }
      else await call('adminSavePaketRow', A.cid, A.qcT, no, f);
    } else { f.jebakan = $('#eTrap').value.trim(); f.jebakan_kata = f.jebakan ? $('#eKata').value.trim() : ''; await call('adminSavePaketRow', A.cid, A.qcT, no, f); }
    A.pk = await call('adminGetPaket', A.cid); closeMod(); renderQc(); toast('Perubahan disimpan.');
  });
}
$('#bQcXls').onclick = () => {
  if (!A.pk || !A.pk.paket.length) return toast('Belum ada tipe ujian.', 'warn');
  const wb = XLSX.utils.book_new();
  [...new Set(A.pk.paket.map(p => Number(p.tipe)))].sort().forEach(k => {
    const rows = [['No', 'Jenis', 'ID bank', 'Soal', 'A', 'B', 'C', 'D', 'E', 'Kunci', 'Bobot']];
    A.pk.paket.filter(p => Number(p.tipe) === k).sort((a, b) => a.no - b.no).forEach(p => { let o = []; try { o = JSON.parse(p.opsi_json || '[]'); } catch (e) { } rows.push([p.no, p.jenis, p.bank_id, p.soal, ...LET.map(L => (o.find(x => x.l === L) || {}).t || ''), p.kunci, A.pk.bobot[k + ':' + p.no] ?? '']); });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Tipe ' + k);
  });
  XLSX.writeFile(wb, 'paket_' + A.cid + '.xlsx');
};

/* ---------- 5. bobot ---------- */
LOAD.bb = async () => {
  const [d, pk] = await Promise.all([call('adminGetBank', A.cid), call('adminGetPaket', A.cid)]), c = d.course; A.bank = d; A.pk = pk;
  const act = q => String(q.aktif) === '' || yes(q.aktif), M = Object.fromEntries(d.mcq.map(q => [String(q.id), q])), E = Object.fromEntries(d.essay.map(q => [String(q.id), q]));
  const tp = [...new Set(pk.paket.map(p => Number(p.tipe)))].sort((a, b) => a - b), tipeEs = {};
  let mcqRows;
  if (tp.length) {
    mcqRows = pk.paket.filter(p => Number(p.tipe) === tp[0] && p.jenis === 'MCQ').sort((a, b) => a.no - b.no).map(p => ({ no: p.no, q: M[String(p.bank_id)] || { id: p.bank_id, soal: p.soal, bobot: 0 } }));
    tp.forEach(k => tipeEs[k] = pk.paket.filter(p => Number(p.tipe) === k && p.jenis === 'ESSAY').map(p => String(p.bank_id)));
  } else {
    mcqRows = d.mcq.filter(act).map((q, i) => ({ no: i + 1, q }));
    let m = null; try { m = JSON.parse(String(c.essay_tipe || '').trim() || 'null'); } catch (e) { }
    const n = Math.max(1, Number(c.jumlah_tipe) || 1), ea = d.essay.filter(act).map(q => String(q.id));
    for (let k = 1; k <= n; k++) tipeEs[k] = m && Object.keys(m).length ? (m[k] || m[String(k)] || []).map(String).filter(id => ea.includes(id)) : ea;
  }
  A.bbTipeEs = tipeEs;
  const used = [...new Set(Object.values(tipeEs).flat())], esRows = d.essay.filter(q => used.includes(String(q.id))).map(q => ({ q, tipe: Object.keys(tipeEs).filter(k => tipeEs[k].includes(String(q.id))) }));
  const mm = String(c.mode_mcq || (c.mode_bobot === 'per_soal' ? 'per_soal' : 'rata')), me = String(c.mode_essay || (c.mode_bobot === 'per_soal' ? 'per_soal' : 'rata'));
  $$('input[name=bbModeM]').forEach(r => r.checked = r.value === (mm === 'per_soal' ? 'per_soal' : 'rata')); $$('input[name=bbModeE]').forEach(r => r.checked = r.value === (me === 'per_soal' ? 'per_soal' : 'rata'));
  $('#bbM').value = fmtNum(c.bobot_mcq === '' ? 60 : parseNum(c.bobot_mcq)); $('#bbE').value = fmtNum(c.bobot_essay === '' ? 40 : parseNum(c.bobot_essay));
  const nT = Object.keys(tipeEs).length, beda = new Set(Object.values(tipeEs).map(x => x.slice().sort().join())).size > 1;
  $('#bbNM').textContent = '(' + mcqRows.length + ' soal)'; $('#bbNE').textContent = beda ? '(berbeda per tipe: ' + Object.keys(tipeEs).map(k => 'T' + k + '=' + tipeEs[k].length).join(', ') + ')' : '(' + (Object.values(tipeEs)[0] || []).length + ' soal)';
  $('#bbBoxM').style.opacity = mcqRows.length ? 1 : .5; $('#bbBoxE').style.opacity = esRows.length ? 1 : .5; if (!esRows.length && !c.bobot_essay) $('#bbE').value = '0';
  const poinIn = q => `<input type="text" inputmode="decimal" class="bbW" value="${esc(fmtNum(q.bobot === '' ? 1 : parseNum(q.bobot)))}" style="text-align:right">`;
  $('#bbTbl').innerHTML = '<tr><th class="r">No</th><th>Jenis</th><th>ID</th><th>Soal</th><th>Tipe</th><th class="r" style="width:120px">Poin</th></tr>' +
    (mcqRows.map(r => `<tr data-j="MCQ" data-id="${esc(r.q.id)}"><td class="r num">${r.no}</td><td>PG</td><td class="mono">${esc(r.q.id)}</td><td>${esc(String(r.q.soal).slice(0, 100))}</td><td class="sm mut">semua</td><td class="r">${poinIn(r.q)}</td></tr>`).join('') +
     esRows.map(r => `<tr data-j="ESSAY" data-id="${esc(r.q.id)}" data-tipe="${r.tipe.join(',')}"><td class="r num">–</td><td>Essay</td><td class="mono">${esc(r.q.id)}</td><td>${esc(String(r.q.soal).slice(0, 100))}</td><td class="sm">${r.tipe.length === nT ? 'semua' : r.tipe.map(k => 'T' + k).join(', ')}</td><td class="r">${poinIn(r.q)}</td></tr>`).join('') || '<tr><td colspan="6" class="mut">Bank soal masih kosong.</td></tr>');
  calcBb();
};
function bbState() {
  const mm = ($('input[name=bbModeM]:checked') || {}).value || 'rata', me = ($('input[name=bbModeE]:checked') || {}).value || 'rata', rows = $$('#bbTbl tr[data-j]');
  const val = r => Math.max(0, parseNum($('.bbW', r).value) || 0), pg = rows.filter(r => r.dataset.j === 'MCQ'), es = rows.filter(r => r.dataset.j === 'ESSAY');
  const TE = A.bbTipeEs || { 1: es.map(r => r.dataset.id) }, totM = parseNum($('#bbM').value) || 0, totE = parseNum($('#bbE').value) || 0;
  const subM = !pg.length ? 0 : mm === 'per_soal' ? pg.reduce((a, r) => a + val(r), 0) : totM;
  const perTipe = Object.keys(TE).map(k => { const ids = TE[k], sE = !ids.length ? 0 : me === 'per_soal' ? es.filter(r => ids.includes(r.dataset.id)).reduce((a, r) => a + val(r), 0) : totE; return { k, n: ids.length, mcq: subM, essay: Math.round(sE * 100) / 100, total: Math.round((subM + sE) * 100) / 100 }; });
  const bad = perTipe.filter(t => Math.abs(t.total - 100) > 0.01);
  return { mm, me, rows, pg, es, val, TE, subM, totE, perTipe, bad, ok: !bad.length && rows.length > 0 };
}
function calcBb() {
  const S = bbState();
  $('#bbM').disabled = S.mm === 'per_soal'; $('#bbE').disabled = S.me === 'per_soal';
  if (S.mm !== 'per_soal') S.pg.forEach(r => { const i = $('.bbW', r); i.disabled = true; i.value = fmtNum(S.pg.length ? S.subM / S.pg.length : 0); }); else S.pg.forEach(r => $('.bbW', r).disabled = false);
  S.es.forEach(r => { const i = $('.bbW', r); if (S.me === 'per_soal') i.disabled = false; else { i.disabled = true; const k = (r.dataset.tipe || '').split(',')[0], n = (S.TE[k] || []).length || 1; i.value = fmtNum(S.totE / n); } });
  const essVals = [...new Set(S.perTipe.map(t => t.essay))];
  $('#bbSubM').textContent = fmtNum(S.subM); $('#bbSubE').textContent = essVals.length > 1 ? S.perTipe.map(t => 'T' + t.k + ' ' + fmtNum(t.essay)).join(' • ') : fmtNum(essVals[0] || 0);
  const show = S.bad[0] || S.perTipe[0] || { total: 0 };
  $('#bbTot').textContent = (S.bad.length && S.perTipe.length > 1 ? 'T' + show.k + ': ' : '') + fmtNum(show.total) + ' / 100';
  $('#bbTotNote').innerHTML = S.ok ? `<span style="color:var(--green);font-weight:700">✓ total ${S.perTipe.length > 1 ? 'tiap tipe ' : ''}sudah 100</span>` :
    S.bad.map(t => `<span style="color:var(--red);font-weight:700;display:block">${S.perTipe.length > 1 ? 'Tipe ' + t.k + ': ' : ''}${t.total > 100 ? 'kelebihan ' + fmtNum(t.total - 100) : 'kurang ' + fmtNum(100 - t.total)}</span>`).join('');
  $('#bbTotBox').style.borderColor = S.ok ? 'var(--green)' : 'var(--red)'; $('#bBbSave').disabled = !S.ok;
  $('#bbHint').textContent = S.ok ? '' : 'Simpan aktif setelah total ' + (S.perTipe.length > 1 ? 'setiap tipe ' : '') + 'tepat 100.';
}
$('#t-bb').addEventListener('input', calcBb); $('#t-bb').addEventListener('change', calcBb);
$('#bBbFix').onclick = () => {
  const S = bbState(), maxE = Math.max(0, ...S.perTipe.map(t => t.essay)), minE = Math.min(...S.perTipe.map(t => t.essay));
  if (!S.es.length && S.pg.length) { $('#bbM').value = '100'; $$('input[name=bbModeM]').forEach(r => r.checked = r.value === 'rata'); }
  else if (!S.pg.length && S.es.length && S.me === 'rata') $('#bbE').value = '100';
  else if (S.me === 'rata') $('#bbE').value = fmtNum(Math.max(0, 100 - S.subM));
  else {
    // poin essay per soal: skalakan essay tiap tipe agar PG + essay = 100 (hanya bila tiap essay dipakai satu tipe)
    const need = 100 - S.subM, multi = S.es.some(r => (r.dataset.tipe || '').split(',').length > 1);
    if (need <= 0) return toast('Kurangi poin pilihan ganda dulu: subtotal PG sudah ' + fmtNum(S.subM) + '.', 'warn');
    if (multi && S.perTipe.length > 1 && maxE !== minE) return toast('Ada soal essay yang dipakai di beberapa tipe dengan total berbeda — atur poinnya manual, atau pilih "Sama rata" untuk essay.', 'warn');
    S.perTipe.forEach(t => { const rs = S.es.filter(r => (S.TE[t.k] || []).includes(r.dataset.id)), sum = rs.reduce((a, r) => a + S.val(r), 0) || rs.length; let acc = 0;
      rs.forEach((r, i) => { const v = i === rs.length - 1 ? Math.round((need - acc) * 100) / 100 : Math.round(((S.val(r) || 1) / sum) * need * 100) / 100; acc += v; $('.bbW', r).value = fmtNum(v); }); });
  }
  calcBb();
};
$('#bBbSave').onclick = () => run('#bBbSave', async () => {
  const S = bbState(), per = {}; S.rows.forEach(r => per[r.dataset.j + ':' + r.dataset.id] = parseNum($('.bbW', r).value) || 0);
  const r = await call('adminSaveBobot', A.cid, { mode_mcq: S.mm, mode_essay: S.me, bobot_mcq: S.subM, bobot_essay: S.totE, per });
  A.pk = null; A.res = null; await reloadCourses();
  toast('Poin tersimpan: ' + (r.perTipe.length > 1 && new Set(r.perTipe.map(t => t.essay)).size > 1 ? r.perTipe.map(t => `tipe ${t.tipe} = PG ${fmtNum(t.mcq)} + essay ${fmtNum(t.essay)}`).join('; ') : `PG ${fmtNum(r.mcq)} + essay ${fmtNum(r.essay)} = ${fmtNum(r.total)}`) + '. Klik "Koreksi semua" bila nilai sudah pernah dihitung.');
});

/* ---------- 6. formula ---------- */
const FS_PROMPT = `Ketik ulang formula sheet pada lampiran ke format teks berikut, tanpa menambah atau mengurangi rumus:
# Judul bagian
Nama rumus :: $rumus dalam LaTeX$ :: keterangan simbol/satuan (boleh kosong)
- butir catatan atau konversi satuan
Aturan: satu rumus per baris; gunakan \\dfrac untuk pecahan; jangan beri penjelasan lain di luar format; balas hanya isi formula sheet-nya.`;
LOAD.fs = async () => { const f = await call('adminFormulaGet', A.cid); $('#fsMarkup').value = f.markup; A.fsImgs = f.gambar.slice(); $('#fsShow').checked = f.tampil; fsPrev(); };
function fsPrev() {
  $('#fsPrev').innerHTML = (renderFS($('#fsMarkup').value) + A.fsImgs.map(g => `<img class="fs-img" src="${g}" alt="">`).join('')) || '<p class="mut">Kosong.</p>';
  $('#fsImgs').innerHTML = A.fsImgs.map((g, i) => `<div style="position:relative"><img src="${g}" alt="" style="height:70px;border:1px solid var(--line);border-radius:6px"><button class="btn warn xs" data-rm="${i}" style="position:absolute;top:-6px;right:-6px">×</button></div>`).join('');
}
$('#fsMarkup').oninput = fsPrev;
$('#fsImgs').onclick = e => { const b = e.target.closest('[data-rm]'); if (b) { A.fsImgs.splice(Number(b.dataset.rm), 1); fsPrev(); } };
function shrink(canvasOrImg, w, h) { const max = 1500, s = Math.min(1, max / w), c = document.createElement('canvas'); c.width = Math.round(w * s); c.height = Math.round(h * s); const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.drawImage(canvasOrImg, 0, 0, c.width, c.height); return c.toDataURL('image/jpeg', 0.78); }
const loadImg = src => new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error('Gambar tidak bisa dibaca.')); i.src = src; });
function htmlToMarkup(html) {
  const d = document.createElement('div'); d.innerHTML = html; const out = [];
  d.querySelectorAll('h1,h2,h3,h4,p,li,tr').forEach(el => {
    if (el.tagName !== 'TR' && el.closest('tr')) return; const t = el.textContent.replace(/\s+/g, ' ').trim(); if (!t) return;
    if (/^H[12]$/.test(el.tagName)) out.push('# ' + t); else if (/^H[34]$/.test(el.tagName)) out.push('## ' + t); else if (el.tagName === 'LI') out.push('- ' + t);
    else if (el.tagName === 'TR') out.push([...el.children].map(c => c.textContent.replace(/\s+/g, ' ').trim()).filter(String).join(' :: ')); else out.push(t);
  });
  return out.join('\n');
}
$('#fsFile').onchange = () => run(null, async () => {
  const f = $('#fsFile').files[0]; if (!f) return; const n = f.name.toLowerCase(); let teks = '';
  $('#fsNote').textContent = 'Membaca ' + f.name + '…';
  if (/\.docx$/.test(n)) {
    const r = await mammoth.convertToHtml({ arrayBuffer: await readBuf(f) }); teks = htmlToMarkup(r.value);
    const d = document.createElement('div'); d.innerHTML = r.value;
    for (const im of [...d.querySelectorAll('img')].slice(0, 4 - A.fsImgs.length)) { try { const i = await loadImg(im.src); if (i.width > 80) A.fsImgs.push(shrink(i, i.width, i.height)); } catch (e) { } }
  } else if (/\.pdf$/.test(n)) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    const pdf = await pdfjsLib.getDocument({ data: await readBuf(f) }).promise, lines = [];
    for (let p = 1; p <= Math.min(pdf.numPages, 4); p++) {
      const pg = await pdf.getPage(p), tc = await pg.getTextContent(); let y = null, cur = '';
      tc.items.forEach(it => { const yy = Math.round(it.transform[5]); if (y !== null && Math.abs(yy - y) > 3) { lines.push(cur.trim()); cur = ''; } cur += it.str + (it.hasEOL ? ' ' : ''); y = yy; }); lines.push(cur.trim());
      if (A.fsImgs.length < 4) { const vp = pg.getViewport({ scale: 1.7 }), c = document.createElement('canvas'); c.width = vp.width; c.height = vp.height; await pg.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise; A.fsImgs.push(shrink(c, c.width, c.height)); }
    }
    teks = lines.filter(String).join('\n');
  } else if (/^image\//.test(f.type)) {
    if (A.fsImgs.length >= 4) throw new Error('Maksimal 4 gambar.');
    const i = await loadImg(URL.createObjectURL(f)); A.fsImgs.push(shrink(i, i.width, i.height));
  } else throw new Error('Format tidak didukung. Gunakan .docx, .pdf, atau gambar.');
  if (teks && (!$('#fsMarkup').value.trim() || confirm('Ganti isi editor dengan teks dari file ini?'))) $('#fsMarkup').value = teks;
  $('#fsFile').value = ''; $('#fsNote').textContent = 'File terbaca. Rapikan teks di editor (rumus dalam $…$), lalu Simpan. Rumus dari Word Equation/PDF sering tidak terbaca sebagai teks — gambarnya tetap ditampilkan.'; fsPrev();
});
$('#bFsSave').onclick = () => run('#bFsSave', async () => { await call('adminFormulaSave', A.cid, $('#fsMarkup').value, A.fsImgs, $('#fsShow').checked); await reloadCourses(); toast('Formula sheet tersimpan.'); });
$('#bFsPrompt').onclick = () => copyText(FS_PROMPT);

/* ---------- 7. monitor & 8. pelanggaran ---------- */
async function loadMon() { A.mon = await call('adminMonitor', A.cid); A.monOff = A.mon.now - Date.now(); renderMon(); renderViol(); }
LOAD.mon = async () => { await loadMon(); clearInterval(A.monT); let n = 0; A.monT = setInterval(() => { n++; tickMon(); if (n % 10 === 0 && $('#monAuto').checked && !A.monBusy) { A.monBusy = true; loadMon().catch(() => { }).finally(() => A.monBusy = false); } }, 1000); };
LOAD.vl = loadMon;
$('#bMonRef').onclick = () => run('#bMonRef', loadMon); $('#monQ').oninput = renderMon;
function tickMon() { const now = Date.now() + (A.monOff || 0); $$('#monTbl [data-dl]').forEach(td => td.textContent = fmtTime(Number(td.dataset.dl) - now)); }
function renderMon() {
  const m = A.mon, q = $('#monQ').value.trim().toLowerCase(), c = s => m.sesi.filter(x => x.status === s).length, off = m.sesi.filter(x => x.status === 'AKTIF' && (x.hb === null || x.hb > 80)).length;
  $('#monStats').innerHTML = [['Peserta terdaftar', m.roster], ['Sedang ujian', c('AKTIF')], ['Selesai', c('SELESAI') + c('WAKTU_HABIS')], ['Terkunci', c('TERKUNCI')], ['Offline', off], ['Belum mulai', Math.max(0, m.roster - m.sesi.length)]].map(x => `<div class="stat"><b>${x[1]}</b><span>${x[0]}</span></div>`).join('');
  $('#monAt').textContent = 'diperbarui ' + new Date().toLocaleTimeString('id-ID');
  const btn = (a, t, cls) => `<button class="btn ${cls || 'sec'} xs" data-a="${a}">${t}</button>`;
  $('#monTbl').innerHTML = '<tr><th>Mahasiswa</th><th>Tipe</th><th>Status</th><th class="r">Sisa waktu</th><th class="r">Terjawab</th><th>Sinyal</th><th class="r">Pelanggaran</th><th class="r">Di luar halaman</th><th>Aksi</th></tr>' + (m.sesi.filter(s => !q || (s.nama + ' ' + s.nim).toLowerCase().includes(q)).sort((a, b) => b.pelanggaran - a.pelanggaran || String(a.nim).localeCompare(String(b.nim))).map(s => {
    const live = s.status === 'AKTIF' || s.status === 'TERKUNCI', sig = !live ? '–' : s.hb === null ? '<span class="badge b-amber">belum ada</span>' : s.hb > 80 ? `<span class="badge b-red">offline ${Math.round(s.hb / 60)} mnt</span>` : '<span class="badge b-green">online</span>';
    const acts = (s.status === 'AKTIF' ? btn('plus5', '+5 menit') + btn('lock', 'Kunci', 'warn') : s.status === 'TERKUNCI' ? btn('plus5', '+5 menit') + btn('unlock', 'Buka kunci', 'ok') : s.status === 'WAKTU_HABIS' ? btn('plus5', '+5 menit') : btn('reopen', 'Buka lagi')) + (s.pelanggaran ? btn('rlog', 'Reset log') : '') + btn('reset', 'Reset sesi');
    return `<tr data-tk="${s.token}" data-nim="${esc(s.nim)}" data-nama="${esc(s.nama)}"><td><b>${esc(s.nama)}</b><div class="sm mut mono">${esc(s.nim)} • ${esc(s.kelas)}</div></td><td class="num">${s.tipe}</td><td>${stBadge(s.status)}</td><td class="r mono" ${live ? `data-dl="${s.dl}"` : ''}>${live ? '' : '–'}</td><td class="r num">${m.kertas ? '<span class="mut">kertas</span>' : s.dijawab + '/' + m.totalSoal}</td><td>${sig}</td><td class="r num">${s.pelanggaran ? `<span class="badge b-red">${s.pelanggaran}</span>` : '0'}</td><td class="r num">${s.diLuar ? `<span class="badge ${s.diLuar >= 60 ? 'b-red' : 'b-amber'}">${durTxt(s.diLuar)}</span>` : '<span class="mut">–</span>'}</td><td><div class="row" style="gap:4px">${acts}</div></td></tr>`;
  }).join('') || '<tr><td colspan="8" class="mut">Belum ada mahasiswa yang login.</td></tr>');
  tickMon();
}
$('#monTbl').onclick = e => {
  const b = e.target.closest('[data-a]'); if (!b) return; const tr = b.closest('tr'), a = b.dataset.a, nama = tr.dataset.nama;
  const ask = { reset: `Reset sesi ${nama}? Ujiannya diulang dari awal dengan waktu penuh; jawaban lama tidak dipakai.`, lock: `Kunci ujian ${nama}? Jawabannya tidak bisa diubah sampai dibuka.`, rlog: `Reset log pelanggaran ${nama}?`, reopen: `Buka lagi ujian ${nama} (minimal 5 menit)?` }[a];
  if (ask && !confirm(ask)) return;
  run(b, async () => { if (a === 'rlog') await call('adminResetLog', A.cid, tr.dataset.nim, false); else await call('adminAct', a, tr.dataset.tk); await loadMon(); toast(({ plus5: '+5 menit untuk ', lock: 'Dikunci: ', unlock: 'Dibuka: ', reset: 'Sesi direset: ', rlog: 'Log direset: ', reopen: 'Dibuka lagi: ' })[a] + nama); });
};
function renderViol() {
  if (!$('#vlTbl')) return; const v = A.mon.viol;
  $('#vlTbl').innerHTML = '<tr><th>Waktu keluar</th><th>Mahasiswa</th><th>Jenis</th><th class="r">Lama di luar</th><th>Detail</th><th>Dihitung</th><th class="r">Total</th><th></th></tr>' + (v.map(x => `<tr><td class="mono sm">${esc(x.waktu)}</td><td>${esc(x.nama)}<div class="sm mut mono">${esc(x.nim)}</div></td><td><b>${esc(x.jenis)}</b></td><td class="r num">${String(x.durasi_detik) !== '' && x.durasi_detik !== undefined ? `<b>${durTxt(x.durasi_detik)}</b><div class="sm mut">kembali ${esc(String(x.kembali).slice(11, 19))}</div>` : /PINDAH_TAB|KELUAR_JENDELA/.test(x.jenis) ? '<span class="sm mut">belum kembali</span>' : '<span class="mut">–</span>'}</td><td class="sm">${esc(x.detail)}</td><td>${x.dihitung === 'YA' ? '<span class="badge b-red">ya</span>' : `<span class="badge">${esc(String(x.dihitung).toLowerCase())}</span>`}</td><td class="r num">${esc(x.total_pelanggaran)}</td><td>${x.dihitung === 'YA' ? `<button class="btn sec xs" data-fg="${x.row}">Maafkan</button>` : ''}</td></tr>`).join('') || '<tr><td colspan="8" class="mut">Tidak ada pelanggaran tercatat.</td></tr>');
}
$('#vlTbl').onclick = e => { const b = e.target.closest('[data-fg]'); if (b) run(b, async () => { await call('adminForgive', Number(b.dataset.fg)); await loadMon(); }); };
$('#bRLog').onclick = () => { const nim = $('#rNim').value.trim(), del = $('#rDel').checked; if (confirm('Reset log pelanggaran ' + (nim ? 'NIM ' + nim : 'SEMUA mahasiswa di mata kuliah ini') + (del ? ' dan HAPUS PERMANEN barisnya' : '') + '?')) run('#bRLog', async () => { const r = await call('adminResetLog', A.cid, nim, del); $('#rNim').value = ''; $('#rDel').checked = false; await loadMon(); toast(`${r.baris} baris log direset, ${r.sesi} sesi diperbarui.`); }); };

/* ---------- 9. koreksi ---------- */
LOAD.gr = async () => { const c = cur(); $('#scCard').style.display = String(c.mode_jawaban).toUpperCase() === 'ONLINE' ? 'none' : ''; rosterMap().catch(() => { }); renderHd(); if (!A.res) A.res = await call('adminResults', A.cid); renderGr(); };
$('#bGrade').onclick = () => run('#bGrade', async () => { A.res = await call('adminGrade', A.cid); A.ch = null; renderGr(); toast(A.res.rows.length + ' mahasiswa dikoreksi.'); });
$('#grQ').oninput = renderGr; $('#grCek').onchange = renderGr;
const belum = r => A.res && A.res.kertas && !['PINDAI', 'TULISAN', 'IMPOR'].includes(r.asal);
function renderGr() {
  const q = $('#grQ').value.trim().toLowerCase(), cek = $('#grCek').checked, need = r => r.kategori.some((k, i) => (k === 'Partially Correct' || k === 'Far') && r.sumber[i] === 'LOKAL');
  const rows = A.res.rows.filter(r => (!q || (r.nama + ' ' + r.nim).toLowerCase().includes(q)) && (!cek || need(r)));
  const nb = Math.max(1, Math.ceil(A.res.rows.length / 10)); $('#aiBatch').innerHTML = Array.from({ length: nb }, (_, i) => `<option value="${i}">${i + 1} dari ${nb}</option>`).join('');
  $('#grTbl').innerHTML = '<tr><th>Mahasiswa</th><th>Tipe</th><th class="r">PG benar</th><th class="r">PG</th><th class="r">Essay</th><th class="r">Total</th><th>Essay</th><th>Catatan</th></tr>' + (rows.map(r => `<tr class="click" data-nim="${esc(r.nim)}"><td><b>${esc(r.nama)}</b><div class="sm mut mono">${esc(r.nim)} • ${esc(r.kelas)}</div></td><td class="num">${r.tipe}</td><td class="r num">${r.benar}/${r.nMcq}</td><td class="r num">${fmtNum(r.mcq)}</td><td class="r num">${fmtNum(r.essay)}</td><td class="r num"><b>${fmtNum(r.total)}</b></td>
    <td>${r.kategori.map((k, i) => katBadge(k) + (r.sumber[i] === 'MANUAL' ? '<sup title="diubah dosen">✎</sup>' : r.sumber[i] === 'AI' ? '<sup title="dinilai chat AI">ai</sup>' : '')).join(' ')}</td>
    <td>${belum(r) ? '<span class="badge">belum dipindai</span> ' : ''}${r.catatan ? `<span class="badge b-amber" title="${esc(r.catatan)}">cek lembar</span> ` : ''}${r.jebakan ? '<span class="badge b-red">jebakan AI</span> ' : ''}${r.pelanggaran ? `<span class="badge b-amber">${r.pelanggaran} pelanggaran</span>` : ''}${need(r) ? ' <span class="badge">lihat ulang</span>' : ''}</td></tr>`).join('') || `<tr><td colspan="8" class="mut">${A.res.rows.length ? 'Tidak ada yang cocok dengan filter.' : 'Belum ada nilai. Klik "Koreksi semua".'}</td></tr>`);
}
$('#grTbl').onclick = e => { const r = e.target.closest('tr[data-nim]'); if (r) run(null, async () => showDetail(await call('adminDetail', A.cid, r.dataset.nim))); };
function showDetail(d) {
  const src = s => s === 'MANUAL' ? '<span class="badge b-blue">diubah dosen</span>' : s === 'AI' ? '<span class="badge">chat AI</span>' : '';
  const scan = Object.keys(d.halaman || {}).length > 0;
  const body = (d.catatan ? `<div class="msg warn sm">${esc(d.catatan)}</div>` : '') + (scan ? `<div class="row" style="margin-bottom:8px;position:sticky;top:52px;background:#fff;z-index:1;padding:4px 0"><span class="sm mut">Foto jawaban:</span>${Object.keys(d.halaman).map(k => `<button class="btn sec xs" data-img="${esc(d.halaman[k])}">Halaman ${k}</button>`).join('')}<button class="btn sec xs" data-tutup="1">Sembunyikan</button></div><div id="dPage"></div>` : '') +
    (d.ocr ? `<details><summary>Teks OCR lengkap</summary><div class="cmp" style="max-height:300px;overflow:auto">${esc(d.ocr)}</div></details>` : '') + d.soal.map(s => {
    const head = `<div class="row"><span class="badge b-blue">No ${s.no}</span><span class="badge">${s.jenis === 'MCQ' ? 'PG' : 'Essay'}</span>${s.jenis === 'ESSAY' ? katBadge(s.kategori) : s.benar ? '<span class="badge b-green">benar</span>' : '<span class="badge b-red">salah</span>'}${src(s.sumber)}${s.trap ? `<span class="badge b-red">jebakan AI: ${esc(s.trap)}</span>` : ''}</div><div style="margin:6px 0">${rich(s.teks)}</div>`;
    const edit = `<div class="row" style="margin-top:8px"><label style="margin:0">Skor</label><input type="text" inputmode="decimal" class="dS" value="${fmtNum(s.skor)}" style="width:90px"><span class="mut">/ ${fmtNum(s.maks)}</span>` +
      (s.jenis === 'ESSAY' ? `<select class="dK" style="width:auto">${['Very Close', 'Close', 'Partially Correct', 'Far', 'Incorrect'].map(k => `<option ${k === s.kategori ? 'selected' : ''}>${k}</option>`).join('')}</select>` : '') +
      `<button class="btn xs" data-sv="${s.no}">Simpan nilai</button>${s.sumber === 'MANUAL' ? `<button class="btn sec xs" data-un="${s.no}">Kembalikan ke otomatis</button>` : ''}</div>`;
    if (s.jenis === 'MCQ') {
      const o = x => { const f = s.opsi.find(y => y.l === x); return f ? x + '. ' + f.t : '(kosong)'; };
      return `<div class="qc" data-no="${s.no}">${head}<div class="row sm"><span>Jawaban:</span><select class="dA" style="width:auto">${['', ...s.opsi.map(x => x.l)].map(l => `<option value="${l}" ${l === s.jawab ? 'selected' : ''}>${l || '(kosong)'}</option>`).join('')}</select><span>${rich(o(s.jawab))}</span><button class="btn sec xs" data-ans="${s.no}">Ubah jawaban</button><span class="mut">Kunci: <b>${rich(o(s.kunci))}</b></span></div><textarea class="dF" style="display:none">${esc(s.feedback)}</textarea>${edit}</div>`;
    }
    return `<div class="qc" data-no="${s.no}">${head}
      <label>${scan ? 'Teks hasil OCR (betulkan bila salah baca, lalu nilai ulang)' : 'Jawaban mahasiswa'}</label><textarea class="dT" style="min-height:90px">${esc(s.jawab)}</textarea><div class="row" style="margin-top:4px"><button class="btn sec xs" data-ans="${s.no}">Simpan teks &amp; nilai ulang otomatis</button></div>
      <details><summary>Acuan penilaian</summary><div class="sm"><b>Jawaban inti:</b> ${esc(s.inti)}</div><div class="sm"><b>Keyword:</b> ${esc(s.keyword)}</div><div class="sm"><b>Rubrik:</b> ${esc(s.rubrik) || '–'}</div></details>
      <label>Feedback</label><textarea class="dF" style="min-height:60px">${esc(s.feedback)}</textarea>${edit}</div>`;
  }).join('');
  openMod(`${d.nama} — ${d.nim} • tipe ${d.tipe} • total ${fmtNum(d.total)}`, body);
  $('#modB').onclick = e => {
    if (e.target.closest('[data-tutup]')) return $('#dPage').innerHTML = '';
    const im = e.target.closest('[data-img]');
    if (im) return run(im, async () => { $('#dPage').innerHTML = `<img class="esimg" style="max-height:75vh;object-fit:contain" src="${await call('adminImage', im.dataset.img)}" alt="Foto jawaban halaman">`; });
    const a = e.target.closest('[data-ans]');
    if (a) { const box = a.closest('.qc'), v = box.querySelector('.dA') ? box.querySelector('.dA').value : box.querySelector('.dT').value;
      return run(a, async () => { const nd = await call('adminSetAnswer', A.cid, d.nim, Number(box.dataset.no), v); A.res = await call('adminResults', A.cid); renderGr(); showDetail(nd); }); }
    const b = e.target.closest('[data-sv],[data-un]'); if (!b) return; const box = b.closest('.qc'), no = Number(box.dataset.no), un = b.dataset.un !== undefined;
    run(b, async () => { const sk = parseNum($('.dS', box).value); if (!un && (isNaN(sk) || sk < 0)) throw new Error('Skor tidak valid. Contoh: 7,5'); const nd = await call('adminSaveManual', A.cid, d.nim, no, sk, $('.dF', box).value, ($('.dK', box) || {}).value || '', un); A.res = await call('adminResults', A.cid); renderGr(); showDetail(nd); });
  };
}
$('#bGrImp').onclick = () => run('#bGrImp', async () => { const f = $('#grFile').files[0]; if (!f) throw new Error('Pilih file jawaban.'); const r = await call('adminImportAnswers', A.cid, await sheetRows(f, /jawaban|answer/i)); $('#grFile').value = ''; A.res = await call('adminGrade', A.cid); renderGr(); toast(`Jawaban ${r.mahasiswa} mahasiswa diimpor dan dikoreksi.`); });
$('#bAiCopy').onclick = () => run('#bAiCopy', async () => { const r = await call('adminEssayPrompt', A.cid, Number($('#aiBatch').value), 10, $('#aiOnly').checked); if (!r.prompt) throw new Error('Tidak ada jawaban essay pada kelompok ini.'); await copyText(r.prompt); });
$('#bAiApply').onclick = () => run('#bAiApply', async () => { const r = await call('adminApplyEssayAI', A.cid, $('#aiJson').value); $('#aiJson').value = ''; A.res = await call('adminResults', A.cid); renderGr(); toast(`${r.diterapkan} nilai essay dari ${r.mahasiswa} mahasiswa diterapkan.`); });

/* ---------- 10. kecurangan ---------- */
LOAD.ch = async () => { if (A.ch) renderCh(); else { $('#chList').innerHTML = '<div class="card mut">Klik <b>Analisis</b> setelah jawaban terkumpul.</div>'; $('#chTrap').innerHTML = ''; $('#chRisk').innerHTML = ''; } };
$('#bCheat').onclick = () => run('#bCheat', async () => { A.ch = await call('adminCheat', A.cid); A.chR = ''; renderCh(); });
$('#chBeda').onchange = renderCh;
$('#chRisk').onclick = e => { const b = e.target.closest('.chip'); if (b) { A.chR = b.dataset.r; renderCh(); } };
function renderCh() {
  const beda = $('#chBeda').checked, R = { High: 3, Medium: 2, Low: 1 }, G = {};
  A.ch.temuan.filter(x => !beda || x.a.tipe !== x.b.tipe).forEach(x => {
    const k = x.a.nim + '|' + x.b.nim, g = G[k] = G[k] || { a: x.a, b: x.b, items: [], risiko: 'Low', skor: 0 };
    g.items.push(x); if (R[x.risiko] > R[g.risiko]) g.risiko = x.risiko; g.skor += R[x.risiko] * 100 + x.persen;
  });
  // pasangan dengan ≥2 temuan Medium naik satu tingkat
  const all = Object.values(G).map(g => { if (g.risiko === 'Medium' && g.items.filter(x => x.risiko !== 'Low').length >= 2) g.risiko = 'High'; return g; }).sort((x, y) => R[y.risiko] - R[x.risiko] || y.skor - x.skor);
  const n = r => all.filter(x => x.risiko === r).length;
  $('#chRisk').innerHTML = [['', 'Semua pasangan', all.length], ['High', 'High', n('High')], ['Medium', 'Medium', n('Medium')], ['Low', 'Low', n('Low')]].map(x => `<button class="chip ${(A.chR || '') === x[0] ? 'on' : ''}" data-r="${x[0]}">${x[1]} (${x[2]})</button>`).join('');
  $('#chTrap').innerHTML = A.ch.jebakan.length ? `<div class="card" style="border-color:#e2a09b"><h3>Terkena jebakan AI (${A.ch.jebakan.length})</h3>${A.ch.jebakan.map(j => `<div class="sm"><b>${esc(j.nama)}</b> <span class="mono mut">${esc(j.nim)}</span> • tipe ${j.tipe} — ${esc(j.temuan)}</div>`).join('')}<div class="sm mut" style="margin-top:6px">Istilah ini hanya ada di teks jebakan soal; kemunculannya di jawaban berarti soal diproses oleh AI.</div></div>` : '';
  const who = p => `<b>${esc(p.nama)}</b> <span class="mono mut sm">${esc(p.nim)}</span> <span class="badge">tipe ${p.tipe}</span>`;
  const item = x => `<div style="border-top:1px dashed #ddd6c6;padding-top:8px;margin-top:8px"><div class="row">${riskBadge(x.risiko)}<span class="badge">${x.jenis === 'MCQ' ? 'Pilihan ganda' : 'Essay'}</span><span class="badge b-amber">${x.persen}% mirip</span><span class="sm mut">nomor soal (kiri/kanan): ${esc(x.soal)}</span></div><div class="sm" style="margin-top:4px">${esc(x.ringkas)}</div>
    ${x.bagian.length ? `<ul class="sm" style="margin:6px 0;padding-left:18px">${x.bagian.map(b => `<li>${esc(b)}</li>`).join('')}</ul>` : ''}
    ${x.teksA ? `<details style="margin:6px 0;padding:6px 10px"><summary class="sm">Bandingkan jawaban</summary><div class="two" style="margin-top:6px"><div class="cmp">${esc(x.teksA)}</div><div class="cmp">${esc(x.teksB)}</div></div></details>` : ''}</div>`;
  const list = all.filter(x => !A.chR || x.risiko === A.chR);
  $('#chList').innerHTML = list.slice(0, 200).map(g => `<div class="qc ${g.risiko === 'High' ? 'err' : ''}"><div class="row">${riskBadge(g.risiko)}<span>${who(g.a)} &nbsp;↔&nbsp; ${who(g.b)}</span><span class="sm mut right">${g.items.length} temuan</span></div>${g.items.map(item).join('')}</div>`).join('') ||
    `<div class="card mut">Tidak ada kemiripan mencurigakan di antara ${A.ch.mahasiswa} mahasiswa${A.chR || beda ? ' untuk filter ini' : ''}.</div>`;
}

/* ---------- 11. nilai akhir ---------- */
LOAD.fn = async () => { A.res = await call('adminResults', A.cid); renderFn(); };
$('#bFnRef').onclick = () => run('#bFnRef', LOAD.fn);
function ranked() { const r = A.res.rows.filter(x => !belum(x)).sort((a, b) => b.total - a.total || String(a.nim).localeCompare(String(b.nim))); r.forEach((x, i) => x.rank = i && r[i - 1].total === x.total ? r[i - 1].rank : i + 1); return r; }
function renderFn() {
  const r = ranked(), t = r.map(x => x.total), f = v => fmtNum(v);
  const nb = A.res.rows.filter(belum).length;
  $('#fnStats').innerHTML = [['Peserta dinilai', r.length + (nb ? ` <span class="sm mut" style="font-size:12px;font-weight:400">(+${nb} belum dipindai)</span>` : '')], ['Rata-rata', r.length ? f(t.reduce((a, b) => a + b, 0) / r.length) : '–'], ['Tertinggi', r.length ? f(Math.max(...t)) : '–'], ['Terendah', r.length ? f(Math.min(...t)) : '–'], ['Nilai maksimum', f(A.res.maks)]].map(x => `<div class="stat"><b>${x[1]}</b><span>${x[0]}</span></div>`).join('');
  $('#fnTbl').innerHTML = '<tr><th class="r">Peringkat</th><th>NIM</th><th>Nama</th><th>Kelas</th><th class="r">PG</th><th class="r">Essay</th><th class="r">Total</th></tr>' + (r.map(x => `<tr><td class="r num">${x.rank}</td><td class="mono">${esc(x.nim)}</td><td>${esc(x.nama)}</td><td>${esc(x.kelas)}</td><td class="r num">${f(x.mcq)}</td><td class="r num">${f(x.essay)}</td><td class="r num"><b>${f(x.total)}</b></td></tr>`).join('') || '<tr><td colspan="7" class="mut">Belum ada nilai. Buka tab Koreksi → "Koreksi semua".</td></tr>');
}
$('#bFnCsv').onclick = () => download('nilai_' + A.cid + '.csv', toCsv([['NIM', 'Nama', 'Score']].concat(ranked().sort((a, b) => String(a.nim).localeCompare(String(b.nim))).map(x => [String(x.nim), x.nama, x.total]))), 'text/csv');
$('#bFnXls').onclick = () => {
  const r = ranked(), wb = XLSX.utils.book_new(), byNim = r.slice().sort((a, b) => String(a.nim).localeCompare(String(b.nim)));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['NIM', 'Nama', 'Score']].concat(byNim.map(x => [String(x.nim), x.nama, x.total]))), 'Nilai');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Peringkat', 'NIM', 'Nama', 'Kelas', 'Tipe', 'PG benar', 'Skor PG', 'Skor Essay', 'Total', 'Pelanggaran', 'Jebakan AI']].concat(r.map(x => [x.rank, String(x.nim), x.nama, x.kelas, x.tipe, x.benar + '/' + x.nMcq, x.mcq, x.essay, x.total, x.pelanggaran, x.jebakan]))), 'Rinci');
  XLSX.writeFile(wb, 'nilai_' + A.cid + '.xlsx');
};

/* ---------- unggah jawaban tulisan tangan ---------- */
A.hd = [];
async function rosterMap() { if (!A.roster) A.roster = await call('adminRoster'); $('#rosterList').innerHTML = A.roster.map(r => `<option value="${esc(r.nim)}">${esc(r.nama)}</option>`).join(''); return Object.fromEntries(A.roster.map(r => [String(r.nim), r])); }
/** Gambar/halaman → JPEG abu-abu yang diratakan bayangannya (membantu OCR). */
function prepPage(src) {
  const s = Math.min(1, 2000 / Math.max(src.width, src.height)), W = Math.round(src.width * s), H = Math.round(src.height * s);
  const c = document.createElement('canvas'); c.width = W; c.height = H; const x = c.getContext('2d', { willReadFrequently: true });
  x.fillStyle = '#fff'; x.fillRect(0, 0, W, H); x.drawImage(src, 0, 0, W, H);
  const img = x.getImageData(0, 0, W, H), d = img.data, n = W * H, g = new Float32Array(n), I = new Float64Array((W + 1) * (H + 1)), r = Math.round(Math.max(W, H) / 40);
  for (let i = 0; i < n; i++) g[i] = d[i * 4] * .299 + d[i * 4 + 1] * .587 + d[i * 4 + 2] * .114;
  for (let y = 0; y < H; y++) { let row = 0; for (let xx = 0; xx < W; xx++) { row += g[y * W + xx]; I[(y + 1) * (W + 1) + xx + 1] = I[y * (W + 1) + xx + 1] + row; } }
  for (let y = 0; y < H; y++) for (let xx = 0; xx < W; xx++) {
    const x0 = Math.max(0, xx - r), x1 = Math.min(W, xx + r + 1), y0 = Math.max(0, y - r), y1 = Math.min(H, y + r + 1);
    const mean = (I[y1 * (W + 1) + x1] - I[y0 * (W + 1) + x1] - I[y1 * (W + 1) + x0] + I[y0 * (W + 1) + x0]) / ((x1 - x0) * (y1 - y0)), bg = Math.max(mean, 50) * 1.05, v = g[y * W + xx];
    const o = v >= bg ? 255 : 255 * Math.pow(v / bg, 1.7), k = (y * W + xx) * 4; d[k] = d[k + 1] = d[k + 2] = o;
  }
  x.putImageData(img, 0, 0); return c.toDataURL('image/jpeg', 0.85);
}
async function fileToPages(f) {
  if (/\.pdf$/i.test(f.name) || f.type === 'application/pdf') {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    const pdf = await pdfjsLib.getDocument({ data: await f.arrayBuffer() }).promise, out = [];
    for (let i = 1; i <= Math.min(pdf.numPages, 12); i++) {
      const pg = await pdf.getPage(i), v0 = pg.getViewport({ scale: 1 }), vp = pg.getViewport({ scale: Math.min(3, 2000 / Math.max(v0.width, v0.height)) });
      const c = document.createElement('canvas'); c.width = vp.width; c.height = vp.height; await pg.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise; out.push(prepPage(c));
    }
    return out;
  }
  const im = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error('Gambar tidak bisa dibuka: ' + f.name)); i.src = URL.createObjectURL(f); });
  return [prepPage(im)];
}
$('#hdFiles').onchange = () => run(null, async () => {
  const R = await rosterMap(), files = [...$('#hdFiles').files].sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true })), G = {};
  files.forEach(f => { const m = (f.name.match(/\d{6,15}/g) || []).find(x => R[x]), key = m || ('#' + f.name); (G[key] = G[key] || { nim: m || '', files: [], status: 'menunggu' }).files.push(f); });
  A.hd = A.hd.filter(h => h.status !== 'menunggu').concat(Object.values(G)); $('#hdFiles').value = ''; renderHd();
  $('#hdProg').textContent = Object.keys(G).length + ' mahasiswa siap dibaca. Klik "Baca & nilai".';
});
function renderHd() {
  const B = { menunggu: '<span class="badge">menunggu</span>', proses: '<span class="badge b-blue">membaca…</span>', ok: '<span class="badge b-green">tersimpan</span>', cek: '<span class="badge b-amber">tersimpan, perlu cek</span>', pilih: '<span class="badge b-red">pilih NIM/tipe</span>', gagal: '<span class="badge b-red">gagal</span>' };
  $('#hdWrap').style.display = A.hd.length ? '' : 'none';
  $('#hdTbl').innerHTML = '<tr><th>File</th><th>NIM</th><th>Nama</th><th>Tipe</th><th class="r">Nomor terbaca</th><th>Catatan</th><th>Status</th><th></th></tr>' + A.hd.map((h, i) => {
    const r = h.res || {}, st = R => (A.roster || []).find(x => String(x.nim) === String(R));
    const pick = h.status === 'pilih' ? `<div class="row" style="gap:4px;margin-top:4px"><input type="text" list="rosterList" placeholder="NIM" data-n="${i}" value="${esc(h.nim)}" style="width:130px;padding:5px"><select data-t="${i}" style="width:auto;padding:5px"><option value="">tipe…</option>${[1, 2, 3, 4, 5, 6].map(k => `<option ${Number(r.tipe) === k ? 'selected' : ''}>${k}</option>`).join('')}</select><button class="btn xs" data-as="${i}">Simpan</button></div>` : '';
    return `<tr><td class="sm">${h.files.map(f => esc(f.name)).join('<br>')}</td><td class="mono">${esc(r.nim || h.nim) || '<span class="mut">?</span>'}</td><td>${esc(r.nama || (st(h.nim) || {}).nama || '')}</td><td class="num">${r.tipe || '–'}</td><td class="r num">${r.total ? r.terbaca + '/' + r.total : '–'}</td>
      <td class="sm">${h.err ? `<span style="color:var(--red)">${esc(h.err)}</span>` : (r.catatan || []).length ? esc(r.catatan.slice(0, 3).join('; ')) + (r.catatan.length > 3 ? ` <span class="mut">+${r.catatan.length - 3}</span>` : '') : '<span class="mut">–</span>'}${pick}</td>
      <td>${B[h.status] || ''}</td><td><div class="row" style="gap:4px">${r.total ? `<button class="btn sec xs" data-lihat="${esc(r.nim)}">Lihat</button>` : ''}<button class="btn sec xs" data-hrm="${i}" aria-label="Buang">×</button></div></td></tr>`;
  }).join('');
}
function hdApply(h, r) {
  h.res = r; h.err = '';
  if (r.perluNim || r.perluTipe) { h.status = 'pilih'; h.err = r.perluNim ? 'NIM tidak terbaca dari tulisan — pilih NIM.' : 'Mahasiswa tidak login & tipe tidak tertulis — pilih tipe.'; if (r.nim) h.nim = r.nim; }
  else { h.nim = r.nim; h.status = r.catatan.length ? 'cek' : 'ok'; }
}
$('#bHdRead').onclick = () => run('#bHdRead', async () => {
  const todo = A.hd.filter(h => h.status === 'menunggu' || h.status === 'gagal');
  if (!todo.length) throw new Error('Pilih file jawaban terlebih dahulu.');
  let done = 0; const q = todo.slice();
  const worker = async () => { while (q.length) { const h = q.shift(); h.status = 'proses'; renderHd();
    try { const pages = []; for (const f of h.files) pages.push(...await fileToPages(f)); hdApply(h, await call('adminHandSubmit', A.cid, { nim: h.nim, pages })); }
    catch (e) { h.status = 'gagal'; h.err = e.message; }
    done++; $('#hdProg').textContent = `${done}/${todo.length} mahasiswa dibaca…`; renderHd(); } };
  await Promise.all([worker(), worker(), worker()]);
  A.res = await call('adminGrade', A.cid); A.ch = null; renderGr();
  const n = s => A.hd.filter(h => h.status === s).length;
  $('#hdProg').textContent = ''; toast(`${n('ok') + n('cek')} jawaban tersimpan & dinilai` + (n('cek') ? ` (${n('cek')} perlu dicek)` : '') + (n('pilih') ? `, ${n('pilih')} perlu dipilih NIM/tipe` : '') + (n('gagal') ? `, ${n('gagal')} gagal` : '') + '.', n('gagal') || n('pilih') ? 'warn' : 'ok');
});
$('#hdTbl').onclick = e => {
  const as = e.target.closest('[data-as]'), lh = e.target.closest('[data-lihat]'), rm = e.target.closest('[data-hrm]');
  if (rm) { A.hd.splice(Number(rm.dataset.hrm), 1); return renderHd(); }
  if (lh) return run(lh, async () => showDetail(await call('adminDetail', A.cid, lh.dataset.lihat)));
  if (as) { const i = Number(as.dataset.as), h = A.hd[i], nim = $(`#hdTbl [data-n="${i}"]`).value.trim(), tipe = Number($(`#hdTbl [data-t="${i}"]`).value) || '';
    if (!nim) return toast('Isi NIM.', 'err');
    run(as, async () => { hdApply(h, await call('adminHandAssign', A.cid, { nim, tipe, text: h.res.text, pageIds: h.res.pageIds, notes: h.res.notes || [] })); renderHd(); A.res = await call('adminGrade', A.cid); renderGr(); }); }
};

/* ---------- PDF (jsPDF + font Unicode DejaVu) ---------- */
const PDF_LIB = ['https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js'];
const FONT_URL = 'https://cdn.jsdelivr.net/npm/dejavu-fonts-ttf@2.37.3/ttf/';
const loadScript = src => new Promise((res, rej) => { if ([...document.scripts].some(x => x.src === src)) return res(); const x = document.createElement('script'); x.src = src; x.onload = res; x.onerror = () => rej(new Error('Gagal memuat pustaka PDF. Periksa koneksi internet.')); document.head.appendChild(x); });
const NAVY = [31, 42, 92];
async function pdfDoc(title, sub) {
  for (const u of PDF_LIB) await loadScript(u);
  if (A.fonts === undefined) {
    const get = async n => { const r = await fetch(FONT_URL + n); if (!r.ok) throw new Error('font'); const b = new Uint8Array(await r.arrayBuffer()); let s = ''; for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000)); return btoa(s); };
    try { A.fonts = { r: await get('DejaVuSans.ttf'), b: await get('DejaVuSans-Bold.ttf') }; } catch (e) { A.fonts = null; }
  }
  const doc = new jspdf.jsPDF({ unit: 'mm', format: 'a4' });
  if (A.fonts) { doc.addFileToVFS('DV.ttf', A.fonts.r); doc.addFont('DV.ttf', 'DV', 'normal'); doc.addFileToVFS('DVB.ttf', A.fonts.b); doc.addFont('DVB.ttf', 'DV', 'bold'); doc.FN = 'DV'; } else doc.FN = 'helvetica';
  doc.setFont(doc.FN, 'bold'); doc.setFontSize(15); doc.setTextColor(...NAVY); doc.text(pdfSafe(doc, title), 14, 16);
  doc.setFont(doc.FN, 'normal'); doc.setFontSize(9.5); doc.setTextColor(90);
  const sl = doc.splitTextToSize(pdfSafe(doc, sub || ''), 182).slice(0, 2); sl.forEach((l, i) => doc.text(l, 14, 22.5 + i * 4.4));
  const ly = 25.5 + (sl.length - 1) * 4.4; doc.setDrawColor(201, 130, 10); doc.setLineWidth(0.6); doc.line(14, ly, 196, ly); doc.setTextColor(20); doc.Y = ly + 5.5; return doc;
}
/** Tanpa font Unicode (gagal unduh), ganti karakter di luar Latin-1 agar tidak jadi kotak. */
function pdfSafe(doc, s) { s = String(s === undefined || s === null ? '' : s); if (doc.FN === 'DV') return s; return s.replace(/[π]/g, 'pi').replace(/[σ]/g, 'sigma').replace(/[τ]/g, 'tau').replace(/[φ]/g, 'phi').replace(/[√]/g, 'sqrt').replace(/[≤]/g, '<=').replace(/[≥]/g, '>=').replace(/[≈]/g, '~').replace(/[⁰¹²³⁴⁵⁶⁷⁸⁹]/g, c => '^' + '⁰¹²³⁴⁵⁶⁷⁸⁹'.indexOf(c)).replace(/[₀₁₂₃₄₅₆₇₈₉]/g, c => '₀₁₂₃₄₅₆₇₈₉'.indexOf(c)).replace(/[^\x00-\xff]/g, '?'); }
function pdfFooter(doc, label) {
  const n = doc.getNumberOfPages(), when = new Date().toLocaleString('id-ID', { dateStyle: 'medium', timeStyle: 'short' });
  for (let i = 1; i <= n; i++) { doc.setPage(i); doc.setFont(doc.FN, 'normal'); doc.setFontSize(8); doc.setTextColor(120); doc.text(doc.splitTextToSize(pdfSafe(doc, label), 100)[0], 14, 290); doc.text(`Halaman ${i} dari ${n} • dibuat ${when}`, 196, 290, { align: 'right' }); }
}
/** Rumus LaTeX → teks biasa yang terbaca di PDF. */
function texPlain(s) {
  const G = { alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', Delta: 'Δ', epsilon: 'ε', varepsilon: 'ε', eta: 'η', theta: 'θ', lambda: 'λ', mu: 'μ', nu: 'ν', pi: 'π', rho: 'ρ', sigma: 'σ', Sigma: 'Σ', tau: 'τ', phi: 'φ', varphi: 'φ', psi: 'ψ', omega: 'ω', Omega: 'Ω', times: '×', cdot: '·', le: '≤', leq: '≤', ge: '≥', geq: '≥', approx: '≈', pm: '±', infty: '∞', circ: '°', degree: '°', rightarrow: '→', to: '→', Rightarrow: '⇒', sum: 'Σ', int: '∫', partial: '∂', ln: 'ln', log: 'log', sin: 'sin', cos: 'cos', tan: 'tan', exp: 'exp' };
  const SUP = { 0: '⁰', 1: '¹', 2: '²', 3: '³', 4: '⁴', 5: '⁵', 6: '⁶', 7: '⁷', 8: '⁸', 9: '⁹', '+': '⁺', '-': '⁻', n: 'ⁿ', i: 'ⁱ', o: '°' }, SUB = { 0: '₀', 1: '₁', 2: '₂', 3: '₃', 4: '₄', 5: '₅', 6: '₆', 7: '₇', 8: '₈', 9: '₉', a: 'ₐ', e: 'ₑ', o: 'ₒ', x: 'ₓ', p: 'ₚ', s: 'ₛ', t: 'ₜ', i: 'ᵢ', r: 'ᵣ', n: 'ₙ', m: 'ₘ', k: 'ₖ', h: 'ₕ', l: 'ₗ' };
  return String(s === undefined || s === null ? '' : s).replace(/\$\$?([^$]+)\$\$?/g, (m, x) => {
    let t = x; for (let i = 0; i < 6; i++) t = t.replace(/\\[dt]?frac\{([^{}]*)\}\{([^{}]*)\}/g, (q, a, b) => (a.length > 1 ? '(' + a + ')' : a) + '/' + (b.length > 1 ? '(' + b + ')' : b));
    t = t.replace(/\\sqrt\{([^{}]*)\}/g, '√($1)').replace(/\\(?:text|mathrm|mathit|operatorname)\{([^{}]*)\}/g, '$1').replace(/\\(left|right|,|;|!|quad|qquad|displaystyle)/g, ' ');
    t = t.replace(/\\([A-Za-z]+)/g, (q, n) => G[n] !== undefined ? G[n] : n).replace(/\{,\}/g, ',');
    t = t.replace(/\^\{([^{}]*)\}|\^(\S)/g, (q, a, b) => { const v = a !== undefined ? a : b; return [...v].every(c => SUP[c]) ? [...v].map(c => SUP[c]).join('') : '^(' + v + ')'; });
    t = t.replace(/_\{([^{}]*)\}|_(\S)/g, (q, a, b) => { const v = a !== undefined ? a : b; return [...v].every(c => SUB[c]) ? [...v].map(c => SUB[c]).join('') : '_' + v; });
    return t.replace(/[{}]/g, '').replace(/\s+/g, ' ').trim();
  });
}
/** Tulis paragraf terbungkus dengan pindah halaman otomatis. */
function para(doc, text, o) {
  o = Object.assign({ x: 14, w: 182, size: 10, bold: false, color: [20, 20, 20], gap: 1.2, lh: 0.42 }, o || {});
  doc.setFont(doc.FN, o.bold ? 'bold' : 'normal'); doc.setFontSize(o.size); doc.setTextColor(...o.color);
  const lines = doc.splitTextToSize(pdfSafe(doc, text), o.w), h = o.size * o.lh;
  lines.forEach(l => { if (doc.Y + h > 282) { doc.addPage(); doc.Y = 16; doc.setFont(doc.FN, o.bold ? 'bold' : 'normal'); doc.setFontSize(o.size); doc.setTextColor(...o.color); } doc.text(l, o.x, doc.Y); doc.Y += h; });
  doc.Y += o.gap;
}
const tbl = (doc, o) => doc.autoTable(Object.assign({ startY: doc.Y, styles: { font: doc.FN, fontSize: 9.5, cellPadding: 1.6, textColor: 20 }, headStyles: { fillColor: NAVY, textColor: 255, fontStyle: 'bold' }, alternateRowStyles: { fillColor: [246, 244, 238] }, margin: { left: 14, right: 14, bottom: 14 } }, o));
const studentUrl = () => location.href.replace(/admin\.html.*$/, '').replace(/[?#].*$/, '');
/** Daftar mahasiswa & tipe soal. */
async function pdfTipe(btn) {
  await run(btn, async () => {
    const L = await call('adminTipeList', A.cid); if (!L.list.length) throw new Error('Belum ada peserta untuk mata kuliah ini.');
    const doc = await pdfDoc('Daftar Mahasiswa & Tipe Soal', L.course);
    para(doc, 'Jumlah per tipe: ' + Object.keys(L.jumlah).sort().map(k => `Tipe ${k} = ${L.jumlah[k]} orang`).join('   •   ') + `   (total ${L.list.length})`, { size: 9.5, color: [80, 80, 80], gap: 2 });
    tbl(doc, { head: [['No', 'NIM', 'Nama', 'Kelas', 'Tipe']], body: L.list.map((r, i) => [i + 1, r.nim, pdfSafe(doc, r.nama), r.kelas, r.tipe]), columnStyles: { 0: { halign: 'right', cellWidth: 11 }, 1: { cellWidth: 30 }, 3: { cellWidth: 18, halign: 'center' }, 4: { cellWidth: 16, halign: 'center', fontStyle: 'bold' } } });
    pdfFooter(doc, 'Daftar tipe soal — ' + L.course); doc.save(`Daftar_Tipe_${A.cid}.pdf`);
    toast('PDF daftar mahasiswa & tipe soal diunduh. Tipe tiap mahasiswa sekarang tetap (dipakai saat login).');
  });
}
$('#bPdfTipe').onclick = () => pdfTipe('#bPdfTipe'); $('#bPdfTipe2').onclick = () => pdfTipe('#bPdfTipe2');
$('#bTipeReset').onclick = () => { if (confirm('Bagi ulang tipe untuk mahasiswa yang BELUM login? (Yang sudah login tetap.) PDF daftar tipe lama tidak berlaku lagi.')) run('#bTipeReset', async () => { const r = await call('adminResetTipe', A.cid); await call('adminTipeList', A.cid); toast(r.dihapus + ' mahasiswa dibagi ulang. Unduh lagi PDF daftar tipe.'); }); };
/** Soal & kunci per tipe. */
$('#bPdfKunci').onclick = () => run('#bPdfKunci', async () => {
  const pk = await call('adminGetPaket', A.cid), c = cur(); if (!pk.paket.length) throw new Error('Buat tipe ujian dulu (tab Bank & Generator).');
  const T = [...new Set(pk.paket.map(p => Number(p.tipe)))].sort((a, b) => a - b), E = Object.fromEntries(pk.essay.map(q => [String(q.id), q]));
  const at = (k, no) => pk.paket.find(p => Number(p.tipe) === k && Number(p.no) === no), nos = pk.paket.filter(p => Number(p.tipe) === T[0]).map(p => Number(p.no)).sort((a, b) => a - b);
  const tp = pk.total || {};
  const doc = await pdfDoc('Soal & Kunci Jawaban per Tipe', `${c.nama} • ${T.length} tipe • ${fmtNum(c.durasi_menit)} menit • total poin ${fmtNum(tp.total)} (PG ${fmtNum(tp.mcq)} + essay ${fmtNum(tp.essay)})`);
  para(doc, 'Ringkasan kunci per tipe', { bold: true, size: 11.5, color: NAVY, gap: 1 });
  const maxNo = Math.max(...pk.paket.map(p => Number(p.no))), allNos = Array.from({ length: maxNo }, (_, i) => i + 1);
  tbl(doc, { head: [['No', ...T.map(k => 'Tipe ' + k), 'Poin']], body: allNos.map(no => { const ps = T.map(k => at(k, no)), pv = [...new Set(ps.map((p, i) => p ? fmtNum(pk.bobot[T[i] + ':' + no]) : null).filter(Boolean))];
      return [no, ...ps.map((p, i) => !p ? '–' : p.jenis === 'MCQ' ? (p.kunci || '?') : 'Essay ' + p.bank_id + (pv.length > 1 ? ' (' + fmtNum(pk.bobot[T[i] + ':' + no]) + ')' : '')), pv.length > 1 ? 'beda' : pv[0]]; }),
    styles: { font: doc.FN, fontSize: 9, halign: 'center', cellPadding: 1.2 }, columnStyles: { 0: { cellWidth: 12 } }, didParseCell: d => { if (d.section === 'body' && d.column.index > 0 && d.column.index <= T.length && /^[A-E]$/.test(d.cell.raw)) d.cell.styles.fontStyle = 'bold'; } });
  doc.Y = doc.lastAutoTable.finalY + 4;
  para(doc, 'Kunci "?" = belum ada kunci di bank soal. "–" = nomor ini tidak ada di tipe tersebut (jumlah essay berbeda per tipe). Poin melekat pada soal; bila essay berbeda per tipe, poin essay tertulis di dalam sel.', { size: 8.5, color: [110, 110, 110] });
  T.forEach(k => {
    doc.addPage(); doc.Y = 18; para(doc, `TIPE ${k}`, { bold: true, size: 14, color: NAVY, gap: 0.5 }); para(doc, c.nama, { size: 9.5, color: [90, 90, 90], gap: 3 });
    pk.paket.filter(p => Number(p.tipe) === k).sort((a, b) => a.no - b.no).forEach(p => {
      if (doc.Y > 255) { doc.addPage(); doc.Y = 16; }
      para(doc, `${p.no}.  ${p.jenis === 'MCQ' ? 'Pilihan ganda' : 'Essay'} • ${fmtNum(pk.bobot[k + ':' + p.no])} poin • bank ${p.bank_id}`, { bold: true, size: 9, color: [110, 110, 110], gap: 0.4 });
      para(doc, texPlain(p.soal), { size: 10.5, gap: 1 });
      if (p.jenis === 'MCQ') {
        let o = []; try { o = JSON.parse(p.opsi_json || '[]'); } catch (e) { }
        o.forEach(x => { const key = x.l === p.kunci; para(doc, `${x.l}.  ${texPlain(x.t)}${key ? '   ✓ KUNCI' : ''}`, { x: 20, w: 176, size: 10, bold: key, color: key ? [28, 122, 74] : [30, 30, 30], gap: 0.3 }); });
        if (!p.kunci) para(doc, 'Belum ada kunci jawaban.', { x: 20, size: 9, color: [179, 38, 30] });
      } else {
        const b = E[String(p.bank_id)] || {}, row = (lab, v) => para(doc, `${lab}: ${texPlain(v) || '–'}`, { x: 20, w: 176, size: 9.5, color: [30, 30, 30], gap: 0.4 });
        row('Jawaban inti', b.jawaban_inti); row('Keyword', b.keyword); if (String(b.rubrik || '').trim()) row('Rubrik', b.rubrik);
        if (p.jebakan_kata) para(doc, `Jebakan AI: istilah "${p.jebakan_kata}"`, { x: 20, size: 8.5, color: [140, 90, 0] });
      }
      doc.Y += 2.5;
    });
  });
  pdfFooter(doc, 'RAHASIA — kunci jawaban ' + c.nama); doc.save(`Soal_Kunci_${A.cid}.pdf`); toast('PDF soal & kunci per tipe diunduh.');
});
/** Password login: tabel + kartu potong. */
$('#bPsPdf').onclick = () => run('#bPsPdf', async () => {
  const c = cur(); let L = null; try { L = await call('adminTipeList', A.cid); } catch (e) { L = null; }
  const ros = L ? L.list : (A.roster || await call('adminRoster')).filter(r => !String(r.course || '').trim() || String(r.course).split(/[,;\s]+/).includes(String(A.cid))).map(r => ({ nim: String(r.nim), nama: r.nama, kelas: String(r.kelas || ''), password: r.password }));
  if (!ros.length) throw new Error('Belum ada peserta. Impor daftar mahasiswa dulu.');
  if (ros.some(r => !r.password)) throw new Error('Ada peserta tanpa password. Klik "Buat password yang kosong" dulu.');
  const url = studentUrl(), withT = !!L, doc = await pdfDoc('Password Login Ujian', `${c.nama} • ${ros.length} peserta`);
  para(doc, `Alamat ujian: ${url}`, { size: 9.5, color: [60, 60, 60], gap: 0.5 }); para(doc, 'Login: pilih mata kuliah, isi NIM dan password 4 karakter (huruf besar/kecil sama saja).', { size: 9, color: [90, 90, 90], gap: 2 });
  tbl(doc, { head: [['No', 'NIM', 'Nama', 'Kelas', ...(withT ? ['Tipe'] : []), 'Password']], body: ros.map((r, i) => [i + 1, r.nim, pdfSafe(doc, r.nama), r.kelas, ...(withT ? [r.tipe] : []), r.password]),
    columnStyles: { 0: { halign: 'right', cellWidth: 11 }, 1: { cellWidth: 28 }, [withT ? 5 : 4]: { font: 'courier', fontStyle: 'bold', fontSize: 12, halign: 'center', cellWidth: 26 }, 3: { halign: 'center', cellWidth: 16 }, ...(withT ? { 4: { halign: 'center', cellWidth: 14 } } : {}) } });
  // kartu potong: 3 kolom × 7 baris
  const W = 60, Hc = 37, X0 = 15, Y0 = 14; doc.addPage();
  ros.forEach((r, i) => {
    const k = i % 21; if (i && k === 0) doc.addPage(); const x = X0 + (k % 3) * (W + 0), y = Y0 + Math.floor(k / 3) * (Hc + 2.5);
    doc.setLineDashPattern([1.2, 1], 0); doc.setDrawColor(150); doc.setLineWidth(0.2); doc.rect(x, y, W, Hc); doc.setLineDashPattern([], 0);
    doc.setFont(doc.FN, 'bold'); doc.setFontSize(9); doc.setTextColor(20); doc.text(doc.splitTextToSize(pdfSafe(doc, r.nama), W - 6)[0], x + 3, y + 5.5);
    doc.setFont(doc.FN, 'normal'); doc.setFontSize(8.5); doc.setTextColor(80); doc.text(doc.splitTextToSize(`NIM ${r.nim}${r.kelas ? ' • Kelas ' + r.kelas : ''}`, W - 6)[0], x + 3, y + 10.5);
    doc.setFontSize(7); doc.text(doc.splitTextToSize(pdfSafe(doc, c.nama), W - 6)[0], x + 3, y + 14.5);
    doc.setFontSize(8); doc.setTextColor(60); doc.text('Password:', x + 3, y + 22); doc.setFont('courier', 'bold'); doc.setFontSize(20); doc.setTextColor(...NAVY); doc.text(String(r.password), x + 22, y + 23.2);
    if (withT) { doc.setFont(doc.FN, 'bold'); doc.setFontSize(9); doc.setTextColor(140, 90, 0); doc.text('Tipe soal: ' + r.tipe, x + 3, y + 29.5); }
    doc.setFont(doc.FN, 'normal'); doc.setFontSize(6.5); doc.setTextColor(110); doc.text(doc.splitTextToSize(url, W - 6)[0], x + 3, y + 34);
  });
  pdfFooter(doc, 'Password login — ' + c.nama); doc.save(`Password_${A.cid}.pdf`); toast('PDF password diunduh (tabel + kartu potong).');
});

/* ---------- impor daftar mahasiswa yang tahan format SIAKAD ---------- */
async function rosterRows(file) {
  const n = file.name.toLowerCase();
  const wb = XLSX.read(await readBuf(file), { type: 'array', raw: false }), ws = wb.Sheets[wb.SheetNames.find(x => /peserta|roster|mahasiswa|daftar/i.test(x)) || wb.SheetNames[0]];
  const R = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false }).map(r => r.map(c => String(c).trim()));
  const hi = R.slice(0, 20).findIndex(r => r.some(c => /^(nim|npm|no\.?\s*induk|nomor\s*induk(\s*mahasiswa)?)$/i.test(c)));
  if (hi >= 0) { const Hd = R[hi]; return R.slice(hi + 1).filter(r => r.some(Boolean)).map(r => Object.fromEntries(Hd.map((h, i) => [h || 'kolom' + i, r[i] || '']))); }
  return R.filter(r => r.some(c => /^\d{6,15}$/.test(c))).map(r => {
    const i = r.findIndex(c => /^\d{6,15}$/.test(c)), nama = r.find((c, j) => j !== i && /[A-Za-z]{2,}/.test(c) && c.length > 3) || '';
    const kelas = r.find((c, j) => j !== i && c !== nama && /^[A-Za-z0-9]{1,3}$/.test(c) && !/^\d{1,3}$/.test(c)) || '';
    return { nim: r[i], nama, kelas };
  });
}
$('#bPsRegen').onclick = () => { if (confirm('Buat ulang SEMUA password? Password lama (dan PDF/kartu yang sudah dibagikan) tidak berlaku lagi.')) run('#bPsRegen', async () => { const r = await call('adminGenPasswords', true); await LOAD.ps(); toast(r.dibuat + ' password baru dibuat. Unduh lagi PDF password.'); }); };

/* ---------- mulai ---------- */
(async () => { if (A.t) { try { return await boot(); } catch (e) { if (A.t) return showPin(e.message); return; } } showPin(); })();
