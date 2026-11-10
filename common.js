/* Fungsi bersama halaman mahasiswa & dosen */
const $ = (s, r) => (r || document).querySelector(s);
const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));
const esc = s => String(s === undefined || s === null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

async function api(fn, ...args) {
  if (!window.CONFIG || /TEMPEL_ID/.test(CONFIG.API_URL)) throw new Error('config.js belum diisi URL Apps Script.');
  let r;
  try { r = await fetch(CONFIG.API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ fn, args }), redirect: 'follow' }); }
  catch (e) { const x = new Error('Koneksi ke server gagal. Periksa internet Anda.'); x.net = true; throw x; }
  let j; try { j = await r.json(); } catch (e) { const x = new Error('Server tidak merespons dengan benar (cek URL di config.js dan akses deployment "Anyone").'); x.net = true; throw x; }
  if (!j.ok) throw new Error(j.error);
  return j.data;
}

/** Angka format Indonesia: titik = ribuan, koma = desimal (1.250,5). */
function fmtNum(v, d) { const n = Number(v); if (v === '' || v === null || v === undefined || isNaN(n)) return v === undefined || v === null ? '' : String(v); return n.toLocaleString('id-ID', { maximumFractionDigits: d === undefined ? 2 : d }); }
/** "1.250,5" / "12,5" / "12.5" / "1.250" → angka. */
function parseNum(s) {
  s = String(s === undefined || s === null ? '' : s).trim().replace(/\s/g, ''); if (!s) return NaN;
  if (s.includes(',')) return Number(s.replace(/\./g, '').replace(',', '.'));
  if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) return Number(s.replace(/\./g, ''));
  return Number(s);
}
/** Teks soal → HTML: aman dari HTML, $rumus$ dirender KaTeX, baris baru dipertahankan. */
function rich(s) {
  const parts = String(s === undefined || s === null ? '' : s).split(/(\$\$[^$]+\$\$|\$[^$\n]+\$)/);
  return parts.map(p => {
    const m = /^(\$\$?)([^$]+)\1$/.exec(p);
    if (!m) return esc(p).replace(/\n/g, '<br>');
    if (window.katex) { try { return katex.renderToString(m[2].replace(/(\d),(?=\d)/g, '$1{,}'), { throwOnError: false, displayMode: m[1] === '$$' }); } catch (e) { } }
    return '<code class="tex">' + esc(m[2]) + '</code>';
  }).join('');
}

/**
 * Markup formula sheet → HTML.
 *   # Judul bagian        ## sub-judul        - butir
 *   Nama :: $rumus$ :: keterangan      (keterangan opsional)
 */
function renderFS(markup) {
  const out = []; let open = false, rows = [], list = [];
  const flush = () => {
    if (rows.length) { out.push('<table class="fs-t">' + rows.map(r => `<tr><th>${rich(r[0])}</th><td class="fs-f">${rich(r[1] || '')}</td><td class="fs-n">${rich(r[2] || '')}</td></tr>`).join('') + '</table>'); rows = []; }
    if (list.length) { out.push('<ul>' + list.map(x => '<li>' + rich(x) + '</li>').join('') + '</ul>'); list = []; }
  };
  const sec = t => { flush(); if (open) out.push('</section>'); out.push('<section class="fs-sec">' + (t ? '<h3>' + rich(t) + '</h3>' : '')); open = true; };
  String(markup || '').split(/\r?\n/).forEach(line => {
    const l = line.trim(); if (!l) { flush(); return; }
    if (/^#\s+/.test(l)) return sec(l.replace(/^#\s+/, ''));
    if (!open) sec('');
    if (/^##\s+/.test(l)) { flush(); out.push('<h4>' + rich(l.replace(/^##\s+/, '')) + '</h4>'); }
    else if (l.includes('::')) { if (list.length) flush(); rows.push(l.split('::').map(x => x.trim())); }
    else if (/^[-•*]\s+/.test(l)) { if (rows.length) flush(); list.push(l.replace(/^[-•*]\s+/, '')); }
    else { flush(); out.push('<p>' + rich(l) + '</p>'); }
  });
  flush(); if (open) out.push('</section>');
  return '<div class="fs">' + out.join('') + '</div>';
}
function fmtTime(ms) { ms = Math.max(0, ms); const s = Math.floor(ms / 1000), h = Math.floor(s / 3600), m = Math.floor(s % 3600 / 60), d = s % 60; return (h ? h + ':' : '') + String(m).padStart(2, '0') + ':' + String(d).padStart(2, '0'); }
function deviceId() { let d = localStorage.getItem('ujian_dev'); if (!d) { d = 'd' + Math.random().toString(36).slice(2, 12) + Date.now().toString(36); localStorage.setItem('ujian_dev', d); } return d; }
