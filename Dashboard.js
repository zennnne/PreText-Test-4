import { prepareWithSegments, layoutWithLines } from 'https://esm.sh/@chenglou/pretext@latest';

// ─── Config ──────────────────────────────────────────────────────────────────
const COLORS = {
  bg: '#0f1b2d', panel: '#162035', panelBorder: '#1e3050',
  accent: '#2d7ef7', accentLight: '#4d9fff', gold: '#f5a623',
  green: '#27ae60', red: '#e74c3c', purple: '#8e44ad', teal: '#1abc9c',
  textPrimary: '#e8eaf0', textSecondary: '#8896aa', textMuted: '#4a5568',
  gridLine: '#1e2d44', white: '#ffffff',
};
const FONT = 'Sarabun, sans-serif';

// ─── SheetJS Loader ───────────────────────────────────────────────────────────
let _xlsxReady = null;
function loadSheetJS() {
  if (_xlsxReady) return _xlsxReady;
  _xlsxReady = new Promise((resolve, reject) => {
    if (window.XLSX) return resolve(window.XLSX);
    const s = document.createElement('script');
    s.src = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
    s.onload = () => resolve(window.XLSX);
    s.onerror = reject;
    document.head.appendChild(s);
  });
  return _xlsxReady;
}

// ─── Excel Parser ─────────────────────────────────────────────────────────────
function parseExcel(buffer) {
  const XLSX = window.XLSX;
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });

  const fmtDate = (v) => {
    if (!v) return null;
    if (v instanceof Date)
      return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,'0')}-${String(v.getDate()).padStart(2,'0')}`;
    return String(v);
  };
  const num = (v) => (v == null || v === '' || isNaN(Number(v))) ? null : Math.round(Number(v) * 100) / 100;

  const sh1 = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true });
  const meta = {
    accountNo: String(sh1[0]?.[1] ?? ''),
    firstDueDate: fmtDate(sh1[1]?.[1]),
    totalLoan: num(sh1[2]?.[1]) ?? 0,
  };
  const schedule = [];
  for (let i = 6; i < sh1.length; i++) {
    const r = sh1[i]; if (r[0] == null) continue;
    schedule.push({ no: Number(r[0]), dueDate: fmtDate(r[1]), pct: num(r[2]), total: num(r[3]), interest: num(r[4]), principal: num(r[5]) });
  }

  const sh2 = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[1]], { header: 1, raw: true });
  const payments = [];
  for (let i = 1; i < sh2.length; i++) {
    const r = sh2[i]; if (r[0] == null) continue;
    payments.push({ no: Number(r[0]), date: fmtDate(r[1]), source: String(r[4] ?? ''), amount: Math.abs(num(r[7]) ?? 0) });
  }
  payments.sort((a, b) => (a.date > b.date ? 1 : -1));

  const sh3 = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[2]], { header: 1, raw: true });
  const recal = [];
  for (let i = 1; i < sh3.length; i++) {
    const r = sh3[i];
    recal.push({
      loanBalance: num(r[1]), installmentNo: num(r[2]), dueDate: fmtDate(r[3]),
      principalBalance: num(r[6]), payDate: fmtDate(r[7]), payAmount: num(r[8]),
      reducePrincipal: num(r[10]), reduceInterest: num(r[11]), reducePenalty: num(r[12]),
      interestAccum: num(r[24]), penaltyAccum: num(r[36]),
    });
  }

  return { meta, schedule, payments, recal };
}

// ─── Stats ────────────────────────────────────────────────────────────────────
function computeStats({ meta, payments, recal }) {
  const totalLoan = meta.totalLoan;
  const totalPaid = payments.reduce((s, p) => s + p.amount, 0);
  let currentBalance = totalLoan;
  for (let i = recal.length - 1; i >= 0; i--) {
    if (recal[i].loanBalance != null) { currentBalance = recal[i].loanBalance; break; }
  }
  const principalPaid = totalLoan - currentBalance;
  const maxInterestAccum = Math.max(0, ...recal.map(r => r.interestAccum || 0));
  const maxPenaltyAccum = Math.max(0, ...recal.map(r => r.penaltyAccum || 0));
  const payByYear = {};
  payments.forEach(p => {
    const yr = (p.date || '').split('-')[0];
    if (yr) payByYear[yr] = (payByYear[yr] || 0) + p.amount;
  });
  const paidInstallments = new Set();
  recal.forEach(r => { if (r.reducePrincipal && r.installmentNo) paidInstallments.add(r.installmentNo); });
  return { totalLoan, totalPaid, currentBalance, principalPaid, maxInterestAccum, maxPenaltyAccum, payByYear, paidInstallments };
}

// ─── Pretext Text Drawing ─────────────────────────────────────────────────────
function makeDrawText(ctx, cache, W) {
  return function drawText(text, x, y, fontSize, color, align = 'left', maxW = 0) {
    text = String(text);
    const fontStr = `${fontSize}px ${FONT}`;
    const key = `${text}|${fontStr}`;
    if (!cache.has(key)) cache.set(key, prepareWithSegments(text, fontStr));
    const { lines } = layoutWithLines(cache.get(key), maxW || W, fontSize * 1.3);
    ctx.fillStyle = color;
    ctx.textBaseline = 'top';
    ctx.font = fontStr;
    lines.forEach((ln, i) => {
      const tw = ctx.measureText(ln.text).width;
      const dx = align === 'center' ? x - tw / 2 : align === 'right' ? x - tw : x;
      ctx.fillText(ln.text, dx, y + i * fontSize * 1.3);
    });
  };
}

// ─── Primitives ───────────────────────────────────────────────────────────────
function makeRR(ctx) {
  return function rr(x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y); ctx.lineTo(x + w - r, y); ctx.quadraticCurveTo(x + w, y, x + w, y + r);
    ctx.lineTo(x + w, y + h - r); ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
    ctx.lineTo(x + r, y + h); ctx.quadraticCurveTo(x, y + h, x, y + h - r);
    ctx.lineTo(x, y + r); ctx.quadraticCurveTo(x, y, x + r, y); ctx.closePath();
  };
}

const fmt = (n, d = 2) => (n ?? 0).toLocaleString('th-TH', { minimumFractionDigits: d, maximumFractionDigits: d });

// ─── Dashboard Renderer (per-canvas) ─────────────────────────────────────────
function createRenderer(canvas, data) {
  const ctx = canvas.getContext('2d');
  const cache = new Map();
  const scale = window.devicePixelRatio || 1;

  function applyScale() {
    const w = canvas.offsetWidth;
    const h = canvas.offsetHeight;
    canvas.width = w * scale;
    canvas.height = h * scale;
    ctx.setTransform(scale, 0, 0, scale, 0, 0);
    cache.clear();
  }

  function render() {
    applyScale();
    const W = canvas.width / scale;
    const H = canvas.height / scale;
    const drawText = makeDrawText(ctx, cache, W);
    const rr = makeRR(ctx);

    // ── Helpers ──
    function panel(x, y, w, h, accent = null) {
      rr(x, y, w, h, 8); ctx.fillStyle = COLORS.panel; ctx.fill();
      ctx.strokeStyle = accent ? accent + '44' : COLORS.panelBorder; ctx.lineWidth = 1; ctx.stroke();
      if (accent) { ctx.fillStyle = accent; rr(x + 1, y + 1, w - 2, 4, 3); ctx.fill(); }
    }

    function drawProgressBar(x, y, w, label, value, total, color) {
      const pct = Math.min(value / Math.max(total, 1), 1);
      const bw = w - 20;
      drawText(label, x + 10, y, 11, COLORS.textSecondary);
      drawText(`${fmt(value)} / ${fmt(total)}  (${(pct * 100).toFixed(1)}%)`, x + w - 10, y, 10, COLORS.textMuted, 'right');
      rr(x + 10, y + 16, bw, 14, 7); ctx.fillStyle = COLORS.panelBorder; ctx.fill();
      if (pct > 0.001) {
        const g = ctx.createLinearGradient(x + 10, 0, x + 10 + bw * pct, 0);
        g.addColorStop(0, color); g.addColorStop(1, color + 'aa');
        ctx.fillStyle = g; rr(x + 10, y + 16, bw * pct, 14, 7); ctx.fill();
      }
      return 36;
    }

    // ── Background ──
    ctx.fillStyle = COLORS.bg; ctx.fillRect(0, 0, W, H);
    ctx.strokeStyle = COLORS.gridLine + '33'; ctx.lineWidth = 0.5;
    for (let i = 0; i < W; i += 32) { ctx.beginPath(); ctx.moveTo(i, 0); ctx.lineTo(i, H); ctx.stroke(); }
    for (let i = 0; i < H; i += 32) { ctx.beginPath(); ctx.moveTo(0, i); ctx.lineTo(W, i); ctx.stroke(); }

    const s = computeStats(data);
    const pad = 12;
    let cy = pad;

    // ── 1. Header ──
    drawText('ภาพรวมบัญชีเงินกู้ กยศ.', W / 2, cy, 16, COLORS.white, 'center'); cy += 22;
    drawText(`บัญชี: ${data.meta.accountNo}   วงเงิน: ${fmt(data.meta.totalLoan)} บาท`, W / 2, cy, 11, COLORS.textSecondary, 'center'); cy += 18;

    // ── 2. KPI cards ──
    const kpis = [
      { label: 'ยอดกู้ทั้งหมด',  value: fmt(s.totalLoan),        color: COLORS.accent  },
      { label: 'ชำระแล้วสะสม',   value: fmt(s.totalPaid),        color: COLORS.green   },
      { label: 'เงินต้นคงเหลือ',  value: fmt(s.currentBalance),   color: COLORS.gold    },
      { label: 'ดอกเบี้ยสะสม',   value: fmt(s.maxInterestAccum), color: COLORS.purple  },
    ];
    const kpiW = (W - pad * 2 - 9) / 4;
    kpis.forEach((k, i) => {
      const kx = pad + i * (kpiW + 3);
      panel(kx, cy, kpiW, 72, k.color);
      drawText(k.label, kx + kpiW / 2, cy + 10, 11, COLORS.textSecondary, 'center');
      drawText(k.value, kx + kpiW / 2, cy + 30, 13, k.color, 'center');
      drawText('บาท', kx + kpiW / 2, cy + 52, 9, COLORS.textMuted, 'center');
    });
    cy += 80;

    // ── 3. Progress bars ──
    const pbH = 36 * 3 + 24;
    panel(pad, cy, W - pad * 2, pbH, COLORS.teal);
    drawText('ความคืบหน้าการชำระ', pad + 10, cy + 10, 12, COLORS.white);
    let pcy = cy + 28;
    pcy += drawProgressBar(pad, pcy, W - pad * 2, 'เงินต้นที่ชำระแล้ว', s.principalPaid, s.totalLoan, COLORS.green) + 2;
    pcy += drawProgressBar(pad, pcy, W - pad * 2, 'ยอดชำระสะสมทั้งหมด', s.totalPaid, s.totalLoan + s.maxInterestAccum, COLORS.accent) + 2;
    drawProgressBar(pad, pcy, W - pad * 2, 'ดอกเบี้ยสะสม', s.maxInterestAccum, Math.max(s.totalLoan * 0.05, s.maxInterestAccum), COLORS.purple);
    cy += pbH + 8;

    // ── 4. Bar chart (timeline) + Payment list  ──
    const yearKeys = Object.keys(s.payByYear).sort();
    const tlW = Math.floor((W - pad * 2) * 0.6);
    const lstW = W - pad * 2 - tlW - 8;
    const rowSectionH = 180;

    // Bar chart
    panel(pad, cy, tlW, rowSectionH);
    drawText('ยอดชำระจริงรายปี (บาท)', pad + 10, cy + 10, 12, COLORS.white);
    if (yearKeys.length > 0) {
      const gp = { t: 30, b: 28, l: 46, r: 10 };
      const gw = tlW - gp.l - gp.r, gh = rowSectionH - gp.t - gp.b;
      const gx = pad + gp.l, gy = cy + gp.t;
      const maxVal = Math.max(...Object.values(s.payByYear), 1);
      const barW = Math.min(gw / yearKeys.length - 6, 36);
      for (let i = 0; i <= 3; i++) {
        const ly = gy + gh * (1 - i / 3);
        ctx.strokeStyle = COLORS.gridLine; ctx.lineWidth = 0.5;
        ctx.beginPath(); ctx.moveTo(gx, ly); ctx.lineTo(gx + gw, ly); ctx.stroke();
        drawText(fmt(maxVal * i / 3, 0), gx - 4, ly - 5, 8, COLORS.textMuted, 'right');
      }
      yearKeys.forEach((yr, i) => {
        const val = s.payByYear[yr];
        const bh = (val / maxVal) * gh;
        const bx = gx + (i + 0.5) * (gw / yearKeys.length) - barW / 2, by = gy + gh - bh;
        const g = ctx.createLinearGradient(bx, by, bx, gy + gh);
        g.addColorStop(0, COLORS.accentLight); g.addColorStop(1, COLORS.accent + '44');
        ctx.fillStyle = g; rr(bx, by, barW, bh, 3); ctx.fill();
        drawText(yr.slice(2), bx + barW / 2, gy + gh + 4, 9, COLORS.textSecondary, 'center');
        if (bh > 18) drawText(fmt(val, 0), bx + barW / 2, by + 3, 8, COLORS.white, 'center');
      });
      ctx.strokeStyle = COLORS.panelBorder; ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(gx, gy); ctx.lineTo(gx, gy + gh); ctx.lineTo(gx + gw, gy + gh); ctx.stroke();
    } else {
      drawText('ยังไม่มีข้อมูลการชำระ', pad + tlW / 2, cy + rowSectionH / 2, 12, COLORS.textMuted, 'center');
    }

    // Payment list
    const lx = pad + tlW + 8;
    panel(lx, cy, lstW, rowSectionH);
    drawText('ประวัติการชำระ', lx + 10, cy + 10, 12, COLORS.white);
    const pRowH = 22, pMaxRows = Math.floor((rowSectionH - 32) / pRowH);
    [...data.payments].reverse().slice(0, pMaxRows).forEach((p, i) => {
      const ry = cy + 28 + i * pRowH;
      if (ry + pRowH > cy + rowSectionH - 6) return;
      if (i % 2 === 0) { ctx.fillStyle = COLORS.bg + '88'; rr(lx + 4, ry, lstW - 8, pRowH - 2, 2); ctx.fill(); }
      ctx.fillStyle = COLORS.accent;
      ctx.beginPath(); ctx.arc(lx + 14, ry + pRowH / 2, 3, 0, Math.PI * 2); ctx.fill();
      drawText(p.date || '', lx + 24, ry + 3, 10, COLORS.textSecondary);
      drawText(fmt(p.amount), lx + lstW - 8, ry + 3, 10, COLORS.green, 'right');
    });
    cy += rowSectionH + 8;

    // ── 5. Schedule table (left) + Interest chart (right) ──
    const tbW = Math.floor((W - pad * 2) * 0.55);
    const chW = W - pad * 2 - tbW - 8;
    const remH = Math.max(200, H - cy - pad);

    // ── 5a. Schedule table ──
    panel(pad, cy, tbW, remH);
    drawText('ตารางงวดผ่อนชำระ', pad + 10, cy + 10, 12, COLORS.white);
    const cols = [
      { label: 'งวด', w: 0.09 }, { label: 'กำหนดชำระ', w: 0.24 },
      { label: 'เงินต้น', w: 0.20 }, { label: 'ดอกเบี้ย', w: 0.20 },
      { label: 'รวม', w: 0.16 },  { label: 'สถานะ', w: 0.11 },
    ];
    const tRowH = 22, tx = pad + 6, tw = tbW - 12, theadY = cy + 30;
    rr(tx, theadY, tw, tRowH, 4); ctx.fillStyle = COLORS.panelBorder; ctx.fill();
    let cx2 = tx;
    cols.forEach(c => {
      drawText(c.label, cx2 + tw * c.w / 2, theadY + 5, 10, COLORS.textSecondary, 'center');
      cx2 += tw * c.w;
    });
    const tMaxRows = Math.floor((remH - 58) / tRowH);
    data.schedule.slice(0, tMaxRows).forEach((row, i) => {
      const ry = theadY + tRowH + i * tRowH;
      if (ry + tRowH > cy + remH - 6) return;
      const isPaid = s.paidInstallments.has(row.no);
      if (i % 2 === 0) { ctx.fillStyle = COLORS.bg + 'aa'; rr(tx, ry, tw, tRowH, 2); ctx.fill(); }
      if (isPaid) { ctx.fillStyle = COLORS.green + '18'; rr(tx, ry, tw, tRowH, 2); ctx.fill(); }
      const vals = [
        { t: row.no,                    a: 'center' },
        { t: row.dueDate || '',         a: 'center' },
        { t: fmt(row.principal),        a: 'right'  },
        { t: fmt(row.interest),         a: 'right'  },
        { t: fmt(row.total),            a: 'right'  },
        { t: isPaid ? '✓' : '○',       a: 'center' },
      ];
      let vcx = tx;
      vals.forEach((v, vi) => {
        const cw = tw * cols[vi].w;
        const color = vi === 5 ? (isPaid ? COLORS.green : COLORS.textMuted) : COLORS.textPrimary;
        const ax = v.a === 'right' ? vcx + cw - 6 : v.a === 'center' ? vcx + cw / 2 : vcx + 4;
        drawText(v.t, ax, ry + 5, 10, color, v.a);
        vcx += cw;
      });
    });

    // ── 5b. Interest chart ──
    const ichX = pad + tbW + 8;
    panel(ichX, cy, chW, remH);
    drawText('ดอกเบี้ยสะสม & เบี้ยปรับสะสม', ichX + 10, cy + 10, 12, COLORS.white);

    const iRows = data.recal.filter(r => r.interestAccum != null && r.payDate);
    if (iRows.length >= 2) {
      const pRows = data.recal.filter(r => r.penaltyAccum != null && r.penaltyAccum > 0 && r.payDate);
      const igp = { t: 36, b: 32, l: 52, r: 12 };
      const igw = chW - igp.l - igp.r, igh = remH - igp.t - igp.b;
      const igx = ichX + igp.l, igy = cy + igp.t;
      const maxI = Math.max(1, ...iRows.map(r => r.interestAccum));
      const dates = iRows.map(r => r.payDate).sort();
      const t0 = +new Date(dates[0]), span = Math.max(1, +new Date(dates[dates.length - 1]) - t0);
      const toX = d => igx + ((+new Date(d) - t0) / span) * igw;
      const toY = v => igy + igh - (v / maxI) * igh;

      for (let i = 0; i <= 4; i++) {
        const ly = igy + igh * (1 - i / 4);
        ctx.strokeStyle = COLORS.gridLine; ctx.lineWidth = 0.5; ctx.setLineDash([3, 3]);
        ctx.beginPath(); ctx.moveTo(igx, ly); ctx.lineTo(igx + igw, ly); ctx.stroke();
        ctx.setLineDash([]);
        drawText(fmt(maxI * i / 4, 0), igx - 4, ly - 5, 8, COLORS.textMuted, 'right');
      }

      // Area fill
      ctx.beginPath();
      iRows.forEach((r, i) => { i === 0 ? ctx.moveTo(toX(r.payDate), toY(r.interestAccum)) : ctx.lineTo(toX(r.payDate), toY(r.interestAccum)); });
      ctx.lineTo(toX(iRows[iRows.length - 1].payDate), igy + igh);
      ctx.lineTo(toX(iRows[0].payDate), igy + igh);
      ctx.closePath(); ctx.fillStyle = COLORS.purple + '33'; ctx.fill();

      // Interest line
      ctx.beginPath();
      iRows.forEach((r, i) => { i === 0 ? ctx.moveTo(toX(r.payDate), toY(r.interestAccum)) : ctx.lineTo(toX(r.payDate), toY(r.interestAccum)); });
      ctx.strokeStyle = COLORS.purple; ctx.lineWidth = 2; ctx.stroke();

      // Penalty line
      if (pRows.length > 1) {
        ctx.beginPath();
        pRows.forEach((r, i) => { i === 0 ? ctx.moveTo(toX(r.payDate), toY(r.penaltyAccum)) : ctx.lineTo(toX(r.payDate), toY(r.penaltyAccum)); });
        ctx.strokeStyle = COLORS.red; ctx.lineWidth = 2; ctx.stroke();
      }

      // Axes
      ctx.strokeStyle = COLORS.panelBorder; ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(igx, igy); ctx.lineTo(igx, igy + igh); ctx.lineTo(igx + igw, igy + igh); ctx.stroke();

      // Legend
      ctx.fillStyle = COLORS.purple; ctx.fillRect(ichX + chW - 110, cy + 16, 12, 3);
      drawText('ดอกเบี้ย', ichX + chW - 94, cy + 11, 10, COLORS.textSecondary);
      if (pRows.length > 1) {
        ctx.fillStyle = COLORS.red; ctx.fillRect(ichX + chW - 56, cy + 16, 12, 3);
        drawText('เบี้ยปรับ', ichX + chW - 40, cy + 11, 10, COLORS.textSecondary);
      }

      // X-axis date labels (first & last)
      drawText(dates[0].slice(0, 7), igx, igy + igh + 6, 8, COLORS.textMuted);
      drawText(dates[dates.length - 1].slice(0, 7), igx + igw, igy + igh + 6, 8, COLORS.textMuted, 'right');
    } else {
      drawText('ข้อมูลไม่เพียงพอสำหรับกราฟ', ichX + chW / 2, cy + remH / 2, 12, COLORS.textMuted, 'center');
    }
  }

  return { render };
}

// ─── Window Manager ───────────────────────────────────────────────────────────
let windows = [];
let zCounter = 10;
let winCount = 0;

function updateToolbar() {
  const n = windows.length;
  document.getElementById('file-count').style.display = n > 0 ? '' : 'none';
  document.getElementById('file-count').textContent = `${n} บัญชี`;
  document.getElementById('tile-btn').style.display = n > 1 ? '' : 'none';
  document.getElementById('closeall-btn').style.display = n > 0 ? '' : 'none';
  document.getElementById('welcome').classList.toggle('hidden', n > 0);
}

function focusWindow(fwin) {
  windows.forEach(w => w.el.classList.remove('focused'));
  fwin.el.classList.add('focused');
  fwin.el.style.zIndex = ++zCounter;
}

function createWindow(data, fileName) {
  winCount++;
  const workspace = document.getElementById('workspace');

  // Stagger position
  const offset = ((winCount - 1) % 8) * 28;
  const startX = 40 + offset;
  const startY = 60 + offset;
  const initW = Math.min(620, window.innerWidth - startX - 20);
  const initH = Math.min(700, window.innerHeight - startY - 20);

  const el = document.createElement('div');
  el.className = 'fwin focused';
  el.style.cssText = `left:${startX}px; top:${startY}px; width:${initW}px; height:${initH}px; z-index:${++zCounter}`;

  // Titlebar
  const title = document.createElement('div');
  title.className = 'fwin-title';
  title.innerHTML = `
    <div class="fwin-dot dot-close" title="ปิด"></div>
    <div class="fwin-dot dot-min" title="ย่อ"></div>
    <div class="fwin-dot dot-max" title="ขยาย"></div>
    <div class="fwin-name"><strong>${data.meta.accountNo || 'บัญชี ' + winCount}</strong> · ${fileName}</div>
  `;

  const body = document.createElement('div');
  body.className = 'fwin-body';

  const canvas = document.createElement('canvas');
  body.appendChild(canvas);

  const resizeHandle = document.createElement('div');
  resizeHandle.className = 'fwin-resize';

  el.appendChild(title);
  el.appendChild(body);
  el.appendChild(resizeHandle);
  workspace.appendChild(el);

  const renderer = createRenderer(canvas, data);

  const win = { el, data, renderer, minimized: false, maximized: false, prevRect: null };
  windows.push(win);

  // ── Dots
  title.querySelector('.dot-close').addEventListener('click', (e) => {
    e.stopPropagation();
    el.style.animation = 'none';
    el.style.transition = 'opacity 0.15s, transform 0.15s';
    el.style.opacity = '0';
    el.style.transform = 'scale(0.9)';
    setTimeout(() => {
      el.remove();
      windows = windows.filter(w => w !== win);
      updateToolbar();
    }, 150);
  });

  title.querySelector('.dot-min').addEventListener('click', (e) => {
    e.stopPropagation();
    win.minimized = !win.minimized;
    el.classList.toggle('minimized', win.minimized);
    if (!win.minimized) setTimeout(() => renderer.render(), 10);
  });

  title.querySelector('.dot-max').addEventListener('click', (e) => {
    e.stopPropagation();
    if (!win.maximized) {
      win.prevRect = { left: el.style.left, top: el.style.top, width: el.style.width, height: el.style.height };
      el.style.left = '0px'; el.style.top = '48px';
      el.style.width = window.innerWidth + 'px';
      el.style.height = (window.innerHeight - 48) + 'px';
      win.maximized = true;
    } else {
      const r = win.prevRect;
      el.style.left = r.left; el.style.top = r.top;
      el.style.width = r.width; el.style.height = r.height;
      win.maximized = false;
    }
    setTimeout(() => renderer.render(), 10);
  });

  // ── Drag
  let dragging = false, dragOffX = 0, dragOffY = 0;
  title.addEventListener('mousedown', (e) => {
    if (e.target.classList.contains('fwin-dot')) return;
    focusWindow(win);
    dragging = true;
    dragOffX = e.clientX - el.offsetLeft;
    dragOffY = e.clientY - el.offsetTop;
    document.addEventListener('mousemove', onDragMove);
    document.addEventListener('mouseup', onDragUp);
  });
  function onDragMove(e) {
    if (!dragging) return;
    let nx = e.clientX - dragOffX;
    let ny = e.clientY - dragOffY;
    nx = Math.max(-el.offsetWidth + 80, Math.min(window.innerWidth - 80, nx));
    ny = Math.max(0, Math.min(window.innerHeight - 38, ny));
    el.style.left = nx + 'px'; el.style.top = ny + 'px';
  }
  function onDragUp() {
    dragging = false;
    document.removeEventListener('mousemove', onDragMove);
    document.removeEventListener('mouseup', onDragUp);
  }

  // ── Resize
  let resizing = false, resizeStartX = 0, resizeStartY = 0, resizeStartW = 0, resizeStartH = 0;
  resizeHandle.addEventListener('mousedown', (e) => {
    e.preventDefault(); e.stopPropagation();
    focusWindow(win);
    resizing = true;
    resizeStartX = e.clientX; resizeStartY = e.clientY;
    resizeStartW = el.offsetWidth; resizeStartH = el.offsetHeight;
    document.addEventListener('mousemove', onResizeMove);
    document.addEventListener('mouseup', onResizeUp);
  });
  function onResizeMove(e) {
    if (!resizing) return;
    const nw = Math.max(380, resizeStartW + (e.clientX - resizeStartX));
    const nh = Math.max(300, resizeStartH + (e.clientY - resizeStartY));
    el.style.width = nw + 'px'; el.style.height = nh + 'px';
    renderer.render();
  }
  function onResizeUp() {
    resizing = false;
    document.removeEventListener('mousemove', onResizeMove);
    document.removeEventListener('mouseup', onResizeUp);
  }

  // ── Focus on click
  el.addEventListener('mousedown', () => focusWindow(win));

  // ── Initial render (after layout)
  setTimeout(() => {
    renderer.render();
    // Entrance animation
    el.style.opacity = '0'; el.style.transform = 'scale(0.92) translateY(8px)';
    el.style.transition = 'opacity 0.25s ease, transform 0.25s ease';
    requestAnimationFrame(() => { el.style.opacity = '1'; el.style.transform = 'scale(1) translateY(0)'; });
  }, 30);

  updateToolbar();
  return win;
}

// ─── Tile Layout ─────────────────────────────────────────────────────────────
window.tileWindows = function() {
  const n = windows.length;
  if (!n) return;
  const cols = Math.ceil(Math.sqrt(n));
  const rows = Math.ceil(n / cols);
  const areaW = window.innerWidth;
  const areaH = window.innerHeight - 48;
  const winW = Math.floor(areaW / cols) - 8;
  const winH = Math.floor(areaH / rows) - 8;
  windows.forEach((win, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    win.el.style.left = (col * (winW + 8) + 4) + 'px';
    win.el.style.top = (48 + row * (winH + 8) + 4) + 'px';
    win.el.style.width = winW + 'px';
    win.el.style.height = winH + 'px';
    win.minimized = false;
    win.el.classList.remove('minimized');
    setTimeout(() => win.renderer.render(), 20);
  });
};

window.closeAll = function() {
  windows.forEach(w => w.el.remove());
  windows = [];
  updateToolbar();
};

// ─── File Handling ────────────────────────────────────────────────────────────
async function handleFiles(files) {
  const arr = [...files].filter(f => f.name.match(/\.xlsx?$/i));
  if (!arr.length) return;
  await loadSheetJS();
  for (const file of arr) {
    try {
      const buf = await file.arrayBuffer();
      const data = parseExcel(buf);
      createWindow(data, file.name);
    } catch (e) {
      console.error('Error parsing', file.name, e);
      alert(`❌ ไม่สามารถอ่านไฟล์ ${file.name}: ${e.message}`);
    }
  }
}

// ─── Init ─────────────────────────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', () => {
  // Load font
  new FontFace('Sarabun', 'url(https://fonts.gstatic.com/s/sarabun/v14/DtVjJx26TKEr37c9YK5sulU.woff2)')
    .load().then(f => document.fonts.add(f)).catch(() => {});

  // File input
  const fileInput = document.getElementById('file-input');
  fileInput.addEventListener('change', e => { handleFiles(e.target.files); fileInput.value = ''; });

  // Global drag & drop
  const overlay = document.getElementById('drop-overlay');
  let dragCounter = 0;

  document.addEventListener('dragenter', (e) => {
    e.preventDefault();
    dragCounter++;
    if ([...e.dataTransfer.types].includes('Files')) overlay.classList.add('visible');
  });
  document.addEventListener('dragleave', () => {
    dragCounter--;
    if (dragCounter <= 0) { dragCounter = 0; overlay.classList.remove('visible'); }
  });
  document.addEventListener('dragover', (e) => { e.preventDefault(); });
  document.addEventListener('drop', (e) => {
    e.preventDefault();
    dragCounter = 0;
    overlay.classList.remove('visible');
    handleFiles(e.dataTransfer.files);
  });

  // Window resize → re-render all
  window.addEventListener('resize', () => {
    clearTimeout(window._rt);
    window._rt = setTimeout(() => {
      windows.forEach(w => { if (!w.minimized) w.renderer.render(); });
    }, 100);
  });
});
