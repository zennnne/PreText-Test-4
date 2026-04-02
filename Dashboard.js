import { prepareWithSegments, layoutWithLines } from 'https://esm.sh/@chenglou/pretext@latest';

// ─── Config ───────────────────────────────────────────────────────────────────
const COLORS = {
  bg: '#0f1b2d', panel: '#162035', panelBorder: '#1e3050',
  accent: '#2d7ef7', accentLight: '#4d9fff', gold: '#f5a623',
  green: '#27ae60', red: '#e74c3c', purple: '#8e44ad', teal: '#1abc9c',
  textPrimary: '#e8eaf0', textSecondary: '#8896aa', textMuted: '#4a5568',
  gridLine: '#1e2d44', white: '#ffffff',
};
const FONT = 'Sarabun, sans-serif';
const fmt = (n, d = 2) => (n ?? 0).toLocaleString('th-TH', { minimumFractionDigits: d, maximumFractionDigits: d });

// ─── SheetJS ──────────────────────────────────────────────────────────────────
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
  const num = (v) => (v == null || v === '' || isNaN(Number(v))) ? null : Math.round(Number(v)*100)/100;

  const sh1 = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header:1, raw:true });
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

  const sh2 = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[1]], { header:1, raw:true });
  const payments = [];
  for (let i = 1; i < sh2.length; i++) {
    const r = sh2[i]; if (r[0] == null) continue;
    payments.push({ no: Number(r[0]), date: fmtDate(r[1]), source: String(r[4]??''), amount: Math.abs(num(r[7])??0) });
  }
  payments.sort((a,b) => (a.date > b.date ? 1 : -1));

  const sh3 = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[2]], { header:1, raw:true });
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
  const totalPaid = payments.reduce((s,p) => s+p.amount, 0);
  let currentBalance = totalLoan;
  for (let i = recal.length-1; i >= 0; i--) {
    if (recal[i].loanBalance != null) { currentBalance = recal[i].loanBalance; break; }
  }
  const principalPaid = totalLoan - currentBalance;
  const maxInterestAccum = Math.max(0, ...recal.map(r => r.interestAccum||0));
  const maxPenaltyAccum  = Math.max(0, ...recal.map(r => r.penaltyAccum||0));
  const payByYear = {};
  payments.forEach(p => {
    const yr = (p.date||'').split('-')[0];
    if (yr) payByYear[yr] = (payByYear[yr]||0) + p.amount;
  });
  const paidInstallments = new Set();
  recal.forEach(r => { if (r.reducePrincipal && r.installmentNo) paidInstallments.add(r.installmentNo); });
  return { totalLoan, totalPaid, currentBalance, principalPaid, maxInterestAccum, maxPenaltyAccum, payByYear, paidInstallments };
}

// ─── Per-canvas renderer ──────────────────────────────────────────────────────
function createRenderer(canvas, data) {
  const ctx = canvas.getContext('2d');
  const textCache = new Map();

  function drawText(text, x, y, fontSize, color, align = 'left', maxW = 0) {
    text = String(text);
    const fontStr = `${fontSize}px ${FONT}`;
    const key = `${text}||${fontStr}`;
    if (!textCache.has(key)) textCache.set(key, prepareWithSegments(text, fontStr));
    const W = canvas.width / (window.devicePixelRatio || 1);
    const { lines } = layoutWithLines(textCache.get(key), maxW || W, fontSize * 1.3);
    ctx.fillStyle = color;
    ctx.textBaseline = 'top';
    ctx.font = fontStr;
    lines.forEach((ln, i) => {
      const tw = ctx.measureText(ln.text).width;
      const dx = align === 'center' ? x - tw/2 : align === 'right' ? x - tw : x;
      ctx.fillText(ln.text, dx, y + i * fontSize * 1.3);
    });
  }

  function rr(x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x+r, y); ctx.lineTo(x+w-r, y); ctx.quadraticCurveTo(x+w, y, x+w, y+r);
    ctx.lineTo(x+w, y+h-r); ctx.quadraticCurveTo(x+w, y+h, x+w-r, y+h);
    ctx.lineTo(x+r, y+h); ctx.quadraticCurveTo(x, y+h, x, y+h-r);
    ctx.lineTo(x, y+r); ctx.quadraticCurveTo(x, y, x+r, y); ctx.closePath();
  }

  function panel(x, y, w, h, accent = null) {
    rr(x, y, w, h, 8); ctx.fillStyle = COLORS.panel; ctx.fill();
    ctx.strokeStyle = accent ? accent+'44' : COLORS.panelBorder; ctx.lineWidth = 1; ctx.stroke();
    if (accent) { ctx.fillStyle = accent; rr(x+1, y+1, w-2, 4, 3); ctx.fill(); }
  }

  function progressBar(x, y, w, label, value, total, color) {
    const pct = Math.min(value / Math.max(total, 1), 1);
    const bw = w - 20;
    drawText(label, x+10, y, 11, COLORS.textSecondary);
    drawText(`${fmt(value)} / ${fmt(total)}  (${(pct*100).toFixed(1)}%)`, x+w-10, y, 10, COLORS.textMuted, 'right');
    rr(x+10, y+16, bw, 14, 7); ctx.fillStyle = COLORS.panelBorder; ctx.fill();
    if (pct > 0.001) {
      const g = ctx.createLinearGradient(x+10, 0, x+10+bw*pct, 0);
      g.addColorStop(0, color); g.addColorStop(1, color+'aa');
      ctx.fillStyle = g; rr(x+10, y+16, bw*pct, 14, 7); ctx.fill();
    }
    return 36;
  }

  function render() {
    const dpr = window.devicePixelRatio || 1;
    const W = canvas.offsetWidth;
    const H = canvas.offsetHeight;
    if (W === 0 || H === 0) return;

    canvas.width  = W * dpr;
    canvas.height = H * dpr;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    textCache.clear();

    const s = computeStats(data);
    const pad = 12;
    let cy = pad;

    // ── Background + dot grid ──
    ctx.fillStyle = COLORS.bg; ctx.fillRect(0, 0, W, H);
    ctx.strokeStyle = COLORS.gridLine + '33'; ctx.lineWidth = 0.5;
    for (let i = 0; i < W; i += 32) { ctx.beginPath(); ctx.moveTo(i,0); ctx.lineTo(i,H); ctx.stroke(); }
    for (let i = 0; i < H; i += 32) { ctx.beginPath(); ctx.moveTo(0,i); ctx.lineTo(W,i); ctx.stroke(); }

    // ── 1. Header ──
    drawText('ภาพรวมบัญชีเงินกู้ กยศ.', W/2, cy, 17, COLORS.white, 'center'); cy += 24;
    drawText(`บัญชี: ${data.meta.accountNo}   วงเงิน: ${fmt(data.meta.totalLoan)} บาท`, W/2, cy, 11, COLORS.textSecondary, 'center'); cy += 20;

    // ── 2. KPI cards ──
    const cards = [
      { label:'ยอดกู้ทั้งหมด',  value:fmt(s.totalLoan),        color:COLORS.accent  },
      { label:'ชำระแล้วสะสม',   value:fmt(s.totalPaid),        color:COLORS.green   },
      { label:'เงินต้นคงเหลือ',  value:fmt(s.currentBalance),   color:COLORS.gold    },
      { label:'ดอกเบี้ยสะสม',   value:fmt(s.maxInterestAccum), color:COLORS.purple  },
    ];
    const cardW = (W - pad*2 - 9) / 4;
    cards.forEach((c, i) => {
      const cx = pad + i*(cardW+3);
      panel(cx, cy, cardW, 72, c.color);
      drawText(c.label, cx+cardW/2, cy+10, 11, COLORS.textSecondary, 'center');
      drawText(c.value, cx+cardW/2, cy+30, 13, c.color, 'center');
      drawText('บาท', cx+cardW/2, cy+52, 9, COLORS.textMuted, 'center');
    });
    cy += 80;

    // ── 3. Progress bars ──
    const pbH = 36*3 + 24;
    panel(pad, cy, W-pad*2, pbH, COLORS.teal);
    drawText('ความคืบหน้าการชำระ', pad+12, cy+10, 12, COLORS.white);
    let pcy = cy+28;
    pcy += progressBar(pad, pcy, W-pad*2, 'เงินต้นที่ชำระแล้ว', s.principalPaid, s.totalLoan, COLORS.green) + 2;
    pcy += progressBar(pad, pcy, W-pad*2, 'ยอดชำระสะสมทั้งหมด', s.totalPaid, s.totalLoan+s.maxInterestAccum, COLORS.accent) + 2;
    progressBar(pad, pcy, W-pad*2, 'ดอกเบี้ยสะสม', s.maxInterestAccum, Math.max(s.totalLoan*0.05, s.maxInterestAccum), COLORS.purple);
    cy += pbH + 8;

    // ── 4. Bar chart (left 60%) + Payment list (right 40%) ──
    const tlW  = Math.floor((W-pad*2) * 0.60);
    const lstW = W - pad*2 - tlW - 8;
    const midH = Math.max(150, Math.min(200, H - cy - 300));

    // 4a. Bar chart
    panel(pad, cy, tlW, midH);
    drawText('ยอดชำระจริงรายปี (บาท)', pad+10, cy+10, 12, COLORS.white);
    const yearKeys = Object.keys(s.payByYear).sort();
    if (yearKeys.length > 0) {
      const gp = { t:30, b:28, l:48, r:10 };
      const gw = tlW-gp.l-gp.r, gh = midH-gp.t-gp.b;
      const gx = pad+gp.l, gy = cy+gp.t;
      const maxVal = Math.max(...Object.values(s.payByYear), 1);
      const barW = Math.min(gw/yearKeys.length - 6, 36);
      for (let i = 0; i <= 3; i++) {
        const ly = gy + gh*(1-i/3);
        ctx.strokeStyle = COLORS.gridLine; ctx.lineWidth = 0.5;
        ctx.beginPath(); ctx.moveTo(gx,ly); ctx.lineTo(gx+gw,ly); ctx.stroke();
        drawText(fmt(maxVal*i/3,0), gx-4, ly-5, 8, COLORS.textMuted, 'right');
      }
      yearKeys.forEach((yr, i) => {
        const val = s.payByYear[yr];
        const bh = (val/maxVal)*gh;
        const bx = gx+(i+0.5)*(gw/yearKeys.length)-barW/2, by = gy+gh-bh;
        const g = ctx.createLinearGradient(bx, by, bx, gy+gh);
        g.addColorStop(0, COLORS.accentLight); g.addColorStop(1, COLORS.accent+'44');
        ctx.fillStyle = g; rr(bx, by, barW, bh, 3); ctx.fill();
        drawText(yr.slice(2), bx+barW/2, gy+gh+4, 9, COLORS.textSecondary, 'center');
        if (bh > 18) drawText(fmt(val,0), bx+barW/2, by+3, 8, COLORS.white, 'center');
      });
      ctx.strokeStyle = COLORS.panelBorder; ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(gx,gy); ctx.lineTo(gx,gy+gh); ctx.lineTo(gx+gw,gy+gh); ctx.stroke();
    } else {
      drawText('ยังไม่มีข้อมูล', pad+tlW/2, cy+midH/2, 12, COLORS.textMuted, 'center');
    }

    // 4b. Payment list
    const lx = pad + tlW + 8;
    panel(lx, cy, lstW, midH);
    drawText('ประวัติการชำระ', lx+10, cy+10, 12, COLORS.white);
    const pRowH = 22, pMaxRows = Math.floor((midH-32)/pRowH);
    [...data.payments].reverse().slice(0, pMaxRows).forEach((p, i) => {
      const ry = cy+28+i*pRowH; if (ry+pRowH > cy+midH-4) return;
      if (i%2===0) { ctx.fillStyle = COLORS.bg+'88'; rr(lx+4, ry, lstW-8, pRowH-2, 2); ctx.fill(); }
      ctx.fillStyle = COLORS.accent;
      ctx.beginPath(); ctx.arc(lx+14, ry+pRowH/2, 3, 0, Math.PI*2); ctx.fill();
      drawText(p.date||'', lx+24, ry+4, 10, COLORS.textSecondary);
      drawText(fmt(p.amount)+' บ.', lx+lstW-8, ry+4, 10, COLORS.green, 'right');
    });
    cy += midH + 8;

    // ── 5. Schedule table (55%) + Interest chart (45%) — fills rest ──
    const tbW = Math.floor((W-pad*2) * 0.54);
    const chW = W - pad*2 - tbW - 8;
    const remH = Math.max(200, H - cy - pad);

    // 5a. Schedule table
    panel(pad, cy, tbW, remH);
    drawText('ตารางงวดผ่อนชำระ', pad+10, cy+10, 12, COLORS.white);

    const cols = [
      { label:'งวด',       fw:0.09 },
      { label:'กำหนดชำระ', fw:0.25 },
      { label:'เงินต้น',   fw:0.20 },
      { label:'ดอกเบี้ย',  fw:0.19 },
      { label:'รวม',       fw:0.15 },
      { label:'สถานะ',     fw:0.12 },
    ];
    const tx = pad+6, tw = tbW-12, theadY = cy+28, tRowH = 22;

    rr(tx, theadY, tw, tRowH, 4); ctx.fillStyle = '#1a2d44'; ctx.fill();
    let hcx = tx;
    cols.forEach(c => {
      drawText(c.label, hcx+tw*c.fw/2, theadY+5, 10, COLORS.textSecondary, 'center');
      hcx += tw*c.fw;
    });

    const tMaxRows = Math.floor((remH-55)/tRowH);
    data.schedule.slice(0, tMaxRows).forEach((row, i) => {
      const ry = theadY+tRowH+i*tRowH; if (ry+tRowH > cy+remH-6) return;
      const isPaid = s.paidInstallments.has(row.no);
      if (i%2===0) { ctx.fillStyle = COLORS.bg+'aa'; rr(tx, ry, tw, tRowH, 2); ctx.fill(); }
      if (isPaid)  { ctx.fillStyle = COLORS.green+'1a'; rr(tx, ry, tw, tRowH, 2); ctx.fill(); }
      const vals = [
        { t:String(row.no),      a:'center' },
        { t:row.dueDate||'',     a:'center' },
        { t:fmt(row.principal),  a:'right'  },
        { t:fmt(row.interest),   a:'right'  },
        { t:fmt(row.total),      a:'right'  },
        { t:isPaid?'✓ ชำระ':'○ รอ', a:'center' },
      ];
      let vcx = tx;
      vals.forEach((v, vi) => {
        const colW = tw*cols[vi].fw;
        const color = vi===5 ? (isPaid ? COLORS.green : COLORS.textMuted) : COLORS.textPrimary;
        const ax = v.a==='right' ? vcx+colW-6 : v.a==='center' ? vcx+colW/2 : vcx+4;
        drawText(v.t, ax, ry+5, 10, color, v.a);
        vcx += colW;
      });
    });

    // 5b. Interest accumulation chart
    const ichX = pad + tbW + 8;
    panel(ichX, cy, chW, remH);
    drawText('ดอกเบี้ยสะสม & เบี้ยปรับสะสม', ichX+10, cy+10, 12, COLORS.white);

    const iRows = data.recal.filter(r => r.interestAccum != null && r.payDate);
    if (iRows.length >= 2) {
      const pRows = data.recal.filter(r => r.penaltyAccum && r.penaltyAccum > 0 && r.payDate);
      const igp = { t:36, b:30, l:54, r:12 };
      const igw = chW-igp.l-igp.r, igh = remH-igp.t-igp.b;
      const igx = ichX+igp.l, igy = cy+igp.t;
      const maxI = Math.max(1, ...iRows.map(r => r.interestAccum));
      const allDates = iRows.map(r => r.payDate).sort();
      const t0   = +new Date(allDates[0]);
      const span = Math.max(1, +new Date(allDates[allDates.length-1]) - t0);
      const toX  = d => igx + ((+new Date(d)-t0)/span)*igw;
      const toY  = v => igy + igh - (v/maxI)*igh;

      for (let i = 0; i <= 4; i++) {
        const ly = igy+igh*(1-i/4);
        ctx.strokeStyle = COLORS.gridLine; ctx.lineWidth = 0.5; ctx.setLineDash([3,3]);
        ctx.beginPath(); ctx.moveTo(igx,ly); ctx.lineTo(igx+igw,ly); ctx.stroke();
        ctx.setLineDash([]);
        drawText(fmt(maxI*i/4, 0), igx-4, ly-5, 8, COLORS.textMuted, 'right');
      }

      // area fill
      ctx.beginPath();
      iRows.forEach((r,i) => i===0 ? ctx.moveTo(toX(r.payDate),toY(r.interestAccum)) : ctx.lineTo(toX(r.payDate),toY(r.interestAccum)));
      ctx.lineTo(toX(iRows[iRows.length-1].payDate), igy+igh);
      ctx.lineTo(toX(iRows[0].payDate), igy+igh);
      ctx.closePath(); ctx.fillStyle = COLORS.purple+'33'; ctx.fill();

      // interest line
      ctx.beginPath();
      iRows.forEach((r,i) => i===0 ? ctx.moveTo(toX(r.payDate),toY(r.interestAccum)) : ctx.lineTo(toX(r.payDate),toY(r.interestAccum)));
      ctx.strokeStyle = COLORS.purple; ctx.lineWidth = 2; ctx.stroke();

      // penalty line
      if (pRows.length >= 2) {
        ctx.beginPath();
        pRows.forEach((r,i) => i===0 ? ctx.moveTo(toX(r.payDate),toY(r.penaltyAccum)) : ctx.lineTo(toX(r.payDate),toY(r.penaltyAccum)));
        ctx.strokeStyle = COLORS.red; ctx.lineWidth = 2; ctx.stroke();
      }

      // axes
      ctx.strokeStyle = COLORS.panelBorder; ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(igx,igy); ctx.lineTo(igx,igy+igh); ctx.lineTo(igx+igw,igy+igh); ctx.stroke();

      // x labels
      drawText(allDates[0].slice(0,7), igx, igy+igh+6, 8, COLORS.textMuted);
      drawText(allDates[allDates.length-1].slice(0,7), igx+igw, igy+igh+6, 8, COLORS.textMuted, 'right');

      // legend
      const legY = cy+12;
      ctx.fillStyle = COLORS.purple; ctx.fillRect(ichX+chW-115, legY+4, 12, 3);
      drawText('ดอกเบี้ย', ichX+chW-99, legY, 10, COLORS.textSecondary);
      if (pRows.length >= 2) {
        ctx.fillStyle = COLORS.red; ctx.fillRect(ichX+chW-58, legY+4, 12, 3);
        drawText('เบี้ยปรับ', ichX+chW-42, legY, 10, COLORS.textSecondary);
      }
    } else {
      drawText('ข้อมูลไม่เพียงพอ', ichX+chW/2, cy+remH/2, 12, COLORS.textMuted, 'center');
    }
  }

  return { render };
}

// ─── Window Manager ───────────────────────────────────────────────────────────
let windows = [];
let zTop = 10;
let winCount = 0;

function updateToolbar() {
  const n = windows.length;
  document.getElementById('file-count').style.display   = n > 0 ? '' : 'none';
  document.getElementById('file-count').textContent      = `${n} บัญชี`;
  document.getElementById('tile-btn').style.display     = n > 1 ? '' : 'none';
  document.getElementById('closeall-btn').style.display = n > 0 ? '' : 'none';
  document.getElementById('welcome').classList.toggle('hidden', n > 0);
}

function focusWin(win) {
  windows.forEach(w => w.el.classList.remove('focused'));
  win.el.classList.add('focused');
  win.el.style.zIndex = ++zTop;
}

function createWindow(data, fileName) {
  winCount++;
  const workspace = document.getElementById('workspace');
  const offset = ((winCount-1) % 8) * 28;
  const sx = 40+offset, sy = 60+offset;
  const iw = Math.min(700, window.innerWidth-sx-20);
  const ih = Math.min(800, window.innerHeight-sy-20);

  const el = document.createElement('div');
  el.className = 'fwin focused';
  el.style.cssText = `left:${sx}px;top:${sy}px;width:${iw}px;height:${ih}px;z-index:${++zTop}`;

  const titlebar = document.createElement('div');
  titlebar.className = 'fwin-title';
  titlebar.innerHTML = `
    <div class="fwin-dot dot-close"></div>
    <div class="fwin-dot dot-min"></div>
    <div class="fwin-dot dot-max"></div>
    <div class="fwin-name"><strong>${data.meta.accountNo || 'บัญชี '+winCount}</strong> · ${fileName}</div>
  `;

  const body = document.createElement('div');
  body.className = 'fwin-body';

  const canvas = document.createElement('canvas');
  canvas.style.cssText = 'display:block;width:100%;height:100%;';
  body.appendChild(canvas);

  const resizeHandle = document.createElement('div');
  resizeHandle.className = 'fwin-resize';

  el.appendChild(titlebar);
  el.appendChild(body);
  el.appendChild(resizeHandle);
  workspace.appendChild(el);

  const renderer = createRenderer(canvas, data);
  const win = { el, renderer, minimized:false, maximized:false, prevRect:null };
  windows.push(win);
  focusWin(win);

  // dots
  titlebar.querySelector('.dot-close').addEventListener('click', e => {
    e.stopPropagation();
    el.style.transition = 'opacity 0.15s, transform 0.15s';
    el.style.opacity='0'; el.style.transform='scale(0.9)';
    setTimeout(() => { el.remove(); windows=windows.filter(w=>w!==win); updateToolbar(); }, 150);
  });
  titlebar.querySelector('.dot-min').addEventListener('click', e => {
    e.stopPropagation();
    win.minimized = !win.minimized;
    el.classList.toggle('minimized', win.minimized);
    if (!win.minimized) setTimeout(()=>renderer.render(), 10);
  });
  titlebar.querySelector('.dot-max').addEventListener('click', e => {
    e.stopPropagation();
    if (!win.maximized) {
      win.prevRect = { left:el.style.left, top:el.style.top, width:el.style.width, height:el.style.height };
      el.style.left='0px'; el.style.top='48px';
      el.style.width=window.innerWidth+'px'; el.style.height=(window.innerHeight-48)+'px';
      win.maximized=true;
    } else {
      const r=win.prevRect;
      el.style.left=r.left; el.style.top=r.top; el.style.width=r.width; el.style.height=r.height;
      win.maximized=false;
    }
    setTimeout(()=>renderer.render(), 20);
  });

  // drag
  let dragging=false, dox=0, doy=0;
  titlebar.addEventListener('mousedown', e => {
    if (e.target.classList.contains('fwin-dot')) return;
    focusWin(win); dragging=true;
    dox=e.clientX-el.offsetLeft; doy=e.clientY-el.offsetTop;
    document.addEventListener('mousemove', onMove); document.addEventListener('mouseup', onUp);
  });
  function onMove(e) {
    if (!dragging) return;
    el.style.left = Math.max(-el.offsetWidth+80, Math.min(window.innerWidth-80,  e.clientX-dox))+'px';
    el.style.top  = Math.max(0,                  Math.min(window.innerHeight-38, e.clientY-doy))+'px';
  }
  function onUp() { dragging=false; document.removeEventListener('mousemove',onMove); document.removeEventListener('mouseup',onUp); }

  // resize
  let resizing=false, rsx=0, rsy=0, rsw=0, rsh=0;
  resizeHandle.addEventListener('mousedown', e => {
    e.preventDefault(); e.stopPropagation();
    focusWin(win); resizing=true;
    rsx=e.clientX; rsy=e.clientY; rsw=el.offsetWidth; rsh=el.offsetHeight;
    document.addEventListener('mousemove', onRM); document.addEventListener('mouseup', onRU);
  });
  function onRM(e) {
    if (!resizing) return;
    el.style.width  = Math.max(520, rsw+(e.clientX-rsx))+'px';
    el.style.height = Math.max(400, rsh+(e.clientY-rsy))+'px';
    renderer.render();
  }
  function onRU() { resizing=false; document.removeEventListener('mousemove',onRM); document.removeEventListener('mouseup',onRU); }

  el.addEventListener('mousedown', ()=>focusWin(win));

  // ResizeObserver triggers render when body actually has dimensions
  const ro = new ResizeObserver(() => renderer.render());
  ro.observe(body);

  // entrance animation
  el.style.opacity='0'; el.style.transform='scale(0.93) translateY(10px)';
  el.style.transition='opacity 0.22s ease, transform 0.22s ease';
  requestAnimationFrame(()=>{ el.style.opacity='1'; el.style.transform='scale(1) translateY(0)'; });

  updateToolbar();
  return win;
}

// ─── Tile / Close All ─────────────────────────────────────────────────────────
window.tileWindows = function() {
  const n=windows.length; if(!n) return;
  const cols=Math.ceil(Math.sqrt(n)), rows=Math.ceil(n/cols);
  const ww=Math.floor(window.innerWidth/cols)-8;
  const wh=Math.floor((window.innerHeight-48)/rows)-8;
  windows.forEach((win,i)=>{
    const col=i%cols, row=Math.floor(i/cols);
    win.el.style.left=(col*(ww+8)+4)+'px';
    win.el.style.top=(48+row*(wh+8)+4)+'px';
    win.el.style.width=ww+'px';
    win.el.style.height=wh+'px';
    win.minimized=false; win.el.classList.remove('minimized');
    setTimeout(()=>win.renderer.render(), 30);
  });
};
window.closeAll = function() { windows.forEach(w=>w.el.remove()); windows=[]; updateToolbar(); };

// ─── File Handling ────────────────────────────────────────────────────────────
async function handleFiles(files) {
  const arr=[...files].filter(f=>f.name.match(/\.xlsx?$/i));
  if(!arr.length) return;
  await loadSheetJS();
  for(const file of arr) {
    try { createWindow(parseExcel(await file.arrayBuffer()), file.name); }
    catch(e) { console.error(e); alert(`❌ อ่านไฟล์ ${file.name} ไม่ได้: ${e.message}`); }
  }
}

// ─── Init ─────────────────────────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', () => {
  new FontFace('Sarabun','url(https://fonts.gstatic.com/s/sarabun/v14/DtVjJx26TKEr37c9YK5sulU.woff2)')
    .load().then(f=>document.fonts.add(f)).catch(()=>{});

  const fileInput=document.getElementById('file-input');
  fileInput.addEventListener('change', e=>{ handleFiles(e.target.files); fileInput.value=''; });

  const overlay=document.getElementById('drop-overlay');
  let dc=0;
  document.addEventListener('dragenter', e=>{ e.preventDefault(); dc++; if([...e.dataTransfer.types].includes('Files')) overlay.classList.add('visible'); });
  document.addEventListener('dragleave', ()=>{ dc--; if(dc<=0){dc=0; overlay.classList.remove('visible');} });
  document.addEventListener('dragover',  e=>e.preventDefault());
  document.addEventListener('drop', e=>{ e.preventDefault(); dc=0; overlay.classList.remove('visible'); handleFiles(e.dataTransfer.files); });

  window.addEventListener('resize', ()=>{
    clearTimeout(window._rt);
    window._rt=setTimeout(()=>windows.forEach(w=>{ if(!w.minimized) w.renderer.render(); }), 100);
  });
});
