'use strict';

// ── Constants ──────────────────────────────────────────────────────────────

const CATEGORIES = {
  food:          { label: 'Food',          icon: '🍽️', color: '#FF7043' },
  transport:     { label: 'Transport',     icon: '🚌', color: '#42A5F5' },
  accommodation: { label: 'Stay',          icon: '🏨', color: '#66BB6A' },
  entertainment: { label: 'Fun',           icon: '🎉', color: '#AB47BC' },
  shopping:      { label: 'Shopping',      icon: '🛍️', color: '#FFA726' },
  other:         { label: 'Other',         icon: '📦', color: '#78909C' },
};

const CURRENCIES = { USD: '$', EUR: '€', GBP: '£', JPY: '¥', CAD: 'C$', AUD: 'A$' };

const MEMBER_COLORS = ['#5C6BC0', '#26C6DA', '#EF9A9A'];

const STORAGE_KEY = 'splittrip_v1';

// ── State ──────────────────────────────────────────────────────────────────

let state = {
  initialized: false,
  tripName: 'Our Trip',
  currency: 'USD',
  members: [],       // [{ id, name, color }]
  expenses: [],      // [{ id, description, amount, category, date, notes, paidBy, splitAmong, splitType, customSplits, createdAt }]
  settlements: [],   // [{ id, from, to, amount, date, notes, createdAt }]
};

// UI state
let activeView    = 'dashboard';
let filterCat     = 'all';
let filterSearch  = '';
let splitType     = 'equal';
let pendingSettle = null;   // { from, to }

// ── Persistence ────────────────────────────────────────────────────────────

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) state = JSON.parse(raw);
  } catch (_) {}
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

// ── Helpers ────────────────────────────────────────────────────────────────

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function sym() {
  return CURRENCIES[state.currency] || '$';
}

function fmt(amount) {
  const s = sym();
  if (state.currency === 'JPY') return s + Math.round(amount).toLocaleString();
  return s + parseFloat(amount).toFixed(2);
}

function fmtDate(str) {
  const d = new Date(str + 'T12:00:00');
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}

function today() {
  return new Date().toISOString().slice(0, 10);
}

function member(id) {
  return state.members.find(m => m.id === id);
}

function initial(name) {
  return (name || '?').charAt(0).toUpperCase();
}

function round2(n) {
  return Math.round(n * 100) / 100;
}

function esc(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ── Balance Engine ─────────────────────────────────────────────────────────

function calcBalances() {
  const bal = {};
  state.members.forEach(m => { bal[m.id] = 0; });

  for (const exp of state.expenses) {
    const { paidBy, amount, splitAmong, splitType: st, customSplits } = exp;
    if (!splitAmong || splitAmong.length === 0) continue;

    bal[paidBy] = (bal[paidBy] || 0) + amount;

    if (st === 'custom' && customSplits) {
      for (const mid of splitAmong) {
        const share = parseFloat(customSplits[mid]) || 0;
        bal[mid] = (bal[mid] || 0) - share;
      }
    } else {
      const share = amount / splitAmong.length;
      for (const mid of splitAmong) {
        bal[mid] = (bal[mid] || 0) - share;
      }
    }
  }

  for (const s of state.settlements) {
    bal[s.from] = (bal[s.from] || 0) + s.amount;
    bal[s.to]   = (bal[s.to]   || 0) - s.amount;
  }

  // Round tiny floating point errors
  for (const id in bal) bal[id] = round2(bal[id]);
  return bal;
}

// Greedy debt simplification — returns [{ from, to, amount }]
function simplifyDebts(balances) {
  const credits = [];
  const debts   = [];

  for (const [id, bal] of Object.entries(balances)) {
    if (bal > 0.005)  credits.push({ id: Number(id), bal });
    if (bal < -0.005) debts.push({ id: Number(id),   bal: -bal });
  }

  credits.sort((a, b) => b.bal - a.bal);
  debts.sort((a, b)   => b.bal - a.bal);

  const txns = [];
  while (credits.length && debts.length) {
    const c = credits[0], d = debts[0];
    const amt = round2(Math.min(c.bal, d.bal));
    txns.push({ from: d.id, to: c.id, amount: amt });
    c.bal = round2(c.bal - amt);
    d.bal = round2(d.bal - amt);
    if (c.bal < 0.005) credits.shift();
    if (d.bal < 0.005) debts.shift();
  }
  return txns;
}

// ── Rendering helpers ──────────────────────────────────────────────────────

function avatarHtml(m, size = 36) {
  return `<div class="member-avatar" style="width:${size}px;height:${size}px;background:${m.color};font-size:${size * 0.4}px">${initial(m.name)}</div>`;
}

function renderActivityItem(item) {
  if (item._type === 'expense') {
    const payer = member(item.paidBy);
    const cat   = CATEGORIES[item.category] || CATEGORIES.other;
    return `
      <div class="activity-item" onclick="App.editExpense('${esc(item.id)}')">
        <div class="activity-icon" style="background:${cat.color}20;color:${cat.color}">${cat.icon}</div>
        <div class="activity-info">
          <div class="activity-desc">${esc(item.description)}</div>
          <div class="activity-meta">${esc(payer ? payer.name : '?')} paid · ${fmtDate(item.date)}</div>
        </div>
        <div class="activity-amount">${fmt(item.amount)}</div>
      </div>`;
  } else {
    const from = member(item.from);
    const to   = member(item.to);
    return `
      <div class="activity-item" style="cursor:default">
        <div class="activity-icon" style="background:#10B98120;color:#10B981">💸</div>
        <div class="activity-info">
          <div class="activity-desc">${esc(from ? from.name : '?')} paid ${esc(to ? to.name : '?')}</div>
          <div class="activity-meta">Settlement · ${fmtDate(item.date)}</div>
        </div>
        <div class="activity-amount settle-amt">${fmt(item.amount)}</div>
      </div>`;
  }
}

// ── Views ──────────────────────────────────────────────────────────────────

function renderDashboard() {
  const bal   = calcBalances();
  const debts = simplifyDebts(bal);

  // Balance cards
  const cards = state.members.map(m => {
    const b = bal[m.id] || 0;
    const cls = b > 0.005 ? 'positive' : b < -0.005 ? 'negative' : 'neutral';
    const label = b > 0.005 ? `gets back ${fmt(b)}` : b < -0.005 ? `owes ${fmt(-b)}` : 'settled up';
    return `
      <div class="balance-card">
        ${avatarHtml(m)}
        <div class="balance-name">${esc(m.name)}</div>
        <div class="balance-amt ${cls}">${cls === 'neutral' ? '✓' : fmt(Math.abs(b))}</div>
        <div style="font-size:11px;color:var(--text-muted)">${label}</div>
      </div>`;
  }).join('');
  document.getElementById('balance-cards').innerHTML = cards;

  // Summary tag
  const totalOwed = Object.values(bal).filter(b => b > 0).reduce((s, b) => s + b, 0);
  const summEl = document.getElementById('balance-summary-text');
  if (totalOwed < 0.01) {
    summEl.textContent = 'All settled up!';
    summEl.className = 'tag tag-success';
  } else {
    summEl.textContent = `${fmt(totalOwed)} to settle`;
    summEl.className = 'tag tag-warn';
  }

  // Who owes whom
  const debtsEl = document.getElementById('debts-list');
  if (debts.length === 0) {
    debtsEl.innerHTML = '<div class="empty-state">🎉 Everyone is even!</div>';
  } else {
    debtsEl.innerHTML = debts.map(d => {
      const f = member(d.from), t = member(d.to);
      return `
        <div class="debt-item">
          <div class="debt-avatars">
            <div class="debt-avatar" style="background:${f.color}">${initial(f.name)}</div>
            <div class="debt-arrow">→</div>
            <div class="debt-avatar" style="background:${t.color}">${initial(t.name)}</div>
          </div>
          <div class="debt-text flex-1"><strong>${esc(f.name)}</strong> owes <strong>${esc(t.name)}</strong></div>
          <div class="debt-amount">${fmt(d.amount)}</div>
        </div>`;
    }).join('');
  }

  // Recent activity (last 5)
  const all = [
    ...state.expenses.map(e => ({ ...e, _type: 'expense',    _ts: e.date + e.createdAt })),
    ...state.settlements.map(s => ({ ...s, _type: 'settlement', _ts: s.date + s.createdAt })),
  ].sort((a, b) => b._ts.localeCompare(a._ts)).slice(0, 5);

  const recentEl = document.getElementById('recent-activity');
  recentEl.innerHTML = all.length === 0
    ? '<div class="empty-state">No activity yet — add your first expense!</div>'
    : all.map(renderActivityItem).join('');
}

function renderExpenses() {
  let exps = [...state.expenses];
  if (filterCat !== 'all') exps = exps.filter(e => e.category === filterCat);
  if (filterSearch) {
    const q = filterSearch.toLowerCase();
    exps = exps.filter(e =>
      e.description.toLowerCase().includes(q) ||
      (e.notes || '').toLowerCase().includes(q)
    );
  }
  exps.sort((a, b) => (b.date + b.createdAt).localeCompare(a.date + a.createdAt));

  // Total bar
  const totalBar = document.getElementById('total-bar');
  if (exps.length > 0) {
    const total = exps.reduce((s, e) => s + e.amount, 0);
    totalBar.innerHTML = `<span>${exps.length} expense${exps.length !== 1 ? 's' : ''}</span><span>${fmt(total)} total</span>`;
    totalBar.classList.remove('hidden');
  } else {
    totalBar.classList.add('hidden');
  }

  const listEl = document.getElementById('expenses-list');
  if (exps.length === 0) {
    listEl.innerHTML = '<div class="empty-state">No expenses found</div>';
    return;
  }

  listEl.innerHTML = exps.map(exp => {
    const payer = member(exp.paidBy);
    const cat   = CATEGORIES[exp.category] || CATEGORIES.other;
    const n     = exp.splitAmong.length;
    const share = exp.splitType === 'equal' && n > 0 ? fmt(exp.amount / n) : null;

    return `
      <div class="expense-item">
        <div class="expense-cat-icon" style="background:${cat.color}20;color:${cat.color}">${cat.icon}</div>
        <div class="expense-info">
          <div class="expense-desc">${esc(exp.description)}</div>
          <div class="expense-meta">
            <span class="payer-badge" style="background:${payer ? payer.color + '25' : '#eee'};color:${payer ? payer.color : '#999'}">${esc(payer ? payer.name : '?')}</span>
            <span class="date-badge">${fmtDate(exp.date)}</span>
          </div>
          ${share ? `<div class="expense-share">${share}/person</div>` : ''}
        </div>
        <div class="expense-right">
          <div class="expense-amount">${fmt(exp.amount)}</div>
          <div class="expense-actions-row">
            <button class="icon-btn-sm" onclick="App.editExpense('${esc(exp.id)}')" title="Edit">✏️</button>
            <button class="icon-btn-sm" onclick="App.deleteExpense('${esc(exp.id)}')" title="Delete">🗑️</button>
          </div>
        </div>
      </div>`;
  }).join('');
}

function renderSettle() {
  const bal   = calcBalances();
  const debts = simplifyDebts(bal);

  const settleEl = document.getElementById('settle-list');
  if (debts.length === 0) {
    settleEl.innerHTML = '<div class="empty-state">🎉 All settled up! No payments needed.</div>';
  } else {
    settleEl.innerHTML = debts.map(d => {
      const f = member(d.from), t = member(d.to);
      return `
        <div class="settle-item">
          <div class="settle-avatars">
            <div class="settle-avatar" style="background:${f.color}">${initial(f.name)}</div>
            <div class="settle-mid">
              <div class="settle-amount-chip">${fmt(d.amount)}</div>
              <div class="settle-arrow-icon">→</div>
            </div>
            <div class="settle-avatar" style="background:${t.color}">${initial(t.name)}</div>
          </div>
          <div class="settle-desc"><strong>${esc(f.name)}</strong> owes <strong>${esc(t.name)}</strong> ${fmt(d.amount)}</div>
          <button class="btn btn-primary btn-sm btn-full" onclick="App.showSettleModal(${d.from},${d.to},${d.amount})">Mark as Paid</button>
        </div>`;
    }).join('');
  }

  // Settlement history
  const histEl = document.getElementById('settlement-history');
  const sorted = [...state.settlements].sort((a, b) => (b.date + b.createdAt).localeCompare(a.date + a.createdAt));
  histEl.innerHTML = sorted.length === 0
    ? '<div class="empty-state">No payments recorded yet</div>'
    : sorted.map(s => renderActivityItem({ ...s, _type: 'settlement' })).join('');
}

function renderSettings() {
  document.getElementById('settings-trip-name').value = state.tripName;
  document.getElementById('settings-name-1').value    = state.members[0]?.name || '';
  document.getElementById('settings-name-2').value    = state.members[1]?.name || '';
  document.getElementById('settings-name-3').value    = state.members[2]?.name || '';
  document.getElementById('settings-currency').value  = state.currency;

  // Color dots
  state.members.forEach((m, i) => {
    const dot = document.getElementById('dot-' + (i + 1));
    if (dot) dot.style.background = m.color;
  });

  const total = state.expenses.reduce((s, e) => s + e.amount, 0);
  document.getElementById('stat-total').textContent      = fmt(total);
  document.getElementById('stat-count').textContent      = state.expenses.length;
  document.getElementById('stat-per-person').textContent = fmt(total / 3);
}

// ── Modal: Add / Edit Expense ──────────────────────────────────────────────

function buildMemberPills(containerId, inputType, name, selectedIds) {
  const container = document.getElementById(containerId);
  container.innerHTML = state.members.map(m => {
    const checked = selectedIds.includes(m.id) ? 'checked' : '';
    const inputId = `pill-${name}-${m.id}`;
    return `
      <input type="${inputType}" id="${inputId}" name="${name}" value="${m.id}" class="member-pill-${inputType}" ${checked}>
      <label for="${inputId}" class="member-pill-label" style="--color:${m.color}">
        <span class="pill-avatar" style="background:${m.color}">${initial(m.name)}</span>
        ${esc(m.name)}
      </label>`;
  }).join('');

  if (inputType === 'checkbox') {
    container.querySelectorAll('input[type=checkbox]').forEach(cb => {
      cb.addEventListener('change', () => {
        if (splitType === 'custom') rebuildCustomSplits();
      });
    });
  }
}

function getCheckedIds(name) {
  return Array.from(document.querySelectorAll(`input[name="${name}"]:checked`)).map(el => Number(el.value));
}

function getRadioId(name) {
  const el = document.querySelector(`input[name="${name}"]:checked`);
  return el ? Number(el.value) : null;
}

function rebuildCustomSplits(existingSplits = {}) {
  const members = getCheckedIds('split-among');
  const container = document.getElementById('custom-splits');
  container.innerHTML = members.map(id => {
    const m   = member(id);
    const val = existingSplits[id] != null ? existingSplits[id] : '';
    return `
      <div class="custom-split-row">
        <div class="custom-split-avatar" style="background:${m.color}">${initial(m.name)}</div>
        <div class="custom-split-name">${esc(m.name)}</div>
        <input type="number" class="custom-split-input" data-mid="${id}" value="${val}" placeholder="0.00" min="0" step="0.01" oninput="App.updateSplitTotal()">
      </div>`;
  }).join('') + `
    <div class="split-total-row">
      <span>Total allocated</span>
      <span id="split-total-display">—</span>
    </div>`;
  App.updateSplitTotal();
}

// ── App Object ─────────────────────────────────────────────────────────────

const App = {
  // ── Boot ──────────────────────────────────────────────────────────────────
  init() {
    loadState();
    if (!state.initialized) {
      document.getElementById('setup-overlay').classList.remove('hidden');
    } else {
      this.boot();
    }
  },

  boot() {
    document.getElementById('trip-name-display').textContent = state.tripName;
    this.showView('dashboard');
  },

  setupComplete() {
    const tripName = document.getElementById('setup-trip-name').value.trim() || 'Our Trip';
    const names    = [1, 2, 3].map(i => document.getElementById('setup-name-' + i).value.trim() || `Person ${i}`);
    const currency = document.getElementById('setup-currency').value;

    state.tripName    = tripName;
    state.currency    = currency;
    state.initialized = true;
    state.members = names.map((name, i) => ({ id: i + 1, name, color: MEMBER_COLORS[i] }));

    saveState();
    document.getElementById('setup-overlay').classList.add('hidden');
    this.boot();
  },

  // ── Navigation ─────────────────────────────────────────────────────────────
  showView(view) {
    activeView = view;
    document.querySelectorAll('.view').forEach(el => el.classList.remove('active'));
    document.getElementById('view-' + view).classList.add('active');
    document.querySelectorAll('.nav-btn[data-view]').forEach(el => el.classList.remove('active'));
    const nb = document.querySelector(`.nav-btn[data-view="${view}"]`);
    if (nb) nb.classList.add('active');

    switch (view) {
      case 'dashboard': renderDashboard(); break;
      case 'expenses':  renderExpenses();  break;
      case 'settle':    renderSettle();    break;
      case 'settings':  renderSettings();  break;
    }
  },

  // ── Expense filters ────────────────────────────────────────────────────────
  filterExpenses() {
    filterSearch = document.getElementById('expense-search').value;
    renderExpenses();
  },

  setCategory(btn, cat) {
    document.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
    btn.classList.add('active');
    filterCat = cat;
    renderExpenses();
  },

  // ── Expense CRUD ───────────────────────────────────────────────────────────
  showAddExpense() {
    document.getElementById('modal-title').textContent = 'Add Expense';
    document.getElementById('expense-id').value         = '';
    document.getElementById('expense-desc').value       = '';
    document.getElementById('expense-amount').value     = '';
    document.getElementById('expense-category').value   = 'food';
    document.getElementById('expense-date').value       = today();
    document.getElementById('expense-notes').value      = '';

    splitType = 'equal';
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelector('.tab[data-type="equal"]').classList.add('active');
    document.getElementById('custom-splits').classList.add('hidden');

    buildMemberPills('paid-by-options',     'radio',    'paid-by',     [state.members[0]?.id]);
    buildMemberPills('split-among-options', 'checkbox', 'split-among', state.members.map(m => m.id));

    document.getElementById('expense-modal').classList.remove('hidden');
  },

  editExpense(id) {
    const exp = state.expenses.find(e => e.id === id);
    if (!exp) return;

    document.getElementById('modal-title').textContent    = 'Edit Expense';
    document.getElementById('expense-id').value           = exp.id;
    document.getElementById('expense-desc').value         = exp.description;
    document.getElementById('expense-amount').value       = exp.amount;
    document.getElementById('expense-category').value     = exp.category;
    document.getElementById('expense-date').value         = exp.date;
    document.getElementById('expense-notes').value        = exp.notes || '';

    splitType = exp.splitType || 'equal';
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelector(`.tab[data-type="${splitType}"]`).classList.add('active');

    buildMemberPills('paid-by-options',     'radio',    'paid-by',     [exp.paidBy]);
    buildMemberPills('split-among-options', 'checkbox', 'split-among', exp.splitAmong);

    if (splitType === 'custom') {
      document.getElementById('custom-splits').classList.remove('hidden');
      rebuildCustomSplits(exp.customSplits || {});
    } else {
      document.getElementById('custom-splits').classList.add('hidden');
    }

    document.getElementById('expense-modal').classList.remove('hidden');
  },

  saveExpense() {
    const desc     = document.getElementById('expense-desc').value.trim();
    const amount   = parseFloat(document.getElementById('expense-amount').value);
    const category = document.getElementById('expense-category').value;
    const date     = document.getElementById('expense-date').value;
    const notes    = document.getElementById('expense-notes').value.trim();
    const paidBy   = getRadioId('paid-by');
    const splitAmong = getCheckedIds('split-among');

    if (!desc)              return this.toast('Please enter a description', 'error');
    if (!amount || amount <= 0) return this.toast('Please enter a valid amount', 'error');
    if (!date)              return this.toast('Please select a date', 'error');
    if (paidBy == null)     return this.toast('Please select who paid', 'error');
    if (splitAmong.length === 0) return this.toast('Please select who to split with', 'error');

    let customSplits = {};
    if (splitType === 'custom') {
      let allocated = 0;
      document.querySelectorAll('.custom-split-input').forEach(inp => {
        const mid = Number(inp.dataset.mid);
        const val = parseFloat(inp.value) || 0;
        customSplits[mid] = val;
        allocated += val;
      });
      if (Math.abs(allocated - amount) > 0.01) {
        return this.toast(`Custom splits must total ${fmt(amount)}`, 'error');
      }
    }

    const existingId = document.getElementById('expense-id').value;
    const payload = { description: desc, amount, category, date, notes, paidBy, splitAmong, splitType, customSplits };

    if (existingId) {
      const idx = state.expenses.findIndex(e => e.id === existingId);
      if (idx !== -1) state.expenses[idx] = { ...state.expenses[idx], ...payload };
      this.toast('Expense updated');
    } else {
      state.expenses.push({ id: uid(), createdAt: new Date().toISOString(), ...payload });
      this.toast('Expense added', 'success');
    }

    saveState();
    this.closeExpenseModal();
    // Refresh whichever view is visible
    if (activeView === 'expenses') renderExpenses();
    else renderDashboard();
  },

  deleteExpense(id) {
    if (!confirm('Delete this expense?')) return;
    state.expenses = state.expenses.filter(e => e.id !== id);
    saveState();
    renderExpenses();
    this.toast('Expense deleted');
  },

  // ── Split type ─────────────────────────────────────────────────────────────
  setSplitType(type, btn) {
    splitType = type;
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    btn.classList.add('active');
    const customEl = document.getElementById('custom-splits');
    if (type === 'custom') {
      customEl.classList.remove('hidden');
      rebuildCustomSplits();
    } else {
      customEl.classList.add('hidden');
    }
  },

  updateSplitTotal() {
    const displayEl = document.getElementById('split-total-display');
    if (!displayEl) return;
    const inputs  = document.querySelectorAll('.custom-split-input');
    const total   = Array.from(inputs).reduce((s, i) => s + (parseFloat(i.value) || 0), 0);
    const target  = parseFloat(document.getElementById('expense-amount').value) || 0;
    const match   = Math.abs(total - target) < 0.01;
    displayEl.textContent = fmt(total);
    displayEl.className   = match ? 'split-match' : 'split-total-display split-mismatch';
  },

  // ── Settlements ────────────────────────────────────────────────────────────
  showSettleModal(fromId, toId, amount) {
    pendingSettle = { from: fromId, to: toId };
    const f = member(fromId), t = member(toId);

    document.getElementById('settle-modal-content').innerHTML = `
      <div class="settle-confirm">
        <div class="settle-avatar-lg" style="background:${f.color}">${initial(f.name)}</div>
        <div class="settle-confirm-arrow">→</div>
        <div class="settle-avatar-lg" style="background:${t.color}">${initial(t.name)}</div>
      </div>
      <p class="settle-confirm-text"><strong>${esc(f.name)}</strong> pays <strong>${esc(t.name)}</strong></p>`;

    document.getElementById('settle-amount').value = amount.toFixed(2);
    document.getElementById('settle-date').value   = today();
    document.getElementById('settle-notes').value  = '';
    document.getElementById('settle-modal').classList.remove('hidden');
  },

  saveSettlement() {
    if (!pendingSettle) return;
    const amount = parseFloat(document.getElementById('settle-amount').value);
    const date   = document.getElementById('settle-date').value;
    const notes  = document.getElementById('settle-notes').value.trim();

    if (!amount || amount <= 0) return this.toast('Enter a valid amount', 'error');
    if (!date)                  return this.toast('Select a date', 'error');

    state.settlements.push({
      id: uid(),
      createdAt: new Date().toISOString(),
      from: pendingSettle.from,
      to:   pendingSettle.to,
      amount, date, notes,
    });
    pendingSettle = null;
    saveState();
    this.closeSettleModal();
    renderSettle();
    this.toast('Payment recorded', 'success');
  },

  // ── Modal close ────────────────────────────────────────────────────────────
  closeExpenseModal() {
    document.getElementById('expense-modal').classList.add('hidden');
  },

  closeSettleModal() {
    document.getElementById('settle-modal').classList.add('hidden');
    pendingSettle = null;
  },

  closeOnBg(event, id) {
    if (event.target === event.currentTarget) {
      document.getElementById(id).classList.add('hidden');
      if (id === 'settle-modal') pendingSettle = null;
    }
  },

  // ── Settings ───────────────────────────────────────────────────────────────
  updateSettings() {
    state.tripName    = document.getElementById('settings-trip-name').value.trim() || 'Our Trip';
    state.currency    = document.getElementById('settings-currency').value;
    const names = [1, 2, 3].map(i => (document.getElementById('settings-name-' + i).value.trim() || `Person ${i}`));
    names.forEach((name, i) => { if (state.members[i]) state.members[i].name = name; });
    document.getElementById('trip-name-display').textContent = state.tripName;
    saveState();
    this.toast('Settings saved');
  },

  confirmReset() {
    if (!confirm('Delete all expenses and settlements? This cannot be undone.')) return;
    state.expenses    = [];
    state.settlements = [];
    saveState();
    renderSettings();
    this.toast('Data cleared');
  },

  // ── Toast ──────────────────────────────────────────────────────────────────
  toast(msg, type = '') {
    const el = document.getElementById('toast');
    el.textContent = msg;
    el.className   = `toast${type ? ' ' + type : ''}`;
    el.classList.remove('hidden');
    clearTimeout(App._toastTimer);
    App._toastTimer = setTimeout(() => el.classList.add('hidden'), 3000);
  },
};

// ── Boot ───────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => App.init());
