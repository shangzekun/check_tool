const fileInput = document.getElementById('fileInput');
const uploadArea = document.getElementById('uploadArea');
const fileStatus = document.getElementById('fileStatus');
const rulesBox = document.getElementById('rules');
const descBox = document.getElementById('ruleDescriptions');
const summaryCards = document.getElementById('summaryCards');
const issuesTableBody = document.querySelector('#issuesTable tbody');
const checkBtn = document.getElementById('checkBtn');
const clearBtn = document.getElementById('clearBtn');
const exportBtn = document.getElementById('exportBtn');

let selectedFile = null;
let latestResult = null;
let rules = [];

const STORAGE_KEY = 'excel_check_tool_state_v1';

function saveState() {
  const state = {
    fileName: selectedFile ? selectedFile.name : null,
    fileSize: selectedFile ? selectedFile.size : null,
    selectedRuleIds: getSelectedRuleIds(),
    result: latestResult,
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function loadState() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return;
  try {
    const state = JSON.parse(raw);
    if (state.fileName) {
      fileStatus.textContent = `上次文件：${state.fileName}（${(state.fileSize / 1024).toFixed(1)} KB）- 刷新后需重新上传原文件才能再次检查`;
    }
    if (state.result) {
      latestResult = state.result;
      renderResult();
      exportBtn.disabled = false;
    }
    if (state.selectedRuleIds && state.selectedRuleIds.length > 0) {
      setTimeout(() => {
        state.selectedRuleIds.forEach((rid) => {
          const cb = document.querySelector(`input[name='rule'][value='${rid}']`);
          if (cb) cb.checked = true;
        });
      }, 0);
    }
  } catch (_) {
    localStorage.removeItem(STORAGE_KEY);
  }
}

function getSelectedRuleIds() {
  return Array.from(document.querySelectorAll("input[name='rule']:checked")).map((i) => i.value);
}

function renderRules() {
  rulesBox.innerHTML = '';
  descBox.innerHTML = '';
  rules.forEach((rule) => {
    const label = document.createElement('label');
    label.innerHTML = `<input type='checkbox' name='rule' value='${rule.id}'> ${rule.name}`;
    rulesBox.appendChild(label);

    const desc = document.createElement('div');
    desc.className = 'card';
    desc.innerHTML = `<strong>${rule.name}</strong><div class='muted'>${rule.description}</div>`;
    descBox.appendChild(desc);
  });
}

function renderResult() {
  if (!latestResult) return;
  summaryCards.innerHTML = '';
  latestResult.summary.forEach((item) => {
    const card = document.createElement('div');
    card.className = 'card';
    card.innerHTML = `
      <h4>${item.rule_name}</h4>
      <div>Error: ${item.error_count}</div>
      <div>Warning: ${item.warning_count}</div>
      <div>Info: ${item.info_count}</div>
      <div>Total: ${item.error_count + item.warning_count + item.info_count}</div>
    `;
    summaryCards.appendChild(card);
  });
  applyFilter();
}

function applyFilter() {
  issuesTableBody.innerHTML = '';
  if (!latestResult) return;
  const level = document.querySelector("input[name='level']:checked").value;
  const filtered = latestResult.issues.filter((i) => level === 'all' || i.severity === level);
  filtered.forEach((issue) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${issue.rule_name}</td>
      <td>${issue.severity}</td>
      <td>${issue.sheet || ''}</td>
      <td>${issue.row ?? ''}</td>
      <td>${issue.column || ''}</td>
      <td>${issue.message}</td>
    `;
    issuesTableBody.appendChild(tr);
  });
}

async function fetchRules() {
  const res = await fetch('/api/rules');
  const data = await res.json();
  rules = data.rules || [];
  renderRules();
  loadState();
}

async function runCheck() {
  if (!selectedFile) {
    alert('请先选择 .xlsx 文件');
    return;
  }
  const selectedRuleIds = getSelectedRuleIds();
  if (selectedRuleIds.length === 0) {
    alert('请至少勾选一个检查规则');
    return;
  }

  const fd = new FormData();
  fd.append('file', selectedFile);
  fd.append('rule_ids', selectedRuleIds.join(','));

  checkBtn.disabled = true;
  checkBtn.textContent = '检查中...';
  try {
    const res = await fetch('/api/check', { method: 'POST', body: fd });
    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.detail || '检查失败');
    }
    latestResult = await res.json();
    renderResult();
    exportBtn.disabled = false;
    saveState();
  } catch (e) {
    alert(e.message);
  } finally {
    checkBtn.disabled = false;
    checkBtn.textContent = '执行检查';
  }
}

async function exportReport() {
  if (!latestResult) return;
  const res = await fetch('/api/report', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(latestResult),
  });
  if (!res.ok) {
    alert('导出失败');
    return;
  }
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'excel_check_report.xlsx';
  a.click();
  URL.revokeObjectURL(url);
}

function clearAll() {
  selectedFile = null;
  latestResult = null;
  fileInput.value = '';
  fileStatus.textContent = '尚未选择文件';
  summaryCards.innerHTML = '';
  issuesTableBody.innerHTML = '';
  exportBtn.disabled = true;
  localStorage.removeItem(STORAGE_KEY);
}

uploadArea.addEventListener('click', () => fileInput.click());
uploadArea.addEventListener('dragover', (e) => {
  e.preventDefault();
  uploadArea.classList.add('dragover');
});
uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragover'));
uploadArea.addEventListener('drop', (e) => {
  e.preventDefault();
  uploadArea.classList.remove('dragover');
  const file = e.dataTransfer.files[0];
  if (!file) return;
  selectedFile = file;
  fileStatus.textContent = `当前文件：${file.name}（${(file.size / 1024).toFixed(1)} KB）`;
  saveState();
});

fileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  selectedFile = file;
  fileStatus.textContent = `当前文件：${file.name}（${(file.size / 1024).toFixed(1)} KB）`;
  saveState();
});

checkBtn.addEventListener('click', runCheck);
clearBtn.addEventListener('click', clearAll);
exportBtn.addEventListener('click', exportReport);
document.querySelectorAll("input[name='level']").forEach((radio) => radio.addEventListener('change', applyFilter));

fetchRules();
