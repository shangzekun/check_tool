const fileInput = document.getElementById('fileInput');
const uploadArea = document.getElementById('uploadArea');
const fileStatus = document.getElementById('fileStatus');
const rulesBox = document.getElementById('rules');
const descBox = document.getElementById('ruleDescriptions');
const summaryCards = document.getElementById('summaryCards');
const issuesTableBody = document.querySelector('#issuesTable tbody');

const statTotal = document.getElementById('statTotal');
const statError = document.getElementById('statError');
const statWarning = document.getElementById('statWarning');
const statInfo = document.getElementById('statInfo');

const checkBtn = document.getElementById('checkBtn');
const clearBtn = document.getElementById('clearBtn');
const exportBtn = document.getElementById('exportBtn');
const rulesToggle = document.getElementById('rulesToggle');
const rulesPanel = document.getElementById('rulesPanel');

let selectedFile = null;
let latestResult = null;
let rules = [];

const STORAGE_KEY = 'excel_check_tool_state_v2';

function saveState() {
  localStorage.setItem(
    STORAGE_KEY,
    JSON.stringify({
      fileName: selectedFile?.name || null,
      fileSize: selectedFile?.size || null,
      selectedRuleIds: getSelectedRuleIds(),
      result: latestResult,
      rulesCollapsed: rulesPanel.classList.contains('collapsed'),
    })
  );
}

function syncStats() {
  if (!latestResult) {
    statTotal.textContent = '0';
    statError.textContent = '0';
    statWarning.textContent = '0';
    statInfo.textContent = '0';
    return;
  }

  const errorCount = latestResult.issues.filter((i) => i.severity === 'error').length;
  const warningCount = latestResult.issues.filter((i) => i.severity === 'warning').length;
  const infoCount = latestResult.issues.filter((i) => i.severity === 'info').length;
  statError.textContent = String(errorCount);
  statWarning.textContent = String(warningCount);
  statInfo.textContent = String(infoCount);
  statTotal.textContent = String(errorCount + warningCount + infoCount);
}

function loadState() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return;

  try {
    const state = JSON.parse(raw);
    if (state.fileName) {
      fileStatus.textContent = `上次文件：${state.fileName}（${(state.fileSize / 1024).toFixed(1)} KB）- 刷新后需重新上传原文件再次检查`;
    }
    if (state.result) {
      latestResult = state.result;
      renderResult();
      exportBtn.disabled = false;
    }
    if (Array.isArray(state.selectedRuleIds)) {
      setTimeout(() => {
        state.selectedRuleIds.forEach((rid) => {
          const input = document.querySelector(`input[name='rule'][value='${rid}']`);
          if (input) input.checked = true;
        });
      }, 0);
    }
    if (state.rulesCollapsed === true) {
      toggleRulesPanel(true);
    }
  } catch (_) {
    localStorage.removeItem(STORAGE_KEY);
  }
}

function toggleRulesPanel(forceCollapse = null) {
  const collapsed = forceCollapse ?? !rulesPanel.classList.contains('collapsed');
  rulesPanel.classList.toggle('collapsed', collapsed);
  rulesToggle.textContent = collapsed ? '展开' : '折叠';
  rulesToggle.setAttribute('aria-expanded', String(!collapsed));
  saveState();
}

function getSelectedRuleIds() {
  return Array.from(document.querySelectorAll("input[name='rule']:checked")).map((item) => item.value);
}

function renderRules() {
  rulesBox.innerHTML = '';
  descBox.innerHTML = '';

  rules.forEach((rule) => {
    const label = document.createElement('label');
    label.innerHTML = `<input type='checkbox' name='rule' value='${rule.id}' /> ${rule.name}`;
    rulesBox.appendChild(label);

    const desc = document.createElement('article');
    desc.className = 'desc-item';
    desc.innerHTML = `<strong>${rule.name}</strong><div class='muted'>${rule.description}</div>`;
    descBox.appendChild(desc);
  });

  rulesBox.querySelectorAll("input[name='rule']").forEach((i) => i.addEventListener('change', saveState));
}

function renderResult() {
  summaryCards.innerHTML = '';
  issuesTableBody.innerHTML = '';

  if (!latestResult) {
    syncStats();
    return;
  }

  latestResult.summary.forEach((item) => {
    const node = document.createElement('article');
    node.className = 'summary-item';
    node.innerHTML = `
      <h4>${item.rule_name}</h4>
      <div>Error: ${item.error_count}</div>
      <div>Warning: ${item.warning_count}</div>
      <div>Info: ${item.info_count}</div>
    `;
    summaryCards.appendChild(node);
  });

  syncStats();
  applyFilter();
}

function applyFilter() {
  issuesTableBody.innerHTML = '';
  if (!latestResult) return;

  const level = document.querySelector("input[name='level']:checked").value;
  const rows = latestResult.issues.filter((i) => level === 'all' || i.severity === level);

  rows.forEach((issue) => {
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
  const response = await fetch('/api/rules');
  const data = await response.json();
  rules = data.rules || [];
  renderRules();
  loadState();
}

async function runCheck() {
  if (!selectedFile) return alert('请先上传 .xlsx 文件');

  const selectedRuleIds = getSelectedRuleIds();
  if (selectedRuleIds.length === 0) return alert('请至少选择一个规则');

  const formData = new FormData();
  formData.append('file', selectedFile);
  formData.append('rule_ids', selectedRuleIds.join(','));

  checkBtn.disabled = true;
  checkBtn.textContent = '检查中...';
  try {
    const response = await fetch('/api/check', { method: 'POST', body: formData });
    if (!response.ok) {
      const err = await response.json();
      throw new Error(err.detail || '检查失败');
    }

    latestResult = await response.json();
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

  const response = await fetch('/api/report', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(latestResult),
  });

  if (!response.ok) return alert('导出失败');

  const blob = await response.blob();
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = 'excel_check_report.xlsx';
  link.click();
  URL.revokeObjectURL(link.href);
}

function clearAll() {
  selectedFile = null;
  latestResult = null;
  fileInput.value = '';
  fileStatus.textContent = '尚未选择文件';
  exportBtn.disabled = true;
  renderResult();
  localStorage.removeItem(STORAGE_KEY);
}

function attachUploadEvents() {
  uploadArea.addEventListener('click', (e) => {
    if (e.target.tagName !== 'BUTTON') fileInput.click();
  });

  const pickerBtn = uploadArea.querySelector('button');
  if (pickerBtn) {
    pickerBtn.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      fileInput.click();
    });
  }

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
}

function attachEvents() {
  attachUploadEvents();
  checkBtn.addEventListener('click', runCheck);
  clearBtn.addEventListener('click', clearAll);
  exportBtn.addEventListener('click', exportReport);
  rulesToggle.addEventListener('click', () => toggleRulesPanel());

  document
    .querySelectorAll("input[name='level']")
    .forEach((radio) => radio.addEventListener('change', applyFilter));

  if (window.innerWidth < 768) {
    toggleRulesPanel(true);
  }
}

attachEvents();
fetchRules();
