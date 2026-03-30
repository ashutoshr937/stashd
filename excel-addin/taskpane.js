// ─────────────────────────────────────────────────────────────
//  Stashd — Excel Task Pane
// ─────────────────────────────────────────────────────────────
const STORAGE_KEY = 'stashd_token';

const FIELD_LABELS = [
  'Name', 'LinkedIn URL', 'Headline', 'About', 'Location',
  'Current Company', 'Current Job Title', 'Company Start Date', 'Company Tenure',
  'Connections', 'Followers', 'Premium', 'Open to Work', 'Hiring', 'Verified',
  'Employment Type', 'Workplace Type', 'Role Location',
  'Previous Company', 'Previous Job Title', 'Total Experience', 'No. of Roles',
  'School', 'Degree & Field', 'Skills', 'Languages',
];

Office.onReady(async () => {
  // Render field list
  const list = document.getElementById('fields-list');
  list.innerHTML = FIELD_LABELS.map(l => `<span>${l}</span>`).join('');

  // Load saved token
  const statusEl = document.getElementById('token-status');
  const tokenInput = document.getElementById('token-input');

  async function loadToken() {
    try {
      const t = await OfficeRuntime.storage.getItem(STORAGE_KEY);
      if (t) {
        tokenInput.value = t;
        setStatus('ok', 'Token saved ✓ — ready to use =STASHD.FETCH()');
      } else {
        setStatus('warn', 'No token saved. Enter your Apify token below.');
      }
    } catch {
      setStatus('warn', 'Could not read stored token.');
    }
  }

  function setStatus(type, msg) {
    statusEl.className = 'status ' + type;
    statusEl.textContent = msg;
  }

  document.getElementById('save-btn').addEventListener('click', async () => {
    const val = tokenInput.value.trim();
    if (!val) { setStatus('error', 'Please enter a token.'); return; }
    try {
      await OfficeRuntime.storage.setItem(STORAGE_KEY, val);
      setStatus('ok', 'Token saved ✓ — ready to use =STASHD.FETCH()');
    } catch {
      setStatus('error', 'Failed to save token.');
    }
  });

  document.getElementById('clear-btn').addEventListener('click', async () => {
    try {
      await OfficeRuntime.storage.removeItem(STORAGE_KEY);
      tokenInput.value = '';
      setStatus('warn', 'Token cleared. Enter a new token to continue.');
    } catch {
      setStatus('error', 'Failed to clear token.');
    }
  });

  await loadToken();
});
