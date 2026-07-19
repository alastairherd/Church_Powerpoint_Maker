const csrf = document.querySelector('meta[name="csrf-token"]')?.content;
const form = document.getElementById('settings-form');
const licence = document.getElementById('ccli-number');
const version = document.getElementById('settings-version');
const save = document.getElementById('save-settings');
const songCount = document.getElementById('admin-song-count');
const toast = document.getElementById('toast');
let toastTimer = null;

async function request(url, options = {}) {
  const headers = new Headers(options.headers || {});
  if (csrf && !['GET', 'HEAD'].includes((options.method || 'GET').toUpperCase())) headers.set('x-csrf-token', csrf);
  if (options.body) headers.set('content-type', 'application/json');
  const response = await fetch(url, { ...options, headers });
  if (!response.ok) {
    const data = await response.json().catch(() => ({}));
    throw new Error(data.error || `Request failed (${response.status})`);
  }
  return response;
}

function showToast(message) {
  toast.textContent = message;
  toast.classList.add('visible');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove('visible'), 3200);
}

function showSettings(settings) {
  licence.value = settings.ccli_licence_number;
  version.textContent = `Version ${settings.version}, last changed by ${settings.created_by}.`;
}

form.addEventListener('submit', async event => {
  event.preventDefault();
  save.disabled = true; save.textContent = 'Saving…';
  try {
    const response = await request('/api/settings', { method: 'PUT', body: JSON.stringify({ ccli_licence_number: licence.value }) });
    showSettings(await response.json()); showToast('CCLI setting saved.');
  } catch (error) { showToast(error.message); }
  finally { save.disabled = false; save.textContent = 'Save CCLI setting'; }
});

document.querySelector('.sign-out').addEventListener('click', async () => {
  await request('/api/logout', { method: 'POST' });
  location.assign('/login');
});

Promise.all([
  request('/api/settings').then(response => response.json()),
  request('/api/songs').then(response => response.json())
]).then(([settings, songs]) => {
  showSettings(settings);
  songCount.textContent = `${songs.length} active song${songs.length === 1 ? '' : 's'}`;
}).catch(error => showToast(error.message));
