function request(url, options = {}) {
  const csrf = document.querySelector('meta[name="csrf-token"]')?.content;
  const headers = new Headers(options.headers || {});
  if (csrf && !['GET', 'HEAD'].includes((options.method || 'GET').toUpperCase())) headers.set('x-csrf-token', csrf);
  return fetch(url, { ...options, headers }).then(async response => {
    if (response.ok) return response;
    const data = await response.json().catch(() => ({}));
    throw new Error(data.error || `Request failed (${response.status})`);
  });
}

function showToast(message) {
  const toast = document.getElementById('toast');
  if (!toast) return;
  toast.textContent = message;
  toast.classList.add('visible');
}

function render(records) {
  const results = document.getElementById('generated-results');
  const count = document.getElementById('generated-count');
  results.replaceChildren();
  results.setAttribute('aria-busy', 'false');
  count.textContent = `${records.length} file${records.length === 1 ? '' : 's'}`;
  if (!records.length) {
    const empty = document.createElement('div');
    empty.className = 'empty-page-state';
    empty.innerHTML = '<span aria-hidden="true">▣</span><h2>No PowerPoints generated yet</h2><p>Generate a service deck and it will appear here for download.</p>';
    results.append(empty);
    return;
  }
  records.forEach(record => {
    const row = document.createElement('article');
    row.className = 'generated-row';
    const copy = document.createElement('div');
    const title = document.createElement('h3');
    title.textContent = `${record.service_name} · Revision ${record.revision}`;
    const details = document.createElement('p');
    details.textContent = `${record.service_date} · Generated ${new Date(record.generated_at).toLocaleString()} by ${record.generated_by} · Source revision ${record.source_revision}`;
    copy.append(title, details);
    const download = document.createElement('a');
    download.className = 'button button-secondary button-link';
    download.href = record.download_url;
    download.download = '';
    download.textContent = 'Download PowerPoint';
    row.append(copy, download);
    results.append(row);
  });
}

function boot() {
  document.querySelector('.sign-out')?.addEventListener('click', async () => {
    try {
      await request('/api/logout', { method: 'POST' });
      globalThis.location.assign('/login');
    } catch (error) {
      showToast(error.message);
    }
  });
  request('/api/generated').then(response => response.json()).then(render).catch(error => {
    const results = document.getElementById('generated-results');
    results.replaceChildren();
    results.setAttribute('aria-busy', 'false');
    const failure = document.createElement('div');
    failure.className = 'empty-page-state error-state';
    failure.innerHTML = '<span aria-hidden="true">!</span><h2>Generated files could not be loaded</h2><p>Try refreshing the page.</p>';
    results.append(failure);
    showToast(error.message);
  });
}

if (typeof document !== 'undefined') boot();
