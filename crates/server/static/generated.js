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

// A plain download link saves whatever comes back, so a failed request used to land in the
// staff member's downloads as a .pptx full of JSON that PowerPoint refuses to open. Fetching
// the deck first means a failure is reported instead of saved.
async function downloadDeck(record, button) {
  const original = button.textContent;
  button.disabled = true;
  button.textContent = 'Downloading…';
  try {
    const response = await request(record.download_url);
    const blob = await response.blob();
    const disposition = response.headers.get('content-disposition') || '';
    const filename = disposition.match(/filename="([^"]+)"/)?.[1]
      || `${record.service_name}-${record.service_date}-r${record.revision}.pptx`;
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 1000);
  } catch (error) {
    showToast(error.message);
  } finally {
    button.disabled = false;
    button.textContent = original;
  }
}

function componentSummary(component) {
  const detail = [
    component.reference,
    component.selection,
    component.title,
    component.key,
    Array.isArray(component.rows) && component.rows.length ? `${component.rows.length} notice rows` : '',
  ].find(value => typeof value === 'string' && value.trim());
  return detail ? detail.trim() : '';
}

function renderSnapshot(service, panel) {
  panel.replaceChildren();
  const summary = document.createElement('p');
  summary.className = 'snapshot-summary';
  summary.textContent = `${service.preset} · ${service.components.length} items · service revision ${service.revision} · last edited by ${service.audit.updated_by}`;
  const list = document.createElement('ol');
  list.className = 'snapshot-components';
  service.components.forEach(component => {
    const item = document.createElement('li');
    const heading = document.createElement('strong');
    heading.textContent = component.heading || component.title || component.type || 'Item';
    item.append(heading);
    const detail = componentSummary(component);
    if (detail && detail !== heading.textContent) {
      const note = document.createElement('span');
      note.textContent = ` — ${detail}`;
      item.append(note);
    }
    list.append(item);
  });
  panel.append(summary, list);
}

async function toggleSnapshot(record, button, panel) {
  if (!panel.hidden) {
    panel.hidden = true;
    button.setAttribute('aria-expanded', 'false');
    button.textContent = 'View contents';
    return;
  }
  button.disabled = true;
  try {
    const service = await request(record.snapshot_url).then(response => response.json());
    renderSnapshot(service, panel);
    panel.hidden = false;
    button.setAttribute('aria-expanded', 'true');
    button.textContent = 'Hide contents';
  } catch (error) {
    showToast(error.message);
  } finally {
    button.disabled = false;
  }
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

    const actions = document.createElement('div');
    actions.className = 'generated-actions';
    const panel = document.createElement('div');
    panel.className = 'snapshot-panel';
    panel.hidden = true;

    // Decks generated before the service was recorded alongside them have nothing to show.
    if (record.snapshot_url) {
      const contents = document.createElement('button');
      contents.type = 'button';
      contents.className = 'button button-secondary';
      contents.textContent = 'View contents';
      contents.setAttribute('aria-expanded', 'false');
      contents.addEventListener('click', () => toggleSnapshot(record, contents, panel));
      actions.append(contents);
    }

    const download = document.createElement('button');
    download.type = 'button';
    download.className = 'button button-primary';
    download.textContent = 'Download PowerPoint';
    download.addEventListener('click', () => downloadDeck(record, download));
    actions.append(download);

    row.append(copy, actions, panel);
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
