const csrf = document.querySelector('meta[name="csrf-token"]')?.content;
const search = document.getElementById('song-search');
const results = document.getElementById('library-results');
const preview = document.getElementById('song-preview');
const count = document.getElementById('song-count');
const toast = document.getElementById('toast');
const addDialog = document.getElementById('add-song-dialog');
const addForm = document.getElementById('add-song-form');
const addError = document.getElementById('add-song-error');
const addSubmit = document.getElementById('add-song-submit');
const replaceInput = document.getElementById('replace-song-file');
const PPTX_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
const MAX_UPLOAD_BYTES = 25 * 1024 * 1024;
let selectedId = null;
let searchTimer = null;
let requestNumber = 0;
let toastTimer = null;

async function request(url, options = {}) {
  const headers = new Headers(options.headers || {});
  if (csrf && !['GET', 'HEAD'].includes((options.method || 'GET').toUpperCase())) headers.set('x-csrf-token', csrf);
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

function friendlyRights(value) {
  return ({ public_domain: 'Public domain', ccli_covered: 'CCLI covered', direct_permission: 'Direct permission', unknown: 'Rights unknown' })[value] || value;
}

export function parseAliases(value) {
  return (value || '').split(',').map(alias => alias.trim()).filter(Boolean);
}

/** Returns a message for the mistakes worth catching before we spend an upload, or null. */
export function describeFileProblem(file) {
  if (!file) return 'Choose a PowerPoint file.';
  if (!file.name.toLowerCase().endsWith('.pptx')) return 'The file must be a PowerPoint .pptx, not a .ppt or an exported PDF.';
  if (!file.size) return 'That file is empty.';
  if (file.size > MAX_UPLOAD_BYTES) return 'That PowerPoint is larger than 25 MB. Compress its images and try again.';
  return null;
}

/** Header values must be latin-1, so smart quotes and the like are dropped before sending. */
export function headerSafeFilename(name) {
  return (name || '').replace(/[^\x20-\x7e]/g, '').slice(0, 180) || 'song.pptx';
}

async function uploadPowerPoint(songId, file) {
  const response = await request(`/api/songs/${encodeURIComponent(songId)}/upload`, {
    method: 'POST',
    headers: { 'content-type': PPTX_CONTENT_TYPE, 'x-source-filename': headerSafeFilename(file.name) },
    body: file
  });
  return response.json();
}

function renderSongs(songs) {
  results.setAttribute('aria-busy', 'false');
  results.replaceChildren();
  count.textContent = `${songs.length} song${songs.length === 1 ? '' : 's'}`;
  if (!songs.length) {
    const empty = document.createElement('div');
    empty.className = 'library-empty';
    const title = document.createElement('h2'); title.textContent = 'No matching songs';
    const copy = document.createElement('p'); copy.textContent = 'Try a shorter title or an alternative spelling.';
    empty.append(title, copy); results.append(empty); return;
  }
  const heading = document.createElement('div');
  heading.className = 'library-row library-row-heading';
  heading.innerHTML = '<span>Title</span><span>Version</span><span>Slides</span><span>Rights</span>';
  results.append(heading);
  songs.forEach(song => {
    const row = document.createElement('button');
    row.type = 'button';
    row.className = `library-row library-song-row${song.id === selectedId ? ' selected' : ''}`;
    row.addEventListener('click', () => selectSong(song, row));
    const titleCell = document.createElement('span');
    const title = document.createElement('strong'); title.textContent = song.title;
    const detail = document.createElement('small');
    detail.textContent = song.variant_label || song.author_owner || song.source_filename || 'Imported song';
    titleCell.append(title, detail);
    const version = document.createElement('span'); version.textContent = `v${song.current_version}`;
    const slides = document.createElement('span'); slides.textContent = song.slide_count;
    const rights = document.createElement('span');
    const badge = document.createElement('span');
    badge.className = `status-badge${song.rights_status === 'unknown' ? ' status-warning' : ' status-ok'}`;
    badge.textContent = friendlyRights(song.rights_status);
    rights.append(badge);
    row.append(titleCell, version, slides, rights);
    results.append(row);
  });
}

async function selectSong(song, row) {
  selectedId = song.id;
  results.querySelectorAll('.library-song-row').forEach(item => item.classList.toggle('selected', item === row));
  preview.setAttribute('aria-busy', 'true');
  preview.innerHTML = '<div class="skeleton-row"></div><div class="skeleton-row"></div><div class="skeleton-row"></div>';
  try {
    const response = await request(`/api/songs/${encodeURIComponent(song.id)}/preview`);
    renderPreview(await response.json());
  } catch (error) {
    preview.removeAttribute('aria-busy');
    preview.innerHTML = '<div class="library-empty"><h2>Preview unavailable</h2><p>The song remains in the catalogue. Try loading its preview again.</p></div>';
    showToast(error.message);
  }
}

function renderPreview(data) {
  preview.removeAttribute('aria-busy');
  preview.replaceChildren();
  const header = document.createElement('div'); header.className = 'preview-heading';
  const copy = document.createElement('div');
  const title = document.createElement('h2'); title.textContent = data.song.title;
  const detail = document.createElement('p');
  detail.textContent = data.song.current_version
    ? [data.song.variant_label, `Version ${data.song.current_version}`, `${data.song.slide_count} slides`].filter(Boolean).join(' · ')
    : [data.song.variant_label, 'Awaiting its PowerPoint'].filter(Boolean).join(' · ');
  copy.append(title, detail);
  const replace = document.createElement('button');
  replace.type = 'button';
  replace.className = 'button button-secondary';
  replace.textContent = data.song.current_version ? 'Upload new version' : 'Upload PowerPoint';
  replace.addEventListener('click', () => { replaceInput.value = ''; replaceInput.click(); });
  header.append(copy, replace); preview.append(header);

  const metadata = document.createElement('dl'); metadata.className = 'song-metadata';
  [['Rights', friendlyRights(data.song.rights_status)], ['Author / owner', data.song.author_owner || 'Needs review'], ['CCLI song number', data.song.ccli_song_number || 'Not recorded']].forEach(([label, value]) => {
    const term = document.createElement('dt'); term.textContent = label;
    const description = document.createElement('dd'); description.textContent = value;
    metadata.append(term, description);
  });
  preview.append(metadata);

  if (!data.slides.length) {
    const empty = document.createElement('p');
    empty.className = 'field-note';
    empty.textContent = 'No slides stored yet. Upload the PowerPoint to finish adding this song.';
    preview.append(empty);
    return;
  }

  const slideHeading = document.createElement('h3'); slideHeading.textContent = 'Extracted slide text'; preview.append(slideHeading);
  const slides = document.createElement('ol'); slides.className = 'preview-slides';
  data.slides.forEach((text, index) => {
    const item = document.createElement('li');
    const number = document.createElement('span'); number.textContent = index + 1;
    const body = document.createElement('p'); body.textContent = text || 'No extractable text';
    item.append(number, body); slides.append(item);
  });
  preview.append(slides);
}

async function loadSongs(query = '') {
  const currentRequest = ++requestNumber;
  results.setAttribute('aria-busy', 'true');
  try {
    const response = await request(`/api/songs?q=${encodeURIComponent(query)}`);
    const songs = await response.json();
    if (currentRequest === requestNumber) renderSongs(songs);
  } catch (error) {
    if (currentRequest !== requestNumber) return;
    results.setAttribute('aria-busy', 'false');
    results.innerHTML = '<div class="library-empty"><h2>Catalogue unavailable</h2><p>Check the R2 configuration and try again.</p></div>';
    count.textContent = 'Could not load catalogue';
    showToast(error.message);
  }
}

async function showPreview(songId) {
  selectedId = songId;
  const response = await request(`/api/songs/${encodeURIComponent(songId)}/preview`);
  renderPreview(await response.json());
}

function setAddError(message) {
  addError.textContent = message || '';
  addError.hidden = !message;
}

async function addSong(event) {
  event.preventDefault();
  const file = document.getElementById('song-file').files[0];
  const title = document.getElementById('song-title').value.trim();
  if (!title) return setAddError('Give the song a title.');
  const problem = describeFileProblem(file);
  if (problem) return setAddError(problem);

  setAddError('');
  addSubmit.disabled = true;
  let created = null;
  try {
    const response = await request('/api/songs', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({
        title,
        aliases: parseAliases(document.getElementById('song-aliases').value),
        variant_label: document.getElementById('song-variant').value.trim(),
        author_owner: document.getElementById('song-author').value.trim(),
        rights_status: document.getElementById('song-rights').value,
        ccli_song_number: document.getElementById('song-ccli').value.trim() || null,
        lyric_slides: [],
        credits: ''
      })
    });
    created = await response.json();
    await uploadPowerPoint(created.id, file);
  } catch (error) {
    // A failed upload leaves the catalogue entry behind, so say so rather than let the
    // song reappear as an empty row with no explanation.
    setAddError(created
      ? `${error.message} "${title}" was added to the catalogue without slides — select it and upload the PowerPoint again.`
      : error.message);
    addSubmit.disabled = false;
    if (created) await loadSongs(search.value.trim());
    return;
  }
  addSubmit.disabled = false;
  addForm.reset();
  addDialog.close();
  showToast(`"${title}" added to the library.`);
  selectedId = created.id;
  await loadSongs(search.value.trim());
  await showPreview(created.id).catch(() => {});
}

document.getElementById('add-song').addEventListener('click', () => {
  addForm.reset();
  setAddError('');
  addDialog.showModal();
});
addForm.addEventListener('submit', addSong);
addForm.querySelectorAll('[data-close]').forEach(button => {
  button.addEventListener('click', () => addDialog.close());
});

replaceInput.addEventListener('change', async () => {
  const file = replaceInput.files[0];
  if (!file || !selectedId) return;
  const problem = describeFileProblem(file);
  if (problem) return showToast(problem);
  try {
    const song = await uploadPowerPoint(selectedId, file);
    showToast(`"${song.title}" updated to version ${song.current_version}.`);
    await loadSongs(search.value.trim());
    await showPreview(song.id);
  } catch (error) {
    showToast(error.message);
  }
});

search.addEventListener('input', () => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => loadSongs(search.value.trim()), 260);
});
document.querySelector('.sign-out').addEventListener('click', async () => {
  await request('/api/logout', { method: 'POST' });
  location.assign('/login');
});
loadSongs();
