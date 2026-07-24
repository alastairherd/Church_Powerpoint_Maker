const csrf = document.querySelector('meta[name="csrf-token"]')?.content;
const search = document.getElementById('song-search');
const results = document.getElementById('library-results');
const preview = document.getElementById('song-preview');
const count = document.getElementById('song-count');
const toast = document.getElementById('toast');
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
  detail.textContent = [data.song.variant_label, `Version ${data.song.current_version}`, `${data.song.slide_count} slides`].filter(Boolean).join(' · ');
  copy.append(title, detail); header.append(copy); preview.append(header);

  const metadata = document.createElement('dl'); metadata.className = 'song-metadata';
  [['Rights', friendlyRights(data.song.rights_status)], ['Author / owner', data.song.author_owner || 'Needs review'], ['CCLI song number', data.song.ccli_song_number || 'Not recorded']].forEach(([label, value]) => {
    const term = document.createElement('dt'); term.textContent = label;
    const description = document.createElement('dd'); description.textContent = value;
    metadata.append(term, description);
  });
  preview.append(metadata);

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

search.addEventListener('input', () => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => loadSongs(search.value.trim()), 260);
});
document.querySelector('.sign-out').addEventListener('click', async () => {
  await request('/api/logout', { method: 'POST' });
  location.assign('/login');
});
loadSongs();
