const csrf = document.querySelector('meta[name="csrf-token"]')?.content;
const ui = Object.fromEntries([
  'service-name', 'service-date', 'service-preset', 'service-heading', 'crumb-name',
  'component-list', 'component-count', 'slide-count', 'review-slides', 'editor-panel',
  'validation-list', 'readiness-score', 'readiness-bar', 'save-state', 'new-dialog',
  'preset-choices', 'review-dialog', 'review-title', 'full-review', 'toast'
].map(id => [id, document.getElementById(id)]));

let presets = [];
let service = null;
let lease = null;
let selectedId = null;
let saveTimer = null;
let saving = null;
let toastTimer = null;

const componentLabels = {
  welcome: 'Welcome', notices: 'Notices', call_to_worship: 'Call to Worship',
  cue_prayer: 'Prayer or cue', song: 'Song', psalm: 'Psalm', reading: 'Reading',
  teaching: 'Teaching', liturgy_block: 'Liturgy', custom_text_image: 'Custom slides'
};

async function request(url, options = {}) {
  const headers = new Headers(options.headers || {});
  if (csrf && !['GET', 'HEAD'].includes((options.method || 'GET').toUpperCase())) {
    headers.set('x-csrf-token', csrf);
  }
  if (lease?.token) headers.set('x-lease-token', lease.token);
  if (options.body && !(options.body instanceof FormData)) headers.set('content-type', 'application/json');
  const response = await fetch(url, { ...options, headers });
  if (!response.ok) {
    const data = await response.json().catch(() => ({}));
    throw new Error(data.error || `Request failed (${response.status})`);
  }
  return response;
}

function today() {
  const date = new Date();
  date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
  return date.toISOString().slice(0, 10);
}

function showToast(message) {
  ui.toast.textContent = message;
  ui.toast.classList.add('visible');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => ui.toast.classList.remove('visible'), 2800);
}

function setSaveState(kind, message) {
  ui['save-state'].className = `save-state ${kind || ''}`;
  ui['save-state'].innerHTML = '<span></span>';
  ui['save-state'].append(document.createTextNode(message));
}

function headingOf(component) {
  return component.heading || component.title || componentLabels[component.type] || 'Service item';
}

function detailOf(component) {
  switch (component.type) {
    case 'song': return component.song ? `Library version ${component.song.version}` : 'Song choice needed';
    case 'psalm': return component.reference || 'Passage needed';
    case 'reading': return component.reference || 'Reference needed';
    case 'call_to_worship': return component.reference || 'Reference needed';
    case 'notices': return `${component.rows.length} notice${component.rows.length === 1 ? '' : 's'}`;
    case 'teaching': return component.selection || 'Selection needed';
    default: return componentLabels[component.type] || '';
  }
}

function estimatedSlides(component) {
  if (component.type === 'notices') return Math.max(1, Math.ceil(component.rows.length / 5));
  if (component.type === 'song') return Math.max(1, component.lyric_slides.length);
  if (component.type === 'psalm') return Math.max(1, component.slide_breaks.length);
  if (component.type === 'custom_text_image') return Math.max(1, component.slides.length);
  if (component.type === 'liturgy_block') return Math.max(1, (component.text || '').split(/\n\s*\n/).filter(Boolean).length);
  return 1;
}

function renderAll() {
  if (!service) return;
  ui['service-name'].value = service.name;
  ui['service-date'].value = service.date;
  ui['service-preset'].value = service.preset;
  ui['service-heading'].textContent = service.name;
  ui['crumb-name'].textContent = service.name;
  renderOrder();
  renderEditor();
  renderValidation();
}

function renderOrder() {
  ui['component-list'].replaceChildren();
  service.components.forEach((component, index) => {
    const item = document.createElement('li');
    item.className = `component-item${component.id === selectedId ? ' selected' : ''}`;
    item.draggable = true;
    item.dataset.id = component.id;

    const handle = button('⠿', 'drag-handle', `Move ${headingOf(component)}`);
    handle.tabIndex = -1;
    const main = button('', 'component-main');
    const title = document.createElement('strong');
    title.textContent = headingOf(component);
    const detail = document.createElement('small');
    detail.textContent = detailOf(component);
    main.append(title, detail);
    main.addEventListener('click', () => selectComponent(component.id));
    main.addEventListener('keydown', event => {
      if (event.altKey && ['ArrowUp', 'ArrowDown'].includes(event.key)) {
        event.preventDefault();
        moveComponent(index, event.key === 'ArrowUp' ? -1 : 1);
      }
    });

    const actions = document.createElement('div');
    actions.className = 'component-actions';
    const duplicate = button('⧉', 'icon-button', `Duplicate ${headingOf(component)}`);
    duplicate.addEventListener('click', () => duplicateComponent(index));
    const remove = button('×', 'icon-button', `Remove ${headingOf(component)}`);
    remove.addEventListener('click', () => removeComponent(index));
    actions.append(duplicate, remove);
    item.append(handle, main, actions);

    item.addEventListener('dragstart', () => item.classList.add('dragging'));
    item.addEventListener('dragend', () => item.classList.remove('dragging'));
    item.addEventListener('dragover', event => {
      event.preventDefault();
      const dragging = ui['component-list'].querySelector('.dragging');
      if (dragging && dragging !== item) {
        const box = item.getBoundingClientRect();
        ui['component-list'].insertBefore(dragging, event.clientY < box.top + box.height / 2 ? item : item.nextSibling);
      }
    });
    ui['component-list'].append(item);
  });
  ui['component-list'].addEventListener('drop', syncDraggedOrder, { once: true });
  const slides = service.components.reduce((total, component) => total + estimatedSlides(component), 0);
  ui['component-count'].textContent = service.components.length;
  ui['slide-count'].textContent = slides;
  ui['review-slides'].textContent = `${slides} slide${slides === 1 ? '' : 's'}`;
}

function syncDraggedOrder() {
  const order = [...ui['component-list'].children].map(item => item.dataset.id);
  service.components.sort((a, b) => order.indexOf(a.id) - order.indexOf(b.id));
  changed();
}

function selectComponent(id) {
  selectedId = id;
  renderOrder();
  renderEditor();
}

function moveComponent(index, delta) {
  const destination = index + delta;
  if (destination < 0 || destination >= service.components.length) return;
  [service.components[index], service.components[destination]] = [service.components[destination], service.components[index]];
  changed();
  requestAnimationFrame(() => ui['component-list'].children[destination]?.querySelector('.component-main')?.focus());
}

function duplicateComponent(index) {
  const copy = structuredClone(service.components[index]);
  copy.id = `component-${Date.now().toString(36)}`;
  service.components.splice(index + 1, 0, copy);
  selectedId = copy.id;
  changed();
}

function removeComponent(index) {
  const [removed] = service.components.splice(index, 1);
  if (removed.id === selectedId) selectedId = service.components[index]?.id || service.components[index - 1]?.id || null;
  changed();
}

function renderEditor() {
  const component = service?.components.find(item => item.id === selectedId);
  if (!component) {
    ui['editor-panel'].innerHTML = '<div class="empty-editor"><div class="empty-icon" aria-hidden="true">✦</div><h2>Choose an item to edit</h2><p>Select an item in the service order. Its wording, song choice, reference, and slide settings will appear here.</p></div>';
    return;
  }
  const card = document.createElement('section');
  card.className = 'editor-card';
  const header = document.createElement('div');
  header.className = 'editor-heading';
  header.innerHTML = `<div><p class="section-label">${componentLabels[component.type]}</p><h2>${escapeHtml(headingOf(component))}</h2><p>${editorHelp(component.type)}</p></div>`;
  const fields = document.createElement('div');
  fields.className = 'editor-fields';
  card.append(header, fields);
  ui['editor-panel'].replaceChildren(card);

  if ('heading' in component) fields.append(textField('Heading', component.heading, value => component.heading = value));
  switch (component.type) {
    case 'welcome': break;
    case 'notices': renderNoticeFields(fields, component); break;
    case 'call_to_worship': renderCallFields(fields, component); break;
    case 'cue_prayer':
      fields.append(textField('Leader or reading cue', component.cue, value => component.cue = value, 'Optional'));
      fields.append(textArea('Text shown on slides', component.text, value => component.text = value, 'Leave blank for a heading-only cue.'));
      break;
    case 'song': renderSongFields(fields, component); break;
    case 'psalm': renderPsalmFields(fields, component); break;
    case 'reading':
      fields.append(textField('Bible reference', component.reference, value => component.reference = value, 'e.g. Luke 7:11–17'));
      fields.append(numberField('Bible page', component.bible_page, value => component.bible_page = value));
      break;
    case 'teaching': renderTeachingFields(fields, component); break;
    case 'liturgy_block':
      fields.append(textArea('Wording for this service', component.text, value => component.text = value, 'Leave blank to use the current staff-approved liturgy. Separate slides with a blank line.'));
      break;
    case 'custom_text_image': renderSlideBlocks(fields, component.slides, 'Slide', () => changed()); break;
  }
}

function editorHelp(type) {
  return {
    welcome: 'The opening slide uses the TWPC welcome layout.',
    notices: 'Complete rows stay together. Five rows are placed on each Notices slide.',
    call_to_worship: 'Fetch ESV wording, then edit it for this service if needed.',
    cue_prayer: 'Use a heading-only cue or add wording for the congregation.',
    song: 'Choose a library version or prepare explicit lyric-slide blocks.',
    psalm: 'The editor proposes readable groups. Staff can adjust every break.',
    reading: 'Readings show the heading, reference, and optional Bible page only.',
    teaching: 'Choose a public source, then edit the selected text for this service.',
    liturgy_block: 'Blank wording uses the current version. Completed services keep their pinned version.',
    custom_text_image: 'Create one or more TWPC-formatted slides. An image can be added in the library workflow.'
  }[type];
}

function renderNoticeFields(fields, component) {
  const list = document.createElement('div');
  list.className = 'notice-editor';
  component.rows.forEach((row, index) => {
    const wrapper = document.createElement('div');
    wrapper.className = 'notice-row';
    wrapper.append(
      bareInput(row.when, 'When', value => row.when = value),
      bareInput(row.title, 'Title', value => row.title = value),
      bareInput(row.details, 'Details', value => row.details = value)
    );
    const remove = button('×', 'icon-button', `Remove notice ${index + 1}`);
    remove.addEventListener('click', () => { component.rows.splice(index, 1); changed(); });
    wrapper.append(remove);
    list.append(wrapper);
  });
  fields.append(list);
  const add = button('Add notice', 'button button-secondary');
  add.addEventListener('click', () => { component.rows.push({ when: '', title: '', details: '', emphasis: false }); changed(); });
  fields.append(add);
}

function renderCallFields(fields, component) {
  const inline = document.createElement('div');
  inline.className = 'inline-action';
  inline.append(textField('ESV reference', component.reference, value => component.reference = value, 'e.g. Psalm 96:2'));
  const fetchButton = button('Fetch ESV text', 'button button-secondary');
  fetchButton.addEventListener('click', async () => {
    fetchButton.disabled = true;
    fetchButton.textContent = 'Fetching…';
    try {
      const response = await request(`/api/scripture?reference=${encodeURIComponent(component.reference)}`);
      const data = await response.json();
      component.text = data.text;
      component.external_source_failed = !data.ok;
      if (!data.ok) showToast(data.warning);
      changed();
    } catch (error) { showToast(error.message); }
    finally { fetchButton.disabled = false; fetchButton.textContent = 'Fetch ESV text'; }
  });
  inline.append(fetchButton);
  fields.append(inline);
  fields.append(textArea('Editable wording', component.text, value => component.text = value, 'Manual entry remains available if the ESV request fails.'));
}

function renderSongFields(fields, component) {
  fields.append(textField('Song title', component.title, value => component.title = value, 'Searchable library selection will pin an exact version.'));
  renderSlideBlocks(fields, component.lyric_slides, 'Lyric slide', () => changed());
  fields.append(textArea('Credits', component.credits, value => component.credits = value, 'Shown on the final lyric slide. CCLI licence 522221 is added automatically.'));
}

function renderPsalmFields(fields, component) {
  fields.append(textField('Psalm reference', component.reference, value => component.reference = value, 'e.g. Psalm 23:1–6'));
  renderSlideBlocks(fields, component.slide_breaks, 'Psalm slide', () => changed());
  const note = document.createElement('p');
  note.className = 'field-note';
  note.textContent = 'If no manual groups are entered, readable slide groups are proposed from the embedded Sing Psalms text during generation.';
  fields.append(note);
}

function renderTeachingFields(fields, component) {
  const source = document.createElement('label');
  source.textContent = 'Source';
  const select = document.createElement('select');
  [['westminster_shorter_catechism', 'Westminster Shorter Catechism'], ['heidelberg1891', 'Heidelberg Catechism, 1891'], ['westminster_confession_original_british', 'Westminster Confession, original British text']].forEach(([value, label]) => select.add(new Option(label, value)));
  select.value = component.source;
  select.addEventListener('change', () => { component.source = select.value; changed(false); });
  source.append(select);
  fields.append(source);
  fields.append(textField('Question or section', component.selection, value => component.selection = value));
  fields.append(textArea('Editable text', component.text, value => component.text = value, 'The wording saved here is pinned when the service is completed.'));
}

function renderSlideBlocks(fields, slides, label, rerender) {
  slides.forEach((text, index) => {
    const row = document.createElement('div');
    row.className = 'slide-block';
    const number = document.createElement('span');
    number.className = 'slide-number';
    number.textContent = index + 1;
    const area = document.createElement('textarea');
    area.value = text;
    area.setAttribute('aria-label', `${label} ${index + 1}`);
    area.addEventListener('input', () => { slides[index] = area.value; changed(false); });
    const remove = button('×', 'icon-button', `Remove ${label.toLowerCase()} ${index + 1}`);
    remove.addEventListener('click', () => { slides.splice(index, 1); rerender(); });
    row.append(number, area, remove);
    fields.append(row);
  });
  const add = button(`Add ${label.toLowerCase()}`, 'button button-secondary');
  add.addEventListener('click', () => { slides.push(''); rerender(); });
  fields.append(add);
}

function textField(labelText, value, setter, placeholder = '') {
  const label = document.createElement('label');
  label.textContent = labelText;
  label.append(bareInput(value, placeholder, setter));
  return label;
}

function numberField(labelText, value, setter) {
  const label = document.createElement('label');
  label.textContent = labelText;
  const input = document.createElement('input');
  input.type = 'number'; input.min = '1'; input.max = '2000'; input.value = value ?? '';
  input.addEventListener('input', () => { setter(input.value ? Number(input.value) : null); changed(false); });
  label.append(input); return label;
}

function bareInput(value, placeholder, setter) {
  const input = document.createElement('input');
  input.value = value || ''; input.placeholder = placeholder;
  input.addEventListener('input', () => { setter(input.value); changed(false); });
  return input;
}

function textArea(labelText, value, setter, note = '') {
  const fragment = document.createDocumentFragment();
  const label = document.createElement('label');
  label.textContent = labelText;
  const area = document.createElement('textarea');
  area.value = value || '';
  area.addEventListener('input', () => { setter(area.value); changed(false); });
  label.append(area); fragment.append(label);
  if (note) { const help = document.createElement('p'); help.className = 'field-note'; help.textContent = note; fragment.append(help); }
  return fragment;
}

function renderValidation() {
  const warnings = [];
  service.components.forEach(component => {
    if (['song'].includes(component.type) && !component.song && !component.lyric_slides.some(text => text.trim())) warnings.push({ title: headingOf(component), detail: 'Choose a library song or enter lyric slides.' });
    if (['psalm', 'reading', 'call_to_worship'].includes(component.type) && !component.reference.trim()) warnings.push({ title: headingOf(component), detail: 'Add a Bible or psalm reference.' });
    if (component.type === 'call_to_worship' && component.external_source_failed) warnings.push({ title: headingOf(component), detail: 'ESV fetch failed. Check the manually entered wording.' });
    if (component.type === 'teaching' && !component.selection.trim() && !component.text.trim()) warnings.push({ title: headingOf(component), detail: 'Choose a section and confirm its wording.' });
    if ((component.lyric_slides || component.slide_breaks || []).some(text => text.length > 620)) warnings.push({ title: headingOf(component), detail: 'A slide is unusually dense. Consider another break.' });
  });
  const possible = Math.max(1, service.components.length);
  const score = Math.max(0, Math.round((1 - warnings.length / possible) * 100));
  ui['readiness-score'].textContent = `${score}%`;
  ui['readiness-bar'].style.width = `${score}%`;
  ui['validation-list'].replaceChildren();
  if (!warnings.length) {
    const ok = document.createElement('div'); ok.className = 'validation-item ok'; ok.innerHTML = '<span>✓</span><div><strong>Ready to generate</strong><small>No missing service choices found.</small></div>'; ui['validation-list'].append(ok);
  } else warnings.slice(0, 7).forEach(warning => {
    const item = document.createElement('div'); item.className = 'validation-item warning';
    const icon = document.createElement('span'); icon.textContent = '!';
    const copy = document.createElement('div'); const title = document.createElement('strong'); title.textContent = warning.title; const detail = document.createElement('small'); detail.textContent = warning.detail; copy.append(title, detail); item.append(icon, copy); ui['validation-list'].append(item);
  });
}

function changed(rerender = true) {
  if (rerender) renderAll(); else { renderOrder(); renderValidation(); }
  setSaveState('saving', 'Changes not saved');
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveService, 900);
}

async function saveService() {
  if (!service || !lease || saving) return saving;
  clearTimeout(saveTimer);
  setSaveState('saving', 'Saving changes…');
  saving = request(`/api/services/${service.id}/autosave`, { method: 'PUT', body: JSON.stringify(service) })
    .then(response => response.json())
    .then(saved => { service = saved; setSaveState('', 'All changes saved'); return saved; })
    .catch(error => { setSaveState('error', 'Save failed'); showToast(error.message); throw error; })
    .finally(() => { saving = null; });
  return saving;
}

async function loadService(record) {
  service = record;
  selectedId = record.components[0]?.id || null;
  try {
    const response = await request(`/api/services/${record.id}/lock`, { method: 'POST' });
    lease = await response.json();
    setSaveState('', 'All changes saved');
  } catch (error) {
    lease = null;
    setSaveState('error', 'Read-only, another staff member is editing');
    showToast(error.message);
  }
  renderAll();
}

async function createService() {
  const chosen = document.querySelector('input[name="preset"]:checked')?.value || 'am';
  const preset = presets.find(item => item.id === chosen);
  const input = {
    name: preset.label,
    date: today(),
    preset: chosen
  };
  const response = await request('/api/services', { method: 'POST', body: JSON.stringify(input) });
  await loadService(await response.json());
  ui['new-dialog'].close();
}

function renderPresetChoices() {
  ui['service-preset'].replaceChildren();
  ui['preset-choices'].replaceChildren();
  presets.forEach((preset, index) => {
    ui['service-preset'].add(new Option(preset.label, preset.id));
    const choice = document.createElement('label'); choice.className = 'preset-choice';
    const radio = document.createElement('input'); radio.type = 'radio'; radio.name = 'preset'; radio.value = preset.id; radio.checked = index === 0;
    const text = document.createElement('span'); const title = document.createElement('strong'); title.textContent = preset.label; const detail = document.createElement('small'); detail.textContent = `${preset.components.length} items · ${preset.components.reduce((sum, item) => sum + estimatedSlides(item), 0)} starting slides`; text.append(title, detail); choice.append(radio, text); ui['preset-choices'].append(choice);
  });
}

function openReview() {
  ui['review-title'].textContent = service.name;
  ui['full-review'].replaceChildren();
  service.components.forEach((component, index) => {
    const row = document.createElement('div'); row.className = 'full-review-row';
    const number = document.createElement('span'); number.textContent = index + 1;
    const text = document.createElement('div'); const title = document.createElement('strong'); title.textContent = headingOf(component); const detail = document.createElement('small'); detail.textContent = detailOf(component); text.append(title, detail);
    const slides = document.createElement('span'); slides.textContent = `${estimatedSlides(component)} slide${estimatedSlides(component) === 1 ? '' : 's'}`; row.append(number, text, slides); ui['full-review'].append(row);
  });
  ui['review-dialog'].showModal();
}

async function generate() {
  try {
    await saveService();
    setSaveState('saving', 'Generating PowerPoint…');
    const response = await request(`/api/services/${service.id}/generate`, { method: 'POST' });
    const blob = await response.blob();
    const disposition = response.headers.get('content-disposition') || '';
    const filename = disposition.match(/filename="([^"]+)"/)?.[1] || `service-${service.date}.pptx`;
    const link = document.createElement('a'); link.href = URL.createObjectURL(blob); link.download = filename; link.click(); setTimeout(() => URL.revokeObjectURL(link.href), 1000);
    service.status = 'completed'; lease = null; setSaveState('', 'PowerPoint generated'); showToast('PowerPoint generated and saved to service history.'); ui['review-dialog'].close();
  } catch (error) { setSaveState('error', 'Generation failed'); showToast(error.message); }
}

function button(text, className, label) {
  const element = document.createElement('button'); element.type = 'button'; element.className = className; element.textContent = text; if (label) element.setAttribute('aria-label', label); return element;
}

function escapeHtml(value) {
  return String(value).replace(/[&<>'"]/g, character => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' })[character]);
}

document.getElementById('new-service').addEventListener('click', () => ui['new-dialog'].showModal());
document.getElementById('create-service').addEventListener('click', event => { event.preventDefault(); createService().catch(error => showToast(error.message)); });
document.getElementById('review-service').addEventListener('click', openReview);
document.getElementById('generate-service').addEventListener('click', generate);
document.getElementById('review-generate').addEventListener('click', event => { event.preventDefault(); generate(); });
document.getElementById('add-component').addEventListener('click', () => {
  const component = { type: 'custom_text_image', id: `component-${Date.now().toString(36)}`, heading: 'Custom slide', slides: [''], image: null };
  service.components.push(component); selectedId = component.id; changed();
});
document.getElementById('sign-out').addEventListener('click', async () => { await request('/api/logout', { method: 'POST' }); location.assign('/login'); });
ui['service-name'].addEventListener('input', () => { service.name = ui['service-name'].value; ui['service-heading'].textContent = service.name; ui['crumb-name'].textContent = service.name; changed(false); });
ui['service-date'].addEventListener('input', () => { service.date = ui['service-date'].value; changed(false); });
ui['service-preset'].addEventListener('change', () => {
  const preset = presets.find(item => item.id === ui['service-preset'].value);
  if (!preset || !confirm('Replace the current order with this preset?')) { ui['service-preset'].value = service.preset; return; }
  service.preset = preset.id; service.components = structuredClone(preset.components); selectedId = service.components[0]?.id; changed();
});

window.addEventListener('beforeunload', event => { if (saveTimer || saving) { event.preventDefault(); event.returnValue = ''; } });

Promise.all([
  request('/api/presets').then(response => response.json()),
  request('/api/services').then(response => response.json())
]).then(([loadedPresets, services]) => {
  presets = loadedPresets;
  renderPresetChoices();
  const current = services.find(item => item.status === 'draft') || services.find(item => item.status !== 'archived');
  if (current) loadService(current); else ui['new-dialog'].showModal();
}).catch(error => showToast(error.message));
