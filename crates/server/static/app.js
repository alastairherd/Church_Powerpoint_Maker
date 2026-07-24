import { createEditorController } from './editor-controller.js';

const componentLabels = {
  welcome: 'Welcome', notices: 'Notices', call_to_worship: 'Call to Worship',
  cue_prayer: 'Prayer or cue', song: 'Song', psalm: 'Psalm', reading: 'Reading',
  teaching: 'Teaching', liturgy_block: 'Liturgy', custom_text_image: 'Custom slides'
};

const DEFAULT_TIMERS = {
  setTimeout: globalThis.setTimeout.bind(globalThis),
  clearTimeout: globalThis.clearTimeout.bind(globalThis),
  setInterval: globalThis.setInterval.bind(globalThis),
  clearInterval: globalThis.clearInterval.bind(globalThis),
};

export function createEditorApp({
  document: doc = globalThis.document,
  request: injectedRequest,
  fetchImpl = globalThis.fetch,
  timers = DEFAULT_TIMERS,
  confirmImpl = globalThis.confirm?.bind(globalThis) || (() => true),
  locationImpl = doc.defaultView?.location,
} = {}) {
  const ui = Object.fromEntries([
    'service-name', 'service-date', 'service-preset', 'service-heading', 'crumb-name',
    'component-list', 'component-count', 'slide-count', 'review-slides', 'editor-panel',
    'validation-list', 'readiness-score', 'readiness-bar', 'save-state', 'save-help',
    'save-now',
    'new-dialog', 'preset-choices', 'review-dialog', 'review-title', 'full-review', 'toast',
    'new-service', 'create-service', 'review-service', 'generate-service', 'review-generate',
    'add-component', 'sign-out',
  ].map(id => [id, doc.getElementById(id)]));
  let presets = [];
  let controller;
  let toastTimer = null;
  let booted = false;
  let generationPromise = null;
  const generationLabels = new Map();

  async function request(url, options = {}) {
    if (injectedRequest) return injectedRequest(url, options);
    const headers = new Headers(options.headers || {});
    const csrf = doc.querySelector('meta[name="csrf-token"]')?.content;
    if (csrf && !['GET', 'HEAD'].includes((options.method || 'GET').toUpperCase())) headers.set('x-csrf-token', csrf);
    if (options.body && !(options.body instanceof FormData)) headers.set('content-type', 'application/json');
    const response = await fetchImpl(url, { ...options, headers });
    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      const error = new Error(data.error || `Request failed (${response.status})`);
      error.status = response.status;
      error.body = data;
      throw error;
    }
    return response;
  }

  function today() {
    const date = new Date();
    date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
    return date.toISOString().slice(0, 10);
  }

  function showToast(message) {
    if (!ui.toast) return;
    ui.toast.textContent = message;
    ui.toast.classList.add('visible');
    if (toastTimer !== null) timers.clearTimeout(toastTimer);
    toastTimer = timers.setTimeout(() => ui.toast.classList.remove('visible'), 2800);
  }

  function setSaveState(kind, message) {
    if (!ui['save-state']) return;
    const stateClass = { Unsaved: 'saving', Saving: 'saving', Saved: '', Failed: 'error' }[kind] || '';
    ui['save-state'].className = `save-state ${stateClass}`;
    ui['save-state'].replaceChildren(doc.createElement('span'), doc.createTextNode(kind));
  }

  function setSaveHelp(message) {
    if (!ui['save-help']) return;
    ui['save-help'].textContent = message || '';
    ui['save-help'].hidden = !message;
  }

  function headingOf(component) {
    return component.heading || component.title || componentLabels[component.type] || 'Service item';
  }

  function detailOf(component) {
    switch (component.type) {
      case 'song': return component.song ? `Library v${component.song.version} · ${component.song.slide_count || 1} slides` : 'Song choice needed';
      case 'psalm': return component.reference || 'Passage needed';
      case 'reading': return component.reference || 'Reference needed';
      case 'call_to_worship': return component.reference || 'Reference needed';
      case 'notices': return `${(component.rows || []).length} notice${(component.rows || []).length === 1 ? '' : 's'}`;
      case 'teaching': return component.selection || 'Selection needed';
      default: return componentLabels[component.type] || '';
    }
  }

  function estimatedSlides(component) {
    if (component.type === 'notices') return Math.max(1, Math.ceil((component.rows || []).length / 5));
    if (component.type === 'song') return Math.max(1, component.song?.slide_count || (component.lyric_slides || []).length);
    if (component.type === 'psalm') return Math.max(1, (component.slide_breaks || []).length);
    if (component.type === 'custom_text_image') return Math.max(1, (component.slides || []).length);
    if (component.type === 'liturgy_block') return Math.max(1, (component.text || '').split(/\n\s*\n/).filter(Boolean).length);
    return 1;
  }

  function updateHeading() {
    const name = controller.getService()?.name || '';
    if (ui['service-heading']) ui['service-heading'].textContent = name;
    if (ui['crumb-name']) ui['crumb-name'].textContent = name;
  }

  function updateCounts() {
    const components = controller.getService()?.components || [];
    const slides = components.reduce((total, component) => total + estimatedSlides(component), 0);
    if (ui['component-count']) ui['component-count'].textContent = components.length;
    if (ui['slide-count']) ui['slide-count'].textContent = slides;
    if (ui['review-slides']) ui['review-slides'].textContent = `${slides} slide${slides === 1 ? '' : 's'}`;
  }

  function renderServiceFields() {
    const service = controller.getService();
    if (!service) return;
    if (ui['service-name']) ui['service-name'].value = service.name || '';
    if (ui['service-date']) ui['service-date'].value = service.date || '';
    if (ui['service-preset']) ui['service-preset'].value = service.preset || '';
    updateHeading();
  }

  function renderAll() {
    if (!controller.getService()) return;
    renderServiceFields();
    renderOrder();
    renderEditor();
    updateCounts();
    renderValidation();
  }

  function renderOrder() {
    const service = controller.getService();
    if (!service || !ui['component-list']) return;
    const selectedId = controller.getState().selectedId;
    ui['component-list'].replaceChildren();
    service.components.forEach((component, index) => {
      const item = doc.createElement('li');
      item.className = `component-item component-item--${component.type}${component.id === selectedId ? ' selected' : ''}`;
      item.draggable = true;
      item.dataset.id = component.id;
      item.dataset.type = component.type;

      const handle = button('⠿', 'drag-handle', `Move ${headingOf(component)}`);
      handle.tabIndex = -1;
      const main = button('', 'component-main');
      const title = doc.createElement('strong');
      title.textContent = headingOf(component);
      const detail = doc.createElement('small');
      detail.textContent = detailOf(component);
      main.append(title, detail);
      main.addEventListener('click', () => selectComponent(component.id));
      main.addEventListener('keydown', event => {
        if (event.altKey && ['ArrowUp', 'ArrowDown'].includes(event.key)) {
          event.preventDefault();
          moveComponent(index, event.key === 'ArrowUp' ? -1 : 1);
        }
      });

      const actions = doc.createElement('div');
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
  }

  function updateOrderItem(componentId) {
    const component = controller.findComponent(componentId);
    const item = [...(ui['component-list']?.children || [])].find(child => child.dataset.id === componentId);
    if (!component || !item) return;
    item.querySelector('strong').textContent = headingOf(component);
    item.querySelector('small').textContent = detailOf(component);
  }

  function syncDraggedOrder() {
    const order = [...ui['component-list'].children].map(item => item.dataset.id);
    controller.updateService(service => {
      service.components.sort((a, b) => order.indexOf(a.id) - order.indexOf(b.id));
    }, 'structural');
  }

  function selectComponent(id) {
    controller.selectComponent(id);
    renderOrder();
    renderEditor();
    const panel = ui['editor-panel'];
    if (panel) {
      panel.scrollTop = 0;
      panel.scrollIntoView?.({ block: 'nearest' });
    }
  }

  function moveComponent(index, delta) {
    const service = controller.getService();
    const destination = index + delta;
    if (!service || destination < 0 || destination >= service.components.length) return;
    controller.updateService(current => {
      [current.components[index], current.components[destination]] = [current.components[destination], current.components[index]];
    }, 'structural');
    const frame = doc.defaultView?.requestAnimationFrame;
    if (frame) frame.call(doc.defaultView, () => ui['component-list'].children[destination]?.querySelector('.component-main')?.focus());
  }

  function duplicateComponent(index) {
    const copy = structuredClone(controller.getService().components[index]);
    copy.id = `component-${Date.now().toString(36)}`;
    controller.updateService(service => {
      service.components.splice(index + 1, 0, copy);
    }, 'structural');
    controller.selectComponent(copy.id);
    renderAll();
  }

  function removeComponent(index) {
    const service = controller.getService();
    const removed = service.components[index];
    controller.updateService(current => { current.components.splice(index, 1); }, 'structural');
    if (removed.id === controller.getState().selectedId) {
      controller.selectComponent(service.components[index]?.id || service.components[index - 1]?.id || null);
      renderAll();
    }
  }

  function renderEditor() {
    const service = controller.getService();
    if (!service || !ui['editor-panel']) return;
    const component = controller.findComponent(controller.getState().selectedId);
    if (!component) {
      const empty = doc.createElement('div'); empty.className = 'empty-editor';
      const icon = doc.createElement('div'); icon.className = 'empty-icon'; icon.setAttribute('aria-hidden', 'true'); icon.textContent = '✦';
      const title = doc.createElement('h2'); title.textContent = 'Choose an item to edit';
      const help = doc.createElement('p'); help.textContent = 'Select an item in the service order. Its wording, song choice, reference, and slide settings will appear here.';
      empty.append(icon, title, help); ui['editor-panel'].replaceChildren(empty); return;
    }
    const card = doc.createElement('section'); card.className = 'editor-card';
    const header = doc.createElement('div'); header.className = 'editor-heading';
    const copy = doc.createElement('div');
    const label = doc.createElement('p'); label.className = 'section-label'; label.textContent = componentLabels[component.type];
    const title = doc.createElement('h2'); title.textContent = headingOf(component);
    const help = doc.createElement('p'); help.textContent = editorHelp(component.type);
    copy.append(label, title, help); header.append(copy);
    const fields = doc.createElement('div'); fields.className = 'editor-fields'; card.append(header, fields);
    ui['editor-panel'].replaceChildren(card);

    if ('heading' in component) fields.append(textField('Heading', component.id, 'heading', component.heading, '', 'heading'));
    switch (component.type) {
      case 'welcome': break;
      case 'notices': renderNoticeFields(fields, component); break;
      case 'call_to_worship': renderCallFields(fields, component); break;
      case 'cue_prayer':
        fields.append(textField('Leader or reading cue', component.id, 'cue', component.cue, 'Optional'));
        fields.append(textArea('Text shown on slides', component.id, 'text', component.text, 'Leave blank for a heading-only cue.'));
        break;
      case 'song': renderSongFields(fields, component); break;
      case 'psalm': renderPsalmFields(fields, component); break;
      case 'reading':
        fields.append(textField('Bible reference', component.id, 'reference', component.reference, 'e.g. Luke 7:11–17', 'summary'));
        fields.append(numberField('Bible page', component.id, 'bible_page', component.bible_page));
        break;
      case 'teaching': renderTeachingFields(fields, component); break;
      case 'liturgy_block':
        fields.append(textArea('Wording for this service', component.id, 'text', component.text, 'Leave blank to use the current staff-approved liturgy. Separate slides with a blank line.'));
        break;
      case 'custom_text_image': renderSlideBlocks(fields, component.id, 'slides', 'Slide'); break;
    }
    if (component.type === 'psalm') renderLoaderState(component.id, 'psalm');
    if (component.type === 'call_to_worship') renderLoaderState(component.id, 'esv');
    if (component.type === 'teaching') renderLoaderState(component.id, 'teaching');
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
    const list = doc.createElement('div'); list.className = 'notice-editor';
    (component.rows || []).forEach((row, index) => {
      const wrapper = doc.createElement('div'); wrapper.className = 'notice-row';
      wrapper.dataset.index = index;
      const handle = button('⠿', 'drag-handle', `Move notice ${index + 1}`);
      handle.tabIndex = -1;
      handle.addEventListener('pointerdown', () => { wrapper.draggable = true; });
      wrapper.append(
        handle,
        bareInput(component.id, `rows.${index}.when`, row.when, 'When', 'field', (current, value) => { current.rows[index].when = value; }),
        bareInput(component.id, `rows.${index}.title`, row.title, 'Title', 'field', (current, value) => { current.rows[index].title = value; }),
        bareInput(component.id, `rows.${index}.details`, row.details, 'Details', 'field', (current, value) => { current.rows[index].details = value; })
      );
      const remove = button('×', 'icon-button', `Remove notice ${index + 1}`);
      remove.addEventListener('click', () => controller.updateComponent(component.id, current => { current.rows.splice(index, 1); }, 'structural'));
      wrapper.append(remove); list.append(wrapper);

      wrapper.addEventListener('dragstart', () => wrapper.classList.add('dragging'));
      wrapper.addEventListener('dragend', () => { wrapper.classList.remove('dragging'); wrapper.draggable = false; });
      wrapper.addEventListener('dragover', event => {
        event.preventDefault();
        const dragging = list.querySelector('.dragging');
        if (dragging && dragging !== wrapper) {
          const box = wrapper.getBoundingClientRect();
          list.insertBefore(dragging, event.clientY < box.top + box.height / 2 ? wrapper : wrapper.nextSibling);
        }
      });
    });
    list.addEventListener('drop', () => {
      const order = [...list.children].map(child => Number(child.dataset.index));
      controller.updateComponent(component.id, current => {
        current.rows = order.map(index => current.rows[index]);
      }, 'structural');
    }, { once: true });
    fields.append(list);
    const add = button('Add notice', 'button button-secondary');
    add.addEventListener('click', () => controller.updateComponent(component.id, current => { current.rows.push({ when: '', title: '', details: '', emphasis: false }); }, 'structural'));
    fields.append(add);
  }

  function renderCallFields(fields, component) {
    const inline = doc.createElement('div'); inline.className = 'inline-action';
    inline.append(textField('ESV reference', component.id, 'reference', component.reference, 'e.g. Psalm 96:2', 'summary'));
    const fetchButton = button('Fetch ESV text', 'button button-secondary');
    const componentId = component.id;
    fetchButton.dataset.loaderKind = 'esv';
    fetchButton.dataset.componentId = componentId;
    const error = loaderErrorNode('esv', componentId);
    fetchButton.setAttribute('aria-describedby', error.id);
    fetchButton.addEventListener('click', () => {
      const current = controller.findComponent(componentId);
      const reference = current?.reference || '';
      if (!/\d/.test(reference)) { showToast('Enter a complete Bible reference, for example Psalm 23:1–6.'); return; }
      void controller.loadEsv(componentId, reference);
    });
    inline.append(fetchButton, error); fields.append(inline);
    fields.append(textArea('Editable wording', component.id, 'text', component.text, 'Manual entry remains available if the ESV request fails.'));
  }

  function renderSongFields(fields, component) {
    const picker = doc.createElement('section'); picker.className = 'song-picker';
    const pickerHeading = doc.createElement('div'); pickerHeading.className = 'song-picker-heading';
    const pickerTitle = doc.createElement('h3'); pickerTitle.textContent = 'Choose from the song library';
    const libraryLink = doc.createElement('a'); libraryLink.href = '/library'; libraryLink.textContent = 'Browse full library';
    libraryLink.addEventListener('click', guardedNavigation);
    pickerHeading.append(pickerTitle, libraryLink); picker.append(pickerHeading);

    if (component.song) {
      const selected = doc.createElement('div'); selected.className = 'selected-song';
      const copy = doc.createElement('div');
      const title = doc.createElement('strong'); title.textContent = component.title;
      const detail = doc.createElement('small'); detail.textContent = `Version ${component.song.version} · ${component.song.slide_count || 1} slides`;
      copy.append(title, detail);
      const clear = button('Change song', 'button button-secondary');
      clear.addEventListener('click', () => controller.updateComponent(component.id, current => { current.song = null; current.title = ''; }, 'structural'));
      selected.append(copy, clear); picker.append(selected);
    }

    const searchLabel = doc.createElement('label'); searchLabel.textContent = component.song ? 'Search for a replacement' : 'Search library';
    const searchInput = doc.createElement('input'); searchInput.type = 'search'; searchInput.autocomplete = 'off'; searchInput.placeholder = 'Start typing a song title';
    searchInput.setAttribute('role', 'combobox'); searchInput.setAttribute('aria-expanded', 'false');
    const results = doc.createElement('div'); results.className = 'song-picker-results'; results.setAttribute('role', 'listbox');
    const status = doc.createElement('p'); status.className = 'field-note song-search-status'; status.textContent = 'Type a title, or focus the field to browse all active songs.';
    searchLabel.append(searchInput); picker.append(searchLabel, results, status); fields.append(picker);

    let timer = null; let sequence = 0;
    async function searchSongs() {
      const currentSequence = ++sequence;
      status.textContent = 'Searching library…'; searchInput.setAttribute('aria-expanded', 'true');
      try {
        const response = await request(`/api/songs?q=${encodeURIComponent(searchInput.value.trim())}`);
        const songs = await response.json();
        if (currentSequence !== sequence) return;
        results.replaceChildren();
        songs.slice(0, 14).forEach(song => {
          const choice = button('', 'song-choice'); choice.setAttribute('role', 'option');
          const copy = doc.createElement('span');
          const title = doc.createElement('strong'); title.textContent = song.title;
          const detail = doc.createElement('small'); detail.textContent = [song.variant_label, `v${song.current_version}`, `${song.slide_count} slides`].filter(Boolean).join(' · ');
          copy.append(title, detail); const use = doc.createElement('span'); use.className = 'song-choice-action'; use.textContent = 'Select'; choice.append(copy, use);
          choice.addEventListener('click', () => controller.updateComponent(component.id, current => {
            current.title = song.title;
            current.song = { entity_id: song.id, version: song.current_version, slide_count: song.slide_count };
            current.lyric_slides = []; current.credits = '';
          }, 'structural'));
          results.append(choice);
        });
        status.textContent = songs.length ? `${songs.length} match${songs.length === 1 ? '' : 'es'}${songs.length > 14 ? ', showing the first 14' : ''}.` : 'No matching songs. Try a shorter title or an alternative spelling.';
        searchInput.setAttribute('aria-expanded', songs.length ? 'true' : 'false');
      } catch (error) {
        results.replaceChildren(); status.textContent = 'The song library could not be loaded.'; searchInput.setAttribute('aria-expanded', 'false'); showToast(error.message);
      }
    }
    searchInput.addEventListener('focus', () => { if (!results.children.length) searchSongs(); });
    searchInput.addEventListener('input', () => { if (timer !== null) timers.clearTimeout(timer); timer = timers.setTimeout(searchSongs, 260); });

    if (!component.song) {
      const divider = doc.createElement('div'); divider.className = 'editor-divider'; divider.textContent = 'Or enter custom lyrics'; fields.append(divider);
      fields.append(textField('Song title', component.id, 'title', component.title, 'Required for custom lyric slides.', 'heading'));
      renderSlideBlocks(fields, component.id, 'lyric_slides', 'Lyric slide');
      fields.append(textArea('Credits', component.id, 'credits', component.credits, 'Shown on the final lyric slide. The global CCLI licence number is added automatically.'));
    }
  }

  function renderPsalmFields(fields, component) {
    const inline = doc.createElement('div'); inline.className = 'inline-action';
    inline.append(textField('Sing Psalms reference', component.id, 'reference', component.reference, 'e.g. Psalm 23:1–6', 'summary'));
    const verseNumbers = doc.createElement('label');
    verseNumbers.className = 'checkbox-field';
    verseNumbers.textContent = 'Show verse numbers';
    const verseNumbersInput = doc.createElement('input');
    verseNumbersInput.type = 'checkbox';
    verseNumbersInput.checked = component.show_verse_numbers !== false;
    verseNumbersInput.dataset.componentId = component.id;
    verseNumbersInput.dataset.field = 'show_verse_numbers';
    verseNumbersInput.addEventListener('change', () => setComponentField(
      component.id,
      'show_verse_numbers',
      verseNumbersInput.checked,
    ));
    verseNumbers.append(verseNumbersInput);
    fields.append(verseNumbers);
    const loadButton = button('Load Psalm text', 'button button-secondary');
    const componentId = component.id;
    loadButton.dataset.loaderKind = 'psalm';
    loadButton.dataset.componentId = componentId;
    const error = loaderErrorNode('psalm', componentId);
    loadButton.setAttribute('aria-describedby', error.id);
    loadButton.addEventListener('click', () => {
      const current = controller.findComponent(componentId);
      const reference = current?.reference || '';
      if (!/\d/.test(reference)) { showToast('Enter a complete Psalm reference, for example Psalm 23:1–6.'); return; }
      void controller.loadPsalm(componentId, reference);
    });
    inline.append(loadButton, error); fields.append(inline);
    renderSlideBlocks(fields, component.id, 'slide_breaks', 'Psalm slide');
    const note = doc.createElement('p'); note.className = 'field-note'; note.textContent = 'Loading proposes readable groups from the embedded Sing Psalms text. You can then edit every break before generation.'; fields.append(note);
  }

  function renderTeachingFields(fields, component) {
    const source = doc.createElement('label'); source.textContent = 'Source';
    const select = doc.createElement('select'); select.dataset.componentId = component.id; select.dataset.field = 'source';
    [['westminster_shorter_catechism', 'Westminster Shorter Catechism'], ['heidelberg1891', 'Heidelberg Catechism, 1891'], ['westminster_confession_original_british', 'Westminster Confession, original British text']].forEach(([value, label]) => {
      const option = doc.createElement('option'); option.value = value; option.textContent = label; select.append(option);
    });
    select.value = component.source;
    select.addEventListener('change', () => {
      setComponentField(component.id, 'source', select.value, 'summary');
      renderEditor();
    });
    source.append(select); fields.append(source);
    const teachingCopy = {
      westminster_shorter_catechism: { load: 'Load WSC question', retry: 'Retry WSC question', placeholder: 'e.g. 1, Q1, or Q. 1', prompt: 'Enter a catechism question, for example Q. 1.', note: 'Load retrieves the embedded WSC question and answer into editable text.' },
      heidelberg1891: { load: 'Load Heidelberg question', retry: 'Retry Heidelberg question', placeholder: 'e.g. 1, Q1, or Q. 1', prompt: 'Enter a catechism question, for example Q. 1.', note: 'Load retrieves the embedded Heidelberg Catechism question and answer into editable text.' },
      westminster_confession_original_british: { load: 'Load WCF section', retry: 'Retry WCF section', placeholder: 'e.g. 21 or 21.8', prompt: 'Enter a confession chapter or section, for example 21.8.', note: 'Load retrieves the embedded Westminster Confession chapter or section into editable text.' },
    }[component.source] || { load: 'Load text', retry: 'Retry', placeholder: '', prompt: 'Enter a selection first.', note: '' };
    fields.append(textField('Question or section', component.id, 'selection', component.selection, teachingCopy.placeholder, 'summary'));
    const inline = doc.createElement('div'); inline.className = 'inline-action';
    const loadButton = button(teachingCopy.load, 'button button-secondary');
    loadButton.dataset.loaderKind = 'teaching';
    loadButton.dataset.componentId = component.id;
    loadButton.dataset.idleLabel = teachingCopy.load;
    loadButton.dataset.retryLabel = teachingCopy.retry;
    const error = loaderErrorNode('teaching', component.id);
    loadButton.setAttribute('aria-describedby', error.id);
    loadButton.addEventListener('click', () => {
      const current = controller.findComponent(component.id);
      if (!current?.selection.trim()) {
        showToast(teachingCopy.prompt);
        return;
      }
      void controller.loadTeaching(component.id, current.source, current.selection);
    });
    inline.append(loadButton, error); fields.append(inline);
    const note = doc.createElement('p'); note.className = 'field-note';
    note.textContent = teachingCopy.note;
    fields.append(note);
    fields.append(textArea('Editable text', component.id, 'text', component.text, 'The wording saved here is pinned when the service is completed.'));
  }

  function renderSlideBlocks(fields, componentId, fieldName, label) {
    const component = controller.findComponent(componentId);
    const slides = component?.[fieldName] || [];
    slides.forEach((text, index) => {
      const row = doc.createElement('div'); row.className = 'slide-block';
      const number = doc.createElement('span'); number.className = 'slide-number'; number.textContent = index + 1;
      const area = doc.createElement('textarea'); area.value = text; area.dataset.componentId = componentId; area.dataset.field = `${fieldName}.${index}`; area.setAttribute('aria-label', `${label} ${index + 1}`);
      area.addEventListener('input', () => setComponentField(componentId, `${fieldName}.${index}`, area.value, 'field', (current, value) => { current[fieldName][index] = value; }));
      const remove = button('×', 'icon-button', `Remove ${label.toLowerCase()} ${index + 1}`);
      remove.addEventListener('click', () => controller.updateComponent(componentId, current => { current[fieldName].splice(index, 1); }, 'structural'));
      row.append(number, area, remove); fields.append(row);
    });
    const add = button(`Add ${label.toLowerCase()}`, 'button button-secondary');
    add.addEventListener('click', () => controller.updateComponent(componentId, current => { current[fieldName].push(''); }, 'structural'));
    fields.append(add);
  }

  function setComponentField(componentId, field, value, scope = 'field', mutate = null) {
    controller.updateComponent(componentId, component => {
      if (mutate) mutate(component, value); else component[field] = value;
    }, scope);
  }

  function textField(labelText, componentId, fieldName, value, placeholder = '', scope = 'field') {
    const label = doc.createElement('label'); label.textContent = labelText;
    label.append(bareInput(componentId, fieldName, value, placeholder, scope)); return label;
  }

  function numberField(labelText, componentId, fieldName, value, scope = 'field') {
    const label = doc.createElement('label');
    const input = doc.createElement('input'); input.type = 'number'; input.min = '1'; input.max = '2000'; input.value = value ?? '';
    input.dataset.componentId = componentId; input.dataset.field = fieldName;
    input.addEventListener('input', () => setComponentField(componentId, fieldName, input.value ? Number(input.value) : null, scope));
    label.textContent = labelText; label.append(input); return label;
  }

  function bareInput(componentId, fieldName, value, placeholder, scope = 'field', mutate = null) {
    const input = doc.createElement('input'); input.value = value || ''; input.placeholder = placeholder;
    input.dataset.componentId = componentId; input.dataset.field = fieldName;
    input.addEventListener('input', () => setComponentField(componentId, fieldName, input.value, scope, mutate)); return input;
  }

  function textArea(labelText, componentId, fieldName, value, note = '', scope = 'field') {
    const fragment = doc.createDocumentFragment();
    const label = doc.createElement('label'); label.textContent = labelText;
    const area = doc.createElement('textarea'); area.value = value || ''; area.dataset.componentId = componentId; area.dataset.field = fieldName;
    area.addEventListener('input', () => setComponentField(componentId, fieldName, area.value, scope)); label.append(area); fragment.append(label);
    if (note) { const help = doc.createElement('p'); help.className = 'field-note'; help.textContent = note; fragment.append(help); }
    return fragment;
  }

  function renderValidation() {
    const service = controller.getService();
    if (!service || !ui['validation-list']) return;
    const warnings = [];
    service.components.forEach(component => {
      if (component.type === 'song' && !component.song && !(component.lyric_slides || []).some(text => text.trim())) warnings.push({ title: headingOf(component), detail: 'Choose a library song or enter lyric slides.' });
      if (['psalm', 'reading', 'call_to_worship'].includes(component.type) && !(component.reference || '').trim()) warnings.push({ title: headingOf(component), detail: 'Add a Bible or psalm reference.' });
      if (component.type === 'call_to_worship' && component.external_source_failed) warnings.push({ title: headingOf(component), detail: 'ESV fetch failed. Check the manually entered wording.' });
      if (component.type === 'teaching' && !(component.text || '').trim()) warnings.push({ title: headingOf(component), detail: 'Load the selected source or enter teaching text manually.' });
      if ((component.lyric_slides || component.slide_breaks || []).some(text => text.length > 620)) warnings.push({ title: headingOf(component), detail: 'A slide is unusually dense. Consider another break.' });
    });
    const possible = Math.max(1, service.components.length);
    const score = Math.max(0, Math.round((1 - warnings.length / possible) * 100));
    if (ui['readiness-score']) ui['readiness-score'].textContent = `${score}%`;
    if (ui['readiness-bar']) ui['readiness-bar'].style.width = `${score}%`;
    ui['validation-list'].replaceChildren();
    if (!warnings.length) {
      const ok = doc.createElement('div'); ok.className = 'validation-item ok';
      const icon = doc.createElement('span'); icon.textContent = '✓'; const copy = doc.createElement('div');
      const title = doc.createElement('strong'); title.textContent = 'Ready to generate'; const detail = doc.createElement('small'); detail.textContent = 'No missing service choices found.';
      copy.append(title, detail); ok.append(icon, copy); ui['validation-list'].append(ok);
    } else warnings.slice(0, 7).forEach(warning => {
      const item = doc.createElement('div'); item.className = 'validation-item warning'; const icon = doc.createElement('span'); icon.textContent = '!';
      const copy = doc.createElement('div'); const title = doc.createElement('strong'); title.textContent = warning.title; const detail = doc.createElement('small'); detail.textContent = warning.detail;
      copy.append(title, detail); item.append(icon, copy); ui['validation-list'].append(item);
    });
  }

  async function loadService(record, options) {
    const result = await controller.loadService(record, options);
    if (controller.getState().status === 'Saved') {
      setSaveState('Saved', 'Saved');
      setSaveHelp('');
    }
    return result;
  }

  async function guardedLeave(leave) {
    if (controller.isDirty() || controller.isSaving()) {
      try {
        await controller.saveNow();
      } catch (error) {
        const generation = controller.getState().editGeneration;
        if (!confirmImpl(`Generation ${generation} is not saved. Leave and discard local edits?`)) return;
      }
    }
    return leave();
  }

  function guardedNavigation(event) {
    const link = event.currentTarget;
    const destination = link?.getAttribute('href');
    if (!link || !['/', '/library', '/generated', '/admin'].includes(destination)) return;
    event.preventDefault();
    void guardedLeave(() => locationImpl?.assign(destination)).catch(error => showToast(error.message));
  }

  async function createService() {
    return guardedLeave(async () => {
      const chosen = doc.querySelector('input[name="preset"]:checked')?.value || 'am';
      const preset = presets.find(item => item.id === chosen);
      const input = { name: preset.label, date: today(), preset: chosen };
      const response = await request('/api/services', { method: 'POST', body: JSON.stringify(input) });
      await loadService(await response.json(), { discardUnsaved: true });
      ui['new-dialog']?.close();
    });
  }

  function renderPresetChoices() {
    if (!ui['service-preset'] || !ui['preset-choices']) return;
    ui['service-preset'].replaceChildren(); ui['preset-choices'].replaceChildren();
    presets.forEach((preset, index) => {
      const option = doc.createElement('option'); option.value = preset.id; option.textContent = preset.label; ui['service-preset'].append(option);
      const choice = doc.createElement('label'); choice.className = 'preset-choice';
      const radio = doc.createElement('input'); radio.type = 'radio'; radio.name = 'preset'; radio.value = preset.id; radio.checked = index === 0;
      const text = doc.createElement('span'); const title = doc.createElement('strong'); title.textContent = preset.label; const detail = doc.createElement('small');
      detail.textContent = `${preset.components.length} items · ${preset.components.reduce((sum, item) => sum + estimatedSlides(item), 0)} starting slides`;
      text.append(title, detail); choice.append(radio, text); ui['preset-choices'].append(choice);
    });
  }

  function openReview() {
    const service = controller.getService();
    if (!service) return;
    ui['review-title'].textContent = service.name; ui['full-review'].replaceChildren();
    service.components.forEach((component, index) => {
      const row = doc.createElement('div'); row.className = 'full-review-row'; const number = doc.createElement('span'); number.textContent = index + 1;
      const text = doc.createElement('div'); const title = doc.createElement('strong'); title.textContent = headingOf(component); const detail = doc.createElement('small'); detail.textContent = detailOf(component);
      text.append(title, detail); const slides = doc.createElement('span'); const count = estimatedSlides(component); slides.textContent = `${count} slide${count === 1 ? '' : 's'}`; row.append(number, text, slides); ui['full-review'].append(row);
    });
    ui['review-dialog']?.showModal();
  }

  function setGenerationPending(pending) {
    for (const element of [ui['generate-service'], ui['review-generate']].filter(Boolean)) {
      if (pending) {
        if (!generationLabels.has(element)) {
          generationLabels.set(element, { text: element.textContent, disabled: element.disabled });
        }
        element.disabled = true;
        element.textContent = 'Generating…';
      } else {
        const original = generationLabels.get(element);
        if (original) {
          element.disabled = original.disabled;
          element.textContent = original.text;
        }
      }
    }
    if (!pending) generationLabels.clear();
  }

  function generate() {
    if (generationPromise) return generationPromise;
    const service = controller.getService();
    if (!service) return;
    setGenerationPending(true);
    generationPromise = (async () => {
      try {
        await controller.saveNow();
        setSaveState('Saving', 'Generating PowerPoint…');
        const response = await request(`/api/services/${service.id}/generate`, { method: 'POST' });
        const blob = await response.blob();
        const disposition = response.headers.get('content-disposition') || '';
        const filename = disposition.match(/filename=\"([^\"]+)\"/)?.[1] || `service-${service.date}.pptx`;
        const link = doc.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = filename;
        link.click();
        timers.setTimeout(() => URL.revokeObjectURL(link.href), 1000);
        service.status = 'completed';
        setSaveState('Saved', 'PowerPoint generated'); showToast('PowerPoint generated and saved to service history.'); ui['review-dialog']?.close();
      } catch (error) {
        setSaveState('Failed', 'Failed');
        setSaveHelp(`Generation failed: ${error.message}`);
        showToast(error.message);
      } finally {
        setGenerationPending(false);
        generationPromise = null;
      }
    })();
    return generationPromise;
  }

  function button(text, className, label) {
    const element = doc.createElement('button'); element.type = 'button'; element.className = className; element.textContent = text;
    if (label) element.setAttribute('aria-label', label); return element;
  }

  function loaderErrorNode(kind, componentId) {
    const error = doc.createElement('p');
    error.className = 'loader-error';
    error.id = `loader-error-${kind}-${componentId}`;
    error.dataset.loaderError = kind;
    error.dataset.componentId = componentId;
    error.setAttribute('role', 'alert');
    error.hidden = true;
    return error;
  }

  function renderLoaderState(componentId, kind) {
    const buttonElement = [...doc.querySelectorAll('[data-loader-kind]')]
      .find(element => element.dataset.loaderKind === kind && element.dataset.componentId === componentId);
    const error = [...doc.querySelectorAll('[data-loader-error]')]
      .find(element => element.dataset.loaderError === kind && element.dataset.componentId === componentId);
    if (!buttonElement || !error) return;
    const loader = controller.getState().loaders.get(`${kind}:${componentId}`) || { pending: false, error: null };
    buttonElement.disabled = loader.pending;
    buttonElement.textContent = loader.pending
      ? (kind === 'psalm' || kind === 'teaching' ? 'Loading…' : 'Fetching…')
      : (loader.error
        ? (buttonElement.dataset.retryLabel || (kind === 'psalm' ? 'Retry Psalm text' : 'Retry ESV text'))
        : (buttonElement.dataset.idleLabel || (kind === 'psalm' ? 'Load Psalm text' : 'Fetch ESV text')));
    error.hidden = !loader.error;
    error.textContent = loader.error || '';
  }

  function boot() {
    if (booted) return;
    booted = true;
    ui['new-service']?.addEventListener('click', () => ui['new-dialog']?.showModal());
    ui['create-service']?.addEventListener('click', event => { event.preventDefault(); createService().catch(error => showToast(error.message)); });
    ui['review-service']?.addEventListener('click', openReview);
    ui['generate-service']?.addEventListener('click', generate);
    ui['review-generate']?.addEventListener('click', event => { event.preventDefault(); generate(); });
    ui['save-now']?.addEventListener('click', () => controller.saveNow().catch(() => {}));
    ui['add-component']?.addEventListener('click', () => {
      const component = { type: 'custom_text_image', id: `component-${Date.now().toString(36)}`, heading: 'Custom slide', slides: [''], image: null };
      controller.updateService(service => { service.components.push(component); }, 'structural'); controller.selectComponent(component.id); renderOrder(); renderEditor();
    });
    ui['sign-out']?.addEventListener('click', () => {
      void guardedLeave(async () => {
        await request('/api/logout', { method: 'POST' });
        locationImpl?.assign('/login');
      }).catch(error => showToast(error.message));
    });
    doc.querySelectorAll('a[href="/"], a[href="/library"], a[href="/generated"], a[href="/admin"]').forEach(link => {
      link.addEventListener('click', guardedNavigation);
    });
    ui['service-name']?.addEventListener('input', () => setServiceField('name', ui['service-name'].value, 'heading'));
    ui['service-date']?.addEventListener('input', () => setServiceField('date', ui['service-date'].value));
    ui['service-preset']?.addEventListener('change', () => {
      const service = controller.getService(); const preset = presets.find(item => item.id === ui['service-preset'].value);
      if (!preset || !confirmImpl('Replace the current order with this preset?')) { ui['service-preset'].value = service.preset; return; }
      controller.updateService(current => { current.preset = preset.id; current.components = structuredClone(preset.components); }, 'structural');
      controller.selectComponent(controller.getService().components[0]?.id || null); renderAll();
    });
    doc.defaultView?.addEventListener('beforeunload', event => {
      if (controller.isDirty() || controller.isSaving()) { event.preventDefault(); event.returnValue = ''; }
    });
    Promise.all([
      request('/api/presets').then(response => response.json()),
      request('/api/services').then(response => response.json()),
    ]).then(([loadedPresets, services]) => {
      presets = loadedPresets; renderPresetChoices();
      const current = services.find(item => item.status === 'draft') || services.find(item => item.status !== 'archived');
      if (current) loadService(current); else ui['new-dialog']?.showModal();
    }).catch(error => showToast(error.message));
  }

  function setServiceField(field, value, scope = 'targeted') {
    controller.updateService(service => { service[field] = value; }, scope);
  }

  controller = createEditorController({
    request,
    timers,
    render: { all: renderAll, order: renderOrder, editor: renderEditor, validation: renderValidation, orderItem: updateOrderItem, counts: updateCounts, heading: updateHeading, loader: renderLoaderState },
    setSaveState,
    setSaveHelp,
    showToast,
  });

  return {
    controller: () => controller,
    loadService,
    boot,
    renderAll,
  };
}

if (typeof document !== 'undefined' && document.getElementById('component-list')) createEditorApp().boot();
