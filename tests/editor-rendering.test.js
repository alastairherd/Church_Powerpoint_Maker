import { beforeEach, describe, expect, it, vi } from 'vitest';
import { createEditorApp } from '../crates/server/static/app.js';
import { deferred, errorResponse, installBuilderDom, jsonResponse, makeService } from './helpers/editor-fixture.js';

describe('editor render boundaries', () => {
  beforeEach(() => {
    installBuilderDom(document);
  });

  it('keeps the Reading field and component list nodes stable while typing', async () => {
    const service = makeService();
    const app = createEditorApp({
      document,
      request: async () => jsonResponse(service),
      timers: { setTimeout: vi.fn(() => 1), clearTimeout: vi.fn(), setInterval: vi.fn(), clearInterval: vi.fn() },
    });
    await app.loadService(service);
    const list = document.getElementById('component-list');
    const editor = document.getElementById('editor-panel');
    const reference = editor.querySelector('[data-field="reference"]');
    reference.focus();
    reference.setSelectionRange(5, 5);
    reference.value = 'Romans 8:28';
    reference.dispatchEvent(new Event('input', { bubbles: true }));

    expect(document.getElementById('component-list')).toBe(list);
    expect(document.getElementById('editor-panel')).toBe(editor);
    expect(editor.querySelector('[data-field="reference"]')).toBe(reference);
    expect(document.activeElement).toBe(reference);
    expect(app.controller().getService().components[0].reference).toBe('Romans 8:28');
  });

  it('updates a heading order item without rebuilding unrelated editor DOM', async () => {
    const service = makeService();
    const app = createEditorApp({ document, request: async () => jsonResponse(service) });
    await app.loadService(service);
    const list = document.getElementById('component-list');
    const editor = document.getElementById('editor-panel');
    const heading = editor.querySelector('[data-field="heading"]');
    heading.value = 'Gospel reading';
    heading.dispatchEvent(new Event('input', { bubbles: true }));

    expect(document.getElementById('component-list')).toBe(list);
    expect(document.getElementById('editor-panel')).toBe(editor);
    expect(list.querySelector('[data-id="reading-1"] strong').textContent).toBe('Gospel reading');
  });

  it('updates a Reading order summary without rebuilding the list or editor', async () => {
    const service = makeService();
    const app = createEditorApp({ document, request: async () => jsonResponse(service) });
    await app.loadService(service);
    const list = document.getElementById('component-list');
    const editor = document.getElementById('editor-panel');
    const reference = editor.querySelector('[data-field="reference"]');
    reference.value = 'Romans 8:28';
    reference.dispatchEvent(new Event('input', { bubbles: true }));

    expect(document.getElementById('component-list')).toBe(list);
    expect(document.getElementById('editor-panel')).toBe(editor);
    expect(list.querySelector('[data-id="reading-1"] small').textContent).toBe('Romans 8:28');
  });

  it('updates a custom song title in its order summary without rebuilding the editor', async () => {
    const service = makeService({ components: [{ id: 'song-1', type: 'song', title: 'Old song', song: null, lyric_slides: ['Lyrics'], credits: '' }] });
    const app = createEditorApp({ document, request: async () => jsonResponse(service) });
    await app.loadService(service);
    const list = document.getElementById('component-list');
    const editor = document.getElementById('editor-panel');
    const title = editor.querySelector('[data-field="title"]');
    title.value = 'New song';
    title.dispatchEvent(new Event('input', { bubbles: true }));

    expect(document.getElementById('component-list')).toBe(list);
    expect(document.getElementById('editor-panel')).toBe(editor);
    expect(list.querySelector('[data-id="song-1"] strong').textContent).toBe('New song');
  });


  it('target-updates every reference or selection used by detailOf', async () => {
    const service = makeService({ components: [
      { id: 'call-1', type: 'call_to_worship', heading: 'Call', reference: 'Psalm 96:2', text: 'Sing.' },
      { id: 'psalm-1', type: 'psalm', heading: 'Psalm', reference: 'Psalm 23:1–6', slide_breaks: ['First'] },
      { id: 'teaching-1', type: 'teaching', heading: 'Teaching', source: 'heidelberg1891', selection: 'Q1', text: 'Text' },
    ] });
    const app = createEditorApp({ document, request: async () => jsonResponse(service) });
    await app.loadService(service);
    const list = document.getElementById('component-list');
    const cases = [
      ['call-1', 'reference', 'Psalm 100:1', 'Psalm 100:1'],
      ['psalm-1', 'reference', 'Psalm 24:1–6', 'Psalm 24:1–6'],
      ['teaching-1', 'selection', 'Q2', 'Q2'],
    ];
    for (const [id, field, value, summary] of cases) {
      document.querySelector(`[data-id="${id}"] .component-main`).click();
      const editor = document.getElementById('editor-panel');
      const input = editor.querySelector(`[data-field="${field}"]`);
      input.value = value;
      input.dispatchEvent(new Event('input', { bubbles: true }));
      expect(document.getElementById('component-list')).toBe(list);
      expect(document.getElementById('editor-panel')).toBe(editor);
      expect(list.querySelector(`[data-id="${id}"] small`).textContent).toBe(summary);
    }
  });

  it('updates counts when a Psalm loader changes slide breaks', async () => {
    const service = makeService();
    const app = createEditorApp({
      document,
      request: async url => {
        if (url.includes('/psalm?')) return jsonResponse({ reference: 'Psalm 23:1–6', slides: ['First', 'Second'], meter: 'Common metre' });
        return jsonResponse(service);
      },
    });
    await app.loadService(service);
    document.querySelector('[data-id="psalm-1"] .component-main').click();
    document.querySelector('button.button-secondary').click();
    await new Promise(resolve => setTimeout(resolve, 0));

    expect(document.getElementById('slide-count').textContent).toBe('4');
    expect(document.getElementById('review-slides').textContent).toBe('4 slides');
  });

  it('renders and persists the Psalm show verse numbers checkbox', async () => {
    const service = makeService();
    const app = createEditorApp({ document, request: async () => jsonResponse(service) });
    await app.loadService(service);
    document.querySelector('[data-id="psalm-1"] .component-main').click();

    const checkbox = document.querySelector('[data-field="show_verse_numbers"]');
    expect(checkbox.type).toBe('checkbox');
    expect(checkbox.checked).toBe(true);
    expect(checkbox.parentElement.textContent).toContain('Show verse numbers');

    checkbox.click();
    expect(app.controller().getService().components[1].show_verse_numbers).toBe(false);
  });

  it('associates persistent loader errors with their action and exposes the save live region', async () => {
    const service = makeService();
    const app = createEditorApp({
      document,
      request: async url => {
        if (url.includes('/psalm?')) return errorResponse('Psalm unavailable', 503);
        return jsonResponse(service);
      },
    });
    await app.loadService(service);
    document.querySelector('[data-id="psalm-1"] .component-main').click();
    const saveState = document.getElementById('save-state');
    const psalmButton = document.querySelector('[data-loader-kind="psalm"]');
    const psalmError = document.querySelector('[data-loader-error="psalm"]');
    psalmButton.click();
    await new Promise(resolve => setTimeout(resolve, 0));
    expect(saveState.getAttribute('aria-live')).toBe('polite');
    expect(psalmButton.getAttribute('aria-describedby')).toBe(psalmError.id);
    expect(psalmError.getAttribute('role')).toBe('alert');
    expect(psalmError.hidden).toBe(false);
    expect(psalmError.textContent).toContain('Psalm unavailable');
    expect(psalmButton.textContent).toBe('Retry Psalm text');
  });

  it('maps save states to CSS classes and resets state after a successful load', async () => {
    const service = makeService();
    const pending = deferred();
    let saveStarted;
    const started = new Promise(resolve => { saveStarted = resolve; });
    const app = createEditorApp({
      document,
      request: async (url, options) => {
        if (url.includes('/autosave')) { saveStarted(); return pending.promise; }
        return jsonResponse(service);
      },
    });
    const state = document.getElementById('save-state');
    const help = document.getElementById('save-help');
    state.className = 'save-state error'; help.hidden = false; help.textContent = 'old failure';
    await app.loadService(service);
    expect(state.className).toBe('save-state ');
    expect(state.textContent).toContain('Saved');
    expect(help.hidden).toBe(true);

    const reference = document.querySelector('[data-field="reference"]');
    reference.value = 'Romans 8:28';
    reference.dispatchEvent(new Event('input', { bubbles: true }));
    expect(state.className).toBe('save-state saving');
    const save = app.controller().saveNow();
    await started;
    expect(state.className).toBe('save-state saving');
    pending.resolve(jsonResponse({ ...service, revision: 5 }));
    await save;
    expect(state.className).toBe('save-state ');

    const failingApp = createEditorApp({ document, request: async url => {
      if (url.includes('/autosave')) return errorResponse('save failed', 500);
      return jsonResponse(service);
    }});
    await failingApp.loadService(makeService());
    const failingReference = document.querySelector('[data-field="reference"]');
    failingReference.value = 'Romans 8:29';
    failingReference.dispatchEvent(new Event('input', { bubbles: true }));
    await expect(failingApp.controller().saveNow()).rejects.toThrow('save failed');
    expect(document.getElementById('save-state').className).toBe('save-state error');
  });

  it('disables both generation actions while saving and downloading, then restores them', async () => {
    const service = makeService();
    const generated = deferred();
    let generationCalls = 0;
    vi.stubGlobal('URL', { createObjectURL: vi.fn(() => 'blob:generated'), revokeObjectURL: vi.fn() });
    const app = createEditorApp({
      document,
      request: async url => {
        if (url === '/api/presets') return jsonResponse([]);
        if (url === '/api/services') return jsonResponse([]);
        if (url.includes('/generate')) {
          generationCalls += 1;
          return generated.promise;
        }
        return jsonResponse(service);
      },
    });
    try {
      await app.loadService(service);
      app.boot();
      vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});
      const primary = document.getElementById('generate-service');
      const review = document.getElementById('review-generate');
      primary.click();
      expect(primary.disabled).toBe(true);
      expect(review.disabled).toBe(true);
      expect(primary.textContent).toBe('Generating…');
      expect(review.textContent).toBe('Generating…');
      primary.click();
      await new Promise(resolve => setTimeout(resolve, 0));
      expect(generationCalls).toBe(1);

      generated.resolve(new Response('pptx', { headers: { 'content-disposition': 'attachment; filename="service.pptx"' } }));
      await new Promise(resolve => setTimeout(resolve, 0));
      expect(primary.disabled).toBe(false);
      expect(review.disabled).toBe(false);
      expect(primary.textContent).toBe('Generate PowerPoint');
      expect(review.textContent).toBe('Generate PowerPoint');
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('restores generation actions after a generation error, including the review action', async () => {
    const service = makeService();
    const app = createEditorApp({
      document,
      request: async url => {
        if (url === '/api/presets') return jsonResponse([]);
        if (url === '/api/services') return jsonResponse([]);
        if (url.includes('/generate')) throw new Error('download failed');
        return jsonResponse(service);
      },
    });
    await app.loadService(service);
    app.boot();
    const primary = document.getElementById('generate-service');
    const review = document.getElementById('review-generate');
    review.click();
    expect(primary.textContent).toBe('Generating…');
    expect(review.textContent).toBe('Generating…');
    await new Promise(resolve => setTimeout(resolve, 0));
    expect(primary.disabled).toBe(false);
    expect(review.disabled).toBe(false);
    expect(primary.textContent).toBe('Generate PowerPoint');
    expect(review.textContent).toBe('Generate PowerPoint');
  });

  it('guards dirty library navigation and beforeunload without hard-coded browser globals', async () => {
    const service = makeService();
    const location = { assign: vi.fn() };
    const confirmImpl = vi.fn().mockReturnValue(false);
    const app = createEditorApp({
      document,
      locationImpl: location,
      confirmImpl,
      request: async (url) => {
        if (url.includes('/autosave')) return errorResponse('save unavailable', 503);
        if (url === '/api/presets' || url === '/api/services') return jsonResponse([]);
        return jsonResponse(service);
      },
      timers: { setTimeout: vi.fn(() => 1), clearTimeout: vi.fn(), setInterval: vi.fn(), clearInterval: vi.fn() },
    });
    app.boot();
    await app.loadService(service);
    app.controller().updateComponent('reading-1', component => { component.reference = 'John 3:16'; });

    const navigation = document.querySelector('a[href="/library"]');
    const click = new MouseEvent('click', { bubbles: true, cancelable: true });
    navigation.dispatchEvent(click);
    await new Promise(resolve => setTimeout(resolve, 0));
    expect(click.defaultPrevented).toBe(true);
    expect(confirmImpl).toHaveBeenCalledWith('Generation 1 is not saved. Leave and discard local edits?');
    expect(location.assign).not.toHaveBeenCalled();

    const beforeunload = new Event('beforeunload', { cancelable: true });
    document.defaultView.dispatchEvent(beforeunload);
    expect(beforeunload.defaultPrevented).toBe(true);
  });

  it('keeps the save live region to its canonical label and puts failure detail in help text', async () => {
    const service = makeService();
    const app = createEditorApp({
      document,
      request: async url => {
        if (url.includes('/autosave')) return errorResponse('offline', 503);
        return jsonResponse(service);
      },
    });
    await app.loadService(service);
    app.controller().updateComponent('reading-1', component => { component.reference = 'John 3:16'; });
    await expect(app.controller().saveNow()).rejects.toThrow('offline');
    expect(document.getElementById('save-state').textContent).toBe('Failed');
    expect(document.getElementById('save-help').textContent).toContain('offline');
  });

  it('guards every builder navigation destination, including dynamically rendered library links', async () => {
    const service = makeService({ components: [{ id: 'song-1', type: 'song', title: 'Song', song: null, lyric_slides: ['Lyrics'], credits: '' }] });
    const location = { assign: vi.fn() };
    const confirmImpl = vi.fn().mockReturnValue(false);
    const app = createEditorApp({
      document,
      locationImpl: location,
      confirmImpl,
      request: async url => {
        if (url === '/api/presets' || url === '/api/services') return jsonResponse([]);
        if (url.includes('/autosave')) return errorResponse('save unavailable', 503);
        return jsonResponse(service);
      },
    });
    app.boot();
    await app.loadService(service);
    app.controller().updateComponent('song-1', component => { component.title = 'Changed'; });
      document.querySelector('[data-id="song-1"] .component-main').click();
    const dynamicLibrary = [...document.querySelectorAll('a[href="/library"]')].find(link => link.textContent.includes('Browse'));
      const destinations = [
        document.querySelector('a[href="/"]'),
        document.querySelector('a[href="/admin"]'),
        document.querySelector('a[href="/generated"]'),
        document.querySelector('a[href="/library"]'),
      dynamicLibrary,
    ];
    for (const link of destinations) {
      const click = new MouseEvent('click', { bubbles: true, cancelable: true });
      link.dispatchEvent(click);
      await new Promise(resolve => setTimeout(resolve, 0));
      expect(click.defaultPrevented).toBe(true);
    }
    expect(location.assign).not.toHaveBeenCalled();
    expect(confirmImpl).toHaveBeenCalledTimes(5);
  });

  it('scrolls the editor panel into view when a component is selected', async () => {
    const service = makeService();
    const app = createEditorApp({ document, request: async () => jsonResponse(service) });
    await app.loadService(service);
    const editor = document.getElementById('editor-panel');
    editor.scrollIntoView = vi.fn();

    document.querySelector('[data-id="psalm-1"] .component-main').click();

    expect(editor.scrollTop).toBe(0);
    expect(editor.scrollIntoView).toHaveBeenCalledWith({ block: 'nearest' });
  });

  it('marks song and psalm order items with their type class for highlighting', async () => {
    const service = makeService({ components: [
      ...makeService().components,
      { id: 'song-1', type: 'song', title: 'Amazing Grace', song: null, lyric_slides: ['Lyrics'], credits: '' },
    ] });
    const app = createEditorApp({ document, request: async () => jsonResponse(service) });
    await app.loadService(service);

    expect(document.querySelector('[data-id="song-1"]').classList.contains('component-item--song')).toBe(true);
    expect(document.querySelector('[data-id="psalm-1"]').classList.contains('component-item--psalm')).toBe(true);
    const reading = document.querySelector('[data-id="reading-1"]');
    expect(reading.classList.contains('component-item--reading')).toBe(true);
    expect(reading.classList.contains('component-item--song')).toBe(false);
    expect(reading.dataset.type).toBe('reading');
  });

  it('reorders notice rows by drag and drop and persists the new order', async () => {
    const service = makeService({ components: [{ id: 'notices-1', type: 'notices', heading: 'Notices', rows: [
      { when: 'Today', title: 'Lunch', details: '', emphasis: false },
      { when: 'Wednesday', title: 'Bible study', details: '', emphasis: false },
      { when: 'Saturday', title: 'Walk', details: '', emphasis: false },
    ] }] });
    const app = createEditorApp({ document, request: async () => jsonResponse(service) });
    await app.loadService(service);
    document.querySelector('[data-id="notices-1"] .component-main').click();

    const list = document.querySelector('.notice-editor');
    expect(list.querySelectorAll('.drag-handle')).toHaveLength(3);
    const [first, second] = list.children;
    first.dispatchEvent(new Event('dragstart'));
    second.dispatchEvent(new Event('dragover', { cancelable: true }));
    list.dispatchEvent(new Event('drop'));

    const rows = app.controller().getService().components[0].rows;
    expect(rows.map(row => row.title)).toEqual(['Bible study', 'Lunch', 'Walk']);
    const rerendered = document.querySelector('.notice-editor');
    expect([...rerendered.querySelectorAll('input[placeholder="Title"]')].map(input => input.value))
      .toEqual(['Bible study', 'Lunch', 'Walk']);
  });

  it('renders an automatic load action for every teaching source', async () => {
    const service = makeService({ components: [
      { id: 'teaching-1', type: 'teaching', heading: 'Teaching', source: 'heidelberg1891', selection: '1', text: '' },
    ] });
    let requested = null;
    const app = createEditorApp({ document, request: async url => {
      if (url.includes('/api/teaching')) {
        requested = url;
        return jsonResponse({ source: 'heidelberg1891', selection: '1', question: 'What is your only comfort?', answer: 'That I am not my own.' });
      }
      return jsonResponse(service);
    } });
    await app.loadService(service);
    document.querySelector('[data-id="teaching-1"] .component-main').click();

    const load = document.querySelector('[data-loader-kind="teaching"]');
    expect(load.textContent).toBe('Load Heidelberg question');
    load.click();
    await new Promise(resolve => setTimeout(resolve, 0));
    expect(requested).toContain('source=heidelberg1891');
    expect(app.controller().getService().components[0].text).toBe('What is your only comfort?\n\nThat I am not my own.');

    const source = document.querySelector('select[data-field="source"]');
    source.value = 'westminster_confession_original_british';
    source.dispatchEvent(new Event('change'));
    expect(document.querySelector('[data-loader-kind="teaching"]').textContent).toBe('Load WCF section');
    expect(document.querySelector('[data-field="selection"]').placeholder).toBe('e.g. 21 or 21.8');
  });

  it('marks the Teaching source select with component and field metadata', async () => {
    const service = makeService({ components: [{ id: 'teaching-1', type: 'teaching', heading: 'Teaching', source: 'heidelberg1891', selection: 'Q1', text: 'Text' }] });
    const app = createEditorApp({ document, request: async () => jsonResponse(service) });
    await app.loadService(service);
    const source = document.querySelector('select[data-component-id="teaching-1"][data-field="source"]');
    expect(source).not.toBeNull();
  });

});
