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

  it('marks the Teaching source select with component and field metadata', async () => {
    const service = makeService({ components: [{ id: 'teaching-1', type: 'teaching', heading: 'Teaching', source: 'heidelberg1891', selection: 'Q1', text: 'Text' }] });
    const app = createEditorApp({ document, request: async () => jsonResponse(service) });
    await app.loadService(service);
    const source = document.querySelector('select[data-component-id="teaching-1"][data-field="source"]');
    expect(source).not.toBeNull();
  });

});
