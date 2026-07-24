import { describe, expect, it, vi } from 'vitest';
import { createEditorController } from '../crates/server/static/editor-controller.js';
import { deferred, errorResponse, jsonResponse, makeService } from './helpers/editor-fixture.js';

function loaderController(request, timers = { setTimeout, clearTimeout }) {
  return createEditorController({
    request,
    timers,
    render: { all: vi.fn(), order: vi.fn(), editor: vi.fn(), validation: vi.fn(), orderItem: vi.fn(), counts: vi.fn(), heading: vi.fn(), loader: vi.fn() },
    setSaveState: vi.fn(),
    setSaveHelp: vi.fn(),
    showToast: vi.fn(),
  });
}

describe('editor loaders', () => {
  it('applies a Psalm result to the current ID after a structural rerender without changing the entered range', async () => {
    const service = makeService();
    const response = deferred();
    const controller = loaderController(async () => response.promise);
    await controller.loadService(service);
    const load = controller.loadPsalm('psalm-1', 'Psalm 23:1–6');
    controller.selectComponent('reading-1');
    response.resolve(jsonResponse({ reference: 'Psalm 23', meter: '11 11 11', slides: ['New group', 'Second group'] }));
    await load;
    expect(service.components.find(component => component.id === 'psalm-1').reference).toBe('Psalm 23:1–6');
    expect(service.components.find(component => component.id === 'psalm-1').slide_breaks).toEqual(['New group', 'Second group']);
  });

  it('discards a Psalm result when the component is removed or the reference changes', async () => {
    const service = makeService();
    const response = deferred();
    const controller = loaderController(async () => response.promise);
    await controller.loadService(service);
    const load = controller.loadPsalm('psalm-1', 'Psalm 23:1–6');
    service.components[1].reference = 'Psalm 24:1–6';
    response.resolve(jsonResponse({ reference: 'Psalm 23', meter: '11 11 11', slides: ['old result'] }));
    await load;
    expect(service.components[1].reference).toBe('Psalm 24:1–6');
    expect(service.components[1].slide_breaks).toEqual(['The LORD is my shepherd']);
  });

  it('rejects malformed Psalm data without blanking existing slide breaks', async () => {
    const service = makeService();
    const controller = loaderController(async () => jsonResponse({ reference: 'Psalm 23', meter: '11 11 11', slides: [7] }));
    await controller.loadService(service);
    await controller.loadPsalm('psalm-1', 'Psalm 23:1–6');
    expect(service.components[1].slide_breaks).toEqual(['The LORD is my shepherd']);
    expect(controller.getState().loaders.get('psalm:psalm-1').error).toMatch(/malformed/i);
  });

  it('discards a completion from an invalidated Psalm request sequence', async () => {
    const service = makeService();
    const response = deferred();
    const controller = loaderController(async () => response.promise);
    await controller.loadService(service);
    const load = controller.loadPsalm('psalm-1', 'Psalm 23:1–6');
    controller.getState().loaders.get('psalm:psalm-1').sequence += 1;
    response.resolve(jsonResponse({ reference: 'Psalm 23', meter: '11 11 11', slides: ['stale'] }));
    await load;
    expect(service.components[1].slide_breaks).toEqual(['The LORD is my shepherd']);
  });

  it('invalidates an in-flight loader when a structural edit replaces its ID', async () => {
    const service = makeService();
    const oldResponse = deferred();
    const newResponse = deferred();
    let calls = 0;
    const controller = loaderController(async () => ++calls === 1 ? oldResponse.promise : newResponse.promise);
    await controller.loadService(service);
    const oldLoad = controller.loadPsalm('psalm-1', 'Psalm 23:1–6');
    controller.updateService(current => {
      current.components[1] = { ...current.components[1], reference: 'Psalm 23:1–6', slide_breaks: ['replacement'] };
    }, 'structural');
    const newLoad = controller.loadPsalm('psalm-1', 'Psalm 23:1–6');

    oldResponse.resolve(jsonResponse({ reference: 'Psalm 23', meter: '11 11 11', slides: ['stale'] }));
    newResponse.resolve(jsonResponse({ reference: 'Psalm 23', meter: '11 11 11', slides: ['fresh'] }));
    await Promise.all([oldLoad, newLoad]);

    expect(calls).toBe(2);
    expect(service.components[1].slide_breaks).toEqual(['fresh']);
  });

  it('reports a stale reference when a loader request fails after the reference changes', async () => {
    const service = makeService();
    const response = deferred();
    const controller = loaderController(async () => response.promise);
    await controller.loadService(service);
    const load = controller.loadPsalm('psalm-1', 'Psalm 23:1–6');
    service.components[1].reference = 'Psalm 24:1–6';
    response.reject(new Error('network down'));
    await load;

    expect(controller.getState().loaders.get('psalm:psalm-1').error).toMatch(/reference changed/i);
  });

  it('accepts an intentionally empty Psalm slide list and disables overlapping loads', async () => {
    const service = makeService();
    const response = deferred();
    let calls = 0;
    const controller = loaderController(async () => { calls += 1; return response.promise; });
    await controller.loadService(service);
    const first = controller.loadPsalm('psalm-1', 'Psalm 23:1–6');
    const second = controller.loadPsalm('psalm-1', 'Psalm 23:1–6');
    expect(calls).toBe(1);
    response.resolve(jsonResponse({ reference: 'Psalm 23', meter: '11 11 11', slides: [] }));
    await Promise.all([first, second]);
    expect(service.components[1].slide_breaks).toEqual([]);
    expect(controller.getState().loaders.get('psalm:psalm-1').pending).toBe(false);
  });

  it('applies valid ESV text to the current component and clears its failure marker', async () => {
    const service = makeService();
    service.components[2].text = 'Manual wording';
    service.components[2].external_source_failed = true;
    const controller = loaderController(async () => jsonResponse({ ok: true, reference: 'Psalm 96:2', text: 'Fetched wording' }));
    await controller.loadService(service);
    await controller.loadEsv('call-1', 'Psalm 96:2');
    expect(service.components[2].text).toBe('Fetched wording');
    expect(service.components[2].external_source_failed).toBe(false);
  });

  it.each([
    ['endpoint failure', () => jsonResponse({ ok: false, reference: 'Psalm 96:2', text: '', warning: 'ESV is unavailable.' })],
    ['malformed success', () => jsonResponse({ ok: true, reference: 'Psalm 96:2' })],
    ['HTTP failure', () => errorResponse('ESV authentication failed', 401)],
  ])('preserves manual ESV wording on %s and exposes retry state', async (_label, responseFactory) => {
    const service = makeService();
    service.components[2].text = 'Keep this wording';
    const controller = loaderController(async () => responseFactory());
    await controller.loadService(service);
    await controller.loadEsv('call-1', 'Psalm 96:2');
    expect(service.components[2].text).toBe('Keep this wording');
    expect(service.components[2].external_source_failed).toBe(true);
    expect(controller.getState().loaders.get('esv:call-1').error).toBeTruthy();
    expect(controller.getState().loaders.get('esv:call-1').pending).toBe(false);
  });

  it('keeps ESV pending after timeout until the aborted request promise settles', async () => {
    const service = makeService();
    const abort = vi.fn();
    let settleAbort;
    const timers = { setTimeout: vi.fn(() => 7), clearTimeout: vi.fn(), setInterval: vi.fn(), clearInterval: vi.fn() };
    const controller = loaderController(async (_url, options) => {
      expect(options.signal).toBeDefined();
      options.signal.addEventListener('abort', () => abort());
      return new Promise((resolve, reject) => {
        settleAbort = () => {
          const error = new Error('aborted');
          error.name = 'AbortError';
          reject(error);
        };
        options.signal.addEventListener('abort', () => {
          // Abort is observable immediately, but this request settles later.
        });
      });
    }, timers);
    await controller.loadService(service);
    const load = controller.loadEsv('call-1', 'Psalm 96:2');
    const timeoutCallback = timers.setTimeout.mock.calls.find(([, duration]) => duration === 10_000)[0];
    timeoutCallback();
    await Promise.resolve();
    expect(abort).toHaveBeenCalledOnce();
    expect(controller.getState().loaders.get('esv:call-1').pending).toBe(true);
    settleAbort();
    await load;
    expect(timers.clearTimeout).toHaveBeenCalledWith(7);
    expect(controller.getState().loaders.get('esv:call-1').pending).toBe(false);
  });

  it('does not start a second ESV request while the first is pending', async () => {
    const service = makeService();
    const pending = deferred();
    let calls = 0;
    const controller = loaderController(async () => { calls += 1; return pending.promise; });
    await controller.loadService(service);
    const first = controller.loadEsv('call-1', 'Psalm 96:2');
    const second = controller.loadEsv('call-1', 'Psalm 96:2');
    expect(calls).toBe(1);
    pending.resolve(jsonResponse({ ok: true, reference: 'Psalm 96:2', text: 'Fetched' }));
    await Promise.all([first, second]);
  });

  it('loads a friendly WSC selection into editable teaching text', async () => {
    const service = makeService({ components: [{ id: 'teaching-1', type: 'teaching', heading: 'Teaching', source: 'westminster_shorter_catechism', selection: 'Q. 1', text: '' }] });
    const controller = loaderController(async url => {
      expect(url).toContain('/api/teaching?');
      expect(url).toContain('selection=Q.%201');
      return jsonResponse({ source: 'westminster_shorter_catechism', selection: 1, question: 'What is the chief end of man?', answer: 'To glorify God, and to enjoy him forever.' });
    });
    await controller.loadService(service);
    await controller.loadTeaching('teaching-1', service.components[0].source, service.components[0].selection);
    expect(service.components[0].text).toBe('What is the chief end of man?\n\nTo glorify God, and to enjoy him forever.');
  });
});
