import { describe, expect, it, vi } from 'vitest';
import { createEditorController, SAVE_DEBOUNCE_MS } from '../crates/server/static/editor-controller.js';
import { deferred, jsonResponse, makeService } from './helpers/editor-fixture.js';

function controllerWithFetch(fetchRequest, overrides = {}) {
  return createEditorController({
    request: fetchRequest,
    render: { all: vi.fn(), order: vi.fn(), editor: vi.fn(), validation: vi.fn(), orderItem: vi.fn(), counts: vi.fn(), heading: vi.fn(), loader: vi.fn() },
    setSaveState: vi.fn(),
    setSaveHelp: vi.fn(),
    setConflict: vi.fn(),
    setConflictRecovery: vi.fn(),
    showToast: vi.fn(),
    ...overrides,
  });
}

describe('editor controller seam', () => {
  it('exposes canonical service state and generation counters', () => {
    const controller = createEditorController({
      request: async () => jsonResponse({}),
      render: {},
      setSaveState: () => {},
      setSaveHelp: () => {},
      setConflict: () => {},
      showToast: () => {},
    });

    expect(controller.getState()).toMatchObject({
      service: null,
      editGeneration: 0,
      savedGeneration: 0,
      status: 'Saved',
    });
  });

  it('retains service and component identities while merging only server metadata', async () => {
    const service = makeService();
    const response = jsonResponse({ ...service, revision: 5, status: 'draft', audit: { updated_by: 'Server' }, components: structuredClone(service.components).reverse(), name: 'Server name' });
    const controller = controllerWithFetch(async () => response);
    await controller.loadService(service);
    const serviceIdentity = controller.getService();
    const componentsIdentity = service.components;
    const readingIdentity = service.components[0];
    controller.updateComponent('reading-1', component => { component.reference = 'Romans 8:28'; }, 'field');
    await controller.saveNow();

    expect(controller.getService()).toBe(serviceIdentity);
    expect(service.components).toBe(componentsIdentity);
    expect(service.components[0]).toBe(readingIdentity);
    expect(service.name).toBe('Morning service');
    expect(service.components[0].reference).toBe('Romans 8:28');
    expect(service.revision).toBe(5);
    expect(service.audit.updated_by).toBe('Server');
  });

  it('increments a generation for every local edit and does not dirty on metadata merge', async () => {
    const service = makeService();
    const controller = controllerWithFetch(async () => jsonResponse(service));
    await controller.loadService(service);
    controller.updateComponent('reading-1', component => { component.bible_page = 843; }, 'field');
    expect(controller.getState()).toMatchObject({ editGeneration: 1, savedGeneration: 0 });
    await controller.saveNow();
    expect(controller.getState()).toMatchObject({ editGeneration: 1, savedGeneration: 1, status: 'Saved' });
  });

  it('sends ordered snapshots and saves an edit made during the first request', async () => {
    const service = makeService();
    const first = deferred();
    const second = deferred();
    const requests = [];
    let resolveFirstStarted;
    const firstStarted = new Promise(resolve => { resolveFirstStarted = resolve; });
    let resolveSecondStarted;
    const secondStarted = new Promise(resolve => { resolveSecondStarted = resolve; });
    const controller = controllerWithFetch(async (url, options) => {
      requests.push({ url, body: JSON.parse(options.body) });
      if (requests.length === 1) resolveFirstStarted();
      if (requests.length === 2) resolveSecondStarted();
      return requests.length === 1 ? first.promise : second.promise;
    });
    await controller.loadService(service);
    controller.updateComponent('reading-1', component => { component.reference = 'Romans 8:28'; });
    const save = controller.saveNow();
    await firstStarted;
    controller.updateComponent('reading-1', component => { component.reference = 'Romans 8:29'; });
    expect(requests).toHaveLength(1);
    expect(requests[0].body.components[0].reference).toBe('Romans 8:28');
    first.resolve(jsonResponse({ ...service, revision: 5 }));
    await secondStarted;
    expect(requests).toHaveLength(2);
    second.resolve(jsonResponse({ ...service, revision: 6 }));
    await save;
    expect(requests[1].body.components[0].reference).toBe('Romans 8:29');
    expect(controller.getState()).toMatchObject({ editGeneration: 2, savedGeneration: 2, status: 'Saved' });
  });

  it('settles an active save before loading a new canonical service', async () => {
    const service = makeService();
    const nextService = makeService({ id: 'service-2', name: 'Evening service', revision: 8 });
    const first = deferred();
    const events = [];
    let resolveSaveStarted;
    const saveStarted = new Promise(resolve => { resolveSaveStarted = resolve; });
    let autosaveCount = 0;
    const controller = controllerWithFetch(async () => {
      autosaveCount += 1;
      events.push(`save-start:${autosaveCount}`);
      resolveSaveStarted();
      return first.promise.then(response => {
        events.push(`save-settled:${autosaveCount}`);
        return response;
      });
    });

    await controller.loadService(service);
    controller.updateComponent('reading-1', component => { component.reference = 'Romans 8:28'; });
    const save = controller.saveNow();
    await saveStarted;
    expect(events).toContain('save-start:1');

    const load = controller.loadService(nextService);
    await Promise.resolve();
    expect(controller.getService()).toBe(service);

    first.resolve(jsonResponse({ ...service, revision: 99, audit: { updated_by: 'Old save' } }));
    await save;
    await load;

    expect(controller.getService()).toBe(nextService);
    expect(nextService.revision).toBe(8);
    expect(nextService.audit.updated_by).toBe('Test Staff');
  });

  it('serializes loading behind every dirty generation, including a follow-up save', async () => {
    const service = makeService();
    const nextService = makeService({ id: 'service-2', name: 'Evening service', revision: 8 });
    const first = deferred();
    const second = deferred();
    const events = [];
    const requests = [];
    let resolveFirstStarted;
    const firstStarted = new Promise(resolve => { resolveFirstStarted = resolve; });
    let resolveSecondStarted;
    const secondStarted = new Promise(resolve => { resolveSecondStarted = resolve; });
    const controller = controllerWithFetch(async (_url, options) => {
      requests.push(JSON.parse(options.body));
      events.push(`save-start:${requests.length}`);
      if (requests.length === 1) resolveFirstStarted();
      if (requests.length === 2) resolveSecondStarted();
      const response = requests.length === 1 ? first.promise : second.promise;
      return response.then(result => {
        events.push(`save-settled:${requests.length}`);
        return result;
      });
    });

    await controller.loadService(service);
    controller.updateComponent('reading-1', component => { component.reference = 'Romans 8:28'; });
    const save = controller.saveNow();
    await firstStarted;
    controller.updateComponent('reading-1', component => { component.reference = 'Romans 8:29'; });
    const load = controller.loadService(nextService);

    first.resolve(jsonResponse({ ...service, revision: 5 }));
    await secondStarted;
    expect(controller.getService()).toBe(service);
    expect(requests[0].components[0].reference).toBe('Romans 8:28');
    expect(requests[1].components[0].reference).toBe('Romans 8:29');

    second.resolve(jsonResponse({ ...service, revision: 6 }));
    await save;
    await load;

    expect(controller.getService()).toBe(nextService);
  });

  it('propagates an active save failure instead of replacing the local service', async () => {
    const service = makeService();
    const nextService = makeService({ id: 'service-2', name: 'Evening service' });
    const first = deferred();
    let pendingTimer;
    const controller = controllerWithFetch(async () => {
      return first.promise;
    }, {
      timers: {
        setTimeout: vi.fn(callback => {
          pendingTimer = callback;
          return 1;
        }),
        clearTimeout: vi.fn(),
      },
    });

    await controller.loadService(service);
    controller.updateComponent('reading-1', component => { component.reference = 'Romans 8:28'; });
    pendingTimer();
    await Promise.resolve();
    const load = controller.loadService(nextService);
    const failure = new Error('save failed');
    first.reject(failure);

    await expect(load).rejects.toThrow('save failed');
    expect(controller.getService()).toBe(service);
    expect(service.components[0].reference).toBe('Romans 8:28');
  });

  it('uses an exact 900ms debounce and never starts a second save while one is active', async () => {
    const service = makeService();
    const first = deferred();
    const timerDelays = [];
    let pendingTimer;
    let autosaveCount = 0;
    const controller = controllerWithFetch(async () => {
      autosaveCount += 1;
      return autosaveCount === 1 ? first.promise : jsonResponse({ ...service, revision: 6 });
    }, {
      timers: {
        setTimeout: vi.fn((callback, delay) => {
          timerDelays.push(delay);
          pendingTimer = callback;
          return 1;
        }),
        clearTimeout: vi.fn(),
      },
    });

    await controller.loadService(service);
    controller.updateComponent('reading-1', component => { component.reference = 'Romans 8:28'; });
    expect(SAVE_DEBOUNCE_MS).toBe(900);
    expect(timerDelays).toEqual([900]);
    pendingTimer();
    await Promise.resolve();
    expect(autosaveCount).toBe(1);

    controller.updateComponent('reading-1', component => { component.reference = 'Romans 8:29'; });
    pendingTimer();
    await Promise.resolve();
    expect(autosaveCount).toBe(1);

    first.resolve(jsonResponse({ ...service, revision: 5 }));
    await controller.saveNow();
    expect(autosaveCount).toBe(2);
  });

  it('Save now bypasses debounce, joins an active request, and waits for a follow-up generation', async () => {
    const service = makeService();
    const first = deferred();
    const second = deferred();
    const puts = [];
    const controller = controllerWithFetch(async (_url, options) => {
      puts.push(JSON.parse(options.body));
      return puts.length === 1 ? first.promise : second.promise;
    }, { timers: { setTimeout: vi.fn(() => 99), clearTimeout: vi.fn(), setInterval: vi.fn(), clearInterval: vi.fn() } });
    await controller.loadService(service);
    controller.updateComponent('reading-1', component => { component.reference = 'John 3:16'; });
    const firstSave = controller.saveNow();
    const joinedSave = controller.saveNow();
    expect(puts).toHaveLength(1);
    first.resolve(jsonResponse({ ...service, revision: 5 }));
    await Promise.resolve();
    controller.updateComponent('reading-1', component => { component.reference = 'John 3:17'; });
    second.resolve(jsonResponse({ ...service, revision: 6 }));
    await Promise.all([firstSave, joinedSave]);
    expect(puts).toHaveLength(2);
    expect(puts[1].components[0].reference).toBe('John 3:17');
  });

  it('keeps local data and exposes a retryable Failed state after a PUT failure', async () => {
    const service = makeService();
    const setSaveState = vi.fn();
    let attempts = 0;
    const controller = controllerWithFetch(async () => {
      attempts += 1;
      if (attempts === 1) return new Response(JSON.stringify({ error: 'network down' }), { status: 503 });
      return jsonResponse({ ...service, revision: 5 });
    }, { setSaveState });
    await controller.loadService(service);
    controller.updateComponent('reading-1', component => { component.reference = 'John 3:16'; });
    await expect(controller.saveNow()).rejects.toThrow('network down');
    expect(controller.getService().components[0].reference).toBe('John 3:16');
    expect(controller.getState()).toMatchObject({ status: 'Failed', editGeneration: 1, savedGeneration: 0 });
    await controller.saveNow();
    expect(attempts).toBe(2);
    expect(controller.getState().status).toBe('Saved');
  });

  it('clears a timer created by an edit during a save failure while retaining dirty Failed state', async () => {
    const service = makeService();
    const first = deferred();
    const timers = {
      setTimeout: vi.fn(() => timers.setTimeout.mock.calls.length),
      clearTimeout: vi.fn(),
    };
    const controller = controllerWithFetch(async () => {
      return first.promise;
    }, { timers });

    await controller.loadService(service);
    controller.updateComponent('reading-1', component => { component.reference = 'John 3:16'; });
    const save = controller.saveNow();
    controller.updateComponent('reading-1', component => { component.reference = 'John 3:17'; });
    first.reject(new Error('network down'));

    await expect(save).rejects.toThrow('network down');
    expect(timers.clearTimeout).toHaveBeenCalledTimes(2);
    expect(controller.getState()).toMatchObject({ saveTimer: null, status: 'Failed', editGeneration: 2, savedGeneration: 0 });
    expect(controller.isDirty()).toBe(true);
  });

  it('marks a 409 as an unresolved conflict that cannot be retried stale', async () => {
    const service = makeService();
    const setSaveHelp = vi.fn();
    const showToast = vi.fn();
    const controller = controllerWithFetch(async () => {
      return new Response(JSON.stringify({ error: 'this service changed in another browser; reload before saving' }), { status: 409 });
    }, { setSaveHelp, showToast });
    await controller.loadService(service);
    controller.updateComponent('reading-1', component => { component.reference = 'John 3:16'; });
    await expect(controller.saveNow()).rejects.toThrow('another browser');
    await expect(controller.saveNow()).rejects.toThrow('another browser');
    expect(controller.getService().components[0].reference).toBe('John 3:16');
    expect(controller.isDirty()).toBe(true);
    expect(controller.getState().conflict).not.toBeNull();
    expect(controller.getState().status).toBe('Failed');
    expect(setSaveHelp).toHaveBeenLastCalledWith(expect.stringContaining('another browser'));
    expect(showToast).toHaveBeenCalledWith(expect.stringContaining('another browser'));
  });

  it('preserves optional metadata omitted by an older compatible save response', async () => {
    const service = makeService();
    const originalAudit = service.audit;
    const controller = controllerWithFetch(async () => jsonResponse({ id: service.id, revision: 5, status: 'draft', components: [] }));
    await controller.loadService(service);
    controller.updateComponent('reading-1', component => { component.reference = 'John 3:16'; });
    await controller.saveNow();
    expect(controller.getService().audit).toBe(originalAudit);
    expect(controller.getService().components[0].reference).toBe('John 3:16');
  });

});
