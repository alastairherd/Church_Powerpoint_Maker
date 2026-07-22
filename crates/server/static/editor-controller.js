export const SAVE_DEBOUNCE_MS = 900;
export const LOADER_TIMEOUT_MS = 10_000;

export function createEditorController({
  request,
  timers = { setTimeout, clearTimeout },
  makeAbortController = () => new AbortController(),
  render = {},
  setSaveState,
  setSaveHelp,
  setConflict,
  setConflictRecovery,
  showToast,
}) {
  const state = {
    service: null,
    lease: null,
    selectedId: null,
    editGeneration: 0,
    savedGeneration: 0,
    saveTimer: null,
    activeSave: null,
    transition: Promise.resolve(),
    transitionDepth: 0,
    status: 'Saved',
    conflict: null,
    activeLeaseRenewal: null,
    loaders: new Map(),
  };
  const noop = () => {};

  function mergeServerMetadata(saved) {
    if (!state.service) return;
    state.service.revision = saved.revision;
    if (saved.status !== undefined) state.service.status = saved.status;
    if (saved.audit !== undefined) state.service.audit = saved.audit;
    if (Object.prototype.hasOwnProperty.call(saved, 'lease')) {
      state.service.lease = saved.lease;
      state.lease = saved.lease;
    }
  }

  function findComponent(id) {
    return state.service?.components.find(component => component.id === id) || null;
  }

  async function checkedRequest(url, options = {}) {
    const response = await request(url, options);
    if (response.ok) return response;
    const data = await response.json().catch(() => ({}));
    const error = new Error(data.error || `Request failed (${response.status})`);
    error.status = response.status;
    error.body = data;
    throw error;
  }

  function loaderKey(kind, componentId) {
    return `${kind}:${componentId}`;
  }

  async function runLoader(kind, componentId, reference, url, parseSuccess, applySuccess, applyFailure = null) {
    const key = loaderKey(kind, componentId);
    const previous = state.loaders.get(key) || { sequence: 0, pending: false, reference: '', error: null };
    if (previous.pending) return;
    const requestSequence = previous.sequence + 1;
    const loader = { sequence: requestSequence, pending: true, reference, error: null };
    state.loaders.set(key, loader);
    (render.loader || noop)(componentId, kind);
    const abortController = makeAbortController();
    let timedOut = false;
    const timeoutId = timers.setTimeout(() => {
      timedOut = true;
      abortController.abort();
      timers.clearTimeout(timeoutId);
    }, LOADER_TIMEOUT_MS);
    try {
      const response = await checkedRequest(url, { signal: abortController.signal });
      const data = await response.json();
      const parsed = parseSuccess(data);
      const current = findComponent(componentId);
      const latest = state.loaders.get(key);
      if (!current || !latest || latest.sequence !== requestSequence || current.reference !== reference) {
        if (current && latest && current.reference !== reference) {
          latest.error = 'The reference changed while this request was loading. Load the new range explicitly.';
          (render.loader || noop)(componentId, kind);
        }
        return;
      }
      applySuccess(current, parsed);
      latest.error = null;
      markDirty('loader');
    } catch (error) {
      const latest = state.loaders.get(key);
      const current = findComponent(componentId);
      if (latest && latest.sequence === requestSequence && current) {
        if (current.reference !== reference) {
          latest.error = 'The reference changed while this request was loading. Load the new range explicitly.';
        } else {
          latest.error = timedOut ? 'The request timed out after 10 seconds. Retry.' : error.message;
          if (applyFailure) {
            applyFailure(current, latest.error);
            markDirty('targeted');
          }
        }
        (render.loader || noop)(componentId, kind);
      }
    } finally {
      timers.clearTimeout(timeoutId);
      const latest = state.loaders.get(key);
      if (latest && latest.sequence === requestSequence) {
        latest.pending = false;
        (render.loader || noop)(componentId, kind);
      }
    }
  }

  function parsePsalm(data) {
    if (typeof data !== 'object' || data === null || typeof data.reference !== 'string' || typeof data.meter !== 'string' || !Array.isArray(data.slides) || data.slides.some(text => typeof text !== 'string')) {
      throw new Error('Psalm response was malformed. Retry.');
    }
    return { meter: data.meter, slides: data.slides };
  }

  function parseEsv(data) {
    if (typeof data !== 'object' || data === null || typeof data.ok !== 'boolean') throw new Error('ESV response was malformed. Retry.');
    if (data.ok !== true) throw new Error(typeof data.warning === 'string' && data.warning ? data.warning : 'ESV text could not be fetched. Enter the text manually.');
    if (typeof data.text !== 'string') throw new Error('ESV response was malformed. Retry.');
    return data.text;
  }

  function loadPsalm(componentId, reference) {
    return runLoader('psalm', componentId, reference, `/api/psalm?reference=${encodeURIComponent(reference)}`, parsePsalm, (component, data) => {
      component.slide_breaks = data.slides;
    });
  }

  function loadEsv(componentId, reference) {
    return runLoader(
      'esv',
      componentId,
      reference,
      `/api/scripture?reference=${encodeURIComponent(reference)}`,
      parseEsv,
      (component, text) => {
        component.text = text;
        component.external_source_failed = false;
      },
      component => {
        component.external_source_failed = true;
      },
    );
  }

  function renderScope(scope) {
    if (scope === 'structural') {
      (render.all || noop)();
    } else if (scope === 'heading') {
      (render.heading || noop)();
      (render.orderItem || noop)(state.selectedId);
      (render.validation || noop)();
    } else if (scope === 'loader') {
      (render.editor || noop)();
      (render.orderItem || noop)(state.selectedId);
      (render.counts || noop)();
      (render.validation || noop)();
    } else if (scope === 'summary') {
      (render.orderItem || noop)(state.selectedId);
      (render.counts || noop)();
      (render.validation || noop)();
    } else {
      (render.counts || noop)();
      (render.validation || noop)();
    }
  }

  function invalidateLoaderSequences() {
    for (const loader of state.loaders.values()) {
      loader.sequence += 1;
      loader.pending = false;
      loader.error = null;
    }
  }

  function markDirty(scope = 'targeted') {
    if (scope === 'structural') invalidateLoaderSequences();
    state.editGeneration += 1;
    state.status = state.activeSave ? 'Saving' : 'Unsaved';
    setSaveState(state.status, state.status);
    if (!state.conflict) setSaveHelp('');
    renderScope(scope);
    scheduleSave();
  }

  function updateComponent(id, mutate, scope = 'field') {
    const component = findComponent(id);
    if (!component) return;
    mutate(component);
    markDirty(scope === 'field' ? 'targeted' : scope);
  }

  function updateService(mutate, scope = 'targeted') {
    if (!state.service) return;
    mutate(state.service);
    markDirty(scope === 'targeted' ? 'targeted' : scope);
  }

  function scheduleSave() {
    if (state.saveTimer !== null) timers.clearTimeout(state.saveTimer);
    state.saveTimer = timers.setTimeout(() => {
      state.saveTimer = null;
      if (state.activeSave) return;
      void startSave().catch(() => {});
    }, SAVE_DEBOUNCE_MS);
  }

  function performLeaseRenewal(force = false) {
    if (!state.service) return null;
    const remaining = Date.parse(state.lease?.expires_at || 0) - Date.now();
    if (!force && state.lease && remaining > 90_000) return state.lease;
    if (state.activeLeaseRenewal) return state.activeLeaseRenewal;
    const operation = checkedRequest(`/api/services/${state.service.id}/lock`, { method: 'POST' }).then(async response => {
      const renewed = await response.json();
      state.lease = renewed;
      state.service.lease = renewed;
      return renewed;
    });
    const renewal = operation.finally(() => {
      if (state.activeLeaseRenewal === renewal) state.activeLeaseRenewal = null;
    });
    state.activeLeaseRenewal = renewal;
    return renewal;
  }

  function renewLease(force = false, { fromSave = false } = {}) {
    if (!fromSave && state.activeSave) return state.activeSave.then(() => state.lease);
    return performLeaseRenewal(force);
  }

  function loadService(record, { discardUnsaved = false } = {}) {
    return enqueueTransition(async () => {
      if (discardUnsaved) {
        clearPendingTimer();
        if (state.activeSave) {
          try { await state.activeSave; } catch {}
        }
      } else {
        await flushDirtyGenerations();
      }
      clearPendingTimer();
      const hadConflict = Boolean(state.conflict);
      state.loaders.clear();
      state.service = record;
      state.lease = record.lease || null;
      state.selectedId = record.components[0]?.id || null;
      state.editGeneration = 0;
      state.savedGeneration = 0;
      state.status = 'Saved';
      state.conflict = null;
      if (hadConflict) setConflict(null);
      setConflictRecovery?.(null);
      try {
        await renewLease(true);
      } catch (error) {
        state.lease = null;
        state.service.lease = null;
        state.status = 'Failed';
        setSaveState('Failed', 'Failed');
        setSaveHelp(error.message);
        showToast(error.message);
      }
      (render.all || noop)();
    });
  }

  async function startSave() {
    if (!state.service) return;
    if (state.conflict) throw state.conflict;
    if (state.activeSave) return state.activeSave;
    clearPendingTimer();
    state.status = 'Saving';
    setSaveState('Saving', 'Saving');
    const operation = (async () => {
      if (!state.lease) throw new Error('This service is read-only. Reload it to acquire the editing lease.');
      const lease = renewLease(false, { fromSave: true });
      if (lease?.then) await lease;
      const sentGeneration = state.editGeneration;
      const body = JSON.stringify(state.service);
      const response = await checkedRequest(`/api/services/${state.service.id}/autosave`, { method: 'PUT', body });
      const saved = await response.json();
      mergeServerMetadata(saved);
      state.savedGeneration = sentGeneration;
      if (state.editGeneration === sentGeneration) {
        state.status = 'Saved';
        setSaveState('Saved', 'Saved');
        setSaveHelp('');
      } else {
        state.status = 'Unsaved';
        setSaveState('Unsaved', 'Unsaved');
        scheduleSave();
      }
    })().catch(error => {
      clearPendingTimer();
      state.status = 'Failed';
      setSaveState('Failed', 'Failed');
      setSaveHelp(error.message);
      showToast(error.message);
      if (error.status === 409) {
        state.conflict = error;
        setConflict(error);
        setConflictRecovery?.(null);
      }
      throw error;
    }).finally(() => {
      state.activeSave = null;
      setSaveState(state.status, state.status);
    });
    state.activeSave = operation;
    return operation;
  }

  function enqueueTransition(operation) {
    state.transitionDepth += 1;
    const run = async () => {
      try {
        return await operation();
      } finally {
        state.transitionDepth -= 1;
      }
    };
    const transition = state.transition.then(run, run);
    state.transition = transition.catch(() => {});
    return transition;
  }

  async function flushDirtyGenerations({ retryActiveFailure = false } = {}) {
    clearPendingTimer();
    if (state.activeSave) {
      try {
        await state.activeSave;
      } catch (error) {
        if (!retryActiveFailure) throw error;
      }
    }
    if (!state.service || !isDirty()) return;
    while (state.editGeneration > state.savedGeneration) {
      await startSave();
    }
  }

  function clearPendingTimer() {
    if (state.saveTimer !== null) timers.clearTimeout(state.saveTimer);
    state.saveTimer = null;
  }

  function saveNow() {
    if (state.transitionDepth === 0) return flushDirtyGenerations({ retryActiveFailure: true });
    return enqueueTransition(() => flushDirtyGenerations({ retryActiveFailure: true }));
  }

  async function flushPendingSave() {
    return saveNow();
  }

  function isDirty() { return state.editGeneration > state.savedGeneration; }
  function isSaving() { return state.activeSave !== null; }

  function keepEditingAfterConflict() {
    if (!state.conflict) return;
    state.status = 'Unsaved';
    setSaveState('Unsaved', 'Unsaved');
    setSaveHelp('The service remains local and unsaved. Review the conflict controls and reload the server version before saving.');
    setConflict(null);
    setConflictRecovery?.(state.conflict);
  }

  function reopenConflictControls() {
    if (!state.conflict) return;
    setConflict(state.conflict);
    setConflictRecovery?.(null);
  }

  return {
    getService: () => state.service,
    getLease: () => state.lease,
    getState: () => state,
    findComponent,
    loadService,
    selectComponent: id => { state.selectedId = id; },
    updateComponent,
    updateService,
    markDirty,
    saveNow,
    flushPendingSave,
    isDirty,
    isSaving,
    renewLease,
    loadPsalm,
    loadEsv,
    keepEditingAfterConflict,
    reopenConflictControls,
  };
}
