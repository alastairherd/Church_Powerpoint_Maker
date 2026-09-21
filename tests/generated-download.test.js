import { beforeEach, describe, expect, it, vi } from 'vitest';

const RECORD = {
  service_id: 'service-1',
  service_name: 'Morning service',
  service_date: '2026-08-09',
  revision: 26,
  generated_at: '2026-07-26T17:53:00Z',
  generated_by: 'Alastair',
  expires_at: '2026-09-26T17:53:00Z',
  source_revision: 4,
  download_url: '/api/services/service-1/revisions/26/download',
  snapshot_url: '/api/services/service-1/revisions/26/snapshot',
  restore_url: '/api/services/service-1/revisions/26/restore',
};

const SNAPSHOT = {
  id: 'service-1',
  name: 'Morning service',
  date: '2026-08-09',
  preset: 'am',
  status: 'completed',
  revision: 4,
  components: [
    { type: 'welcome', id: 'c1', heading: 'Welcome' },
    { type: 'song', id: 'c2', title: 'Tell Out My Soul' },
    { type: 'reading', id: 'c3', heading: 'First Reading', reference: 'Genesis 1:1' },
  ],
  audit: { created_at: '', created_by: 'Alastair', updated_at: '', updated_by: 'Alastair' },
};

/// generated.js wires itself to the page as soon as it is imported, so the markup and the
/// catalogue response have to be in place first.
async function bootPage(handlers) {
  document.head.innerHTML = '<meta name="csrf-token" content="test-csrf">';
  document.body.innerHTML = `
    <div id="saved-orders-results" aria-busy="true"></div>
    <p id="saved-orders-count"></p>
    <div id="generated-results" aria-busy="true"></div>
    <p id="generated-count"></p>
    <button class="sign-out"></button>
    <div id="toast"></div>
  `;
  global.fetch = vi.fn(async (url, options) => {
    const handler = handlers[url] || (url === '/api/services' ? () => ({ ok: true, json: async () => [] }) : handlers.default);
    if (!handler) throw new Error(`unexpected request to ${url}`);
    return handler(options);
  });
  globalThis.URL.createObjectURL = vi.fn(() => 'blob:deck');
  globalThis.URL.revokeObjectURL = vi.fn();
  vi.resetModules();
  await import('../crates/server/static/generated.js');
  await vi.waitFor(() => {
    expect(document.getElementById('generated-results').getAttribute('aria-busy')).toBe('false');
    expect(document.getElementById('saved-orders-results').getAttribute('aria-busy')).toBe('false');
  });
}

const listingOk = () => ({ ok: true, json: async () => [RECORD] });

function button(label) {
  return [...document.querySelectorAll('.generated-row button')]
    .find(element => element.textContent === label);
}

beforeEach(() => {
  vi.restoreAllMocks();
});

describe('downloading a generated deck', () => {
  it('saves the file when the server returns one', async () => {
    await bootPage({
      '/api/generated': listingOk,
      [RECORD.download_url]: () => ({
        ok: true,
        headers: new Headers({ 'content-disposition': 'attachment; filename="Morning-service-2026-08-09-r26.pptx"' }),
        blob: async () => new Blob([new Uint8Array([0x50, 0x4b])]),
      }),
    });

    const anchors = [];
    const createElement = document.createElement.bind(document);
    vi.spyOn(document, 'createElement').mockImplementation(tag => {
      const element = createElement(tag);
      if (tag === 'a') {
        element.click = vi.fn();
        anchors.push(element);
      }
      return element;
    });

    button('Download PowerPoint').click();
    await vi.waitFor(() => expect(anchors.length).toBe(1));
    expect(anchors[0].download).toBe('Morning-service-2026-08-09-r26.pptx');
    expect(anchors[0].click).toHaveBeenCalled();
    expect(document.getElementById('toast').textContent).toBe('');
  });

  // The bug this guards: a plain download link saved whatever came back, so a failed request
  // landed in the staff member's downloads as a .pptx full of JSON that PowerPoint refuses.
  it('reports a failure instead of saving the error response as a .pptx', async () => {
    await bootPage({
      '/api/generated': listingOk,
      [RECORD.download_url]: () => ({
        ok: false,
        status: 404,
        json: async () => ({ error: 'record not found' }),
      }),
    });

    button('Download PowerPoint').click();
    await vi.waitFor(() =>
      expect(document.getElementById('toast').textContent).toBe('record not found'));
    expect(globalThis.URL.createObjectURL).not.toHaveBeenCalled();
  });
});

describe('viewing what a deck was generated from', () => {
  it('shows the order of service the PowerPoint was built from', async () => {
    await bootPage({
      '/api/generated': listingOk,
      [RECORD.snapshot_url]: () => ({ ok: true, json: async () => SNAPSHOT }),
    });

    button('View contents').click();
    await vi.waitFor(() =>
      expect(document.querySelectorAll('.snapshot-components li').length).toBe(3));
    const text = document.querySelector('.snapshot-panel').textContent;
    expect(text).toContain('Welcome');
    expect(text).toContain('Tell Out My Soul');
    expect(text).toContain('Genesis 1:1');
    expect(text).toContain('3 items');
  });

  it('offers nothing to view for decks generated before the service was recorded', async () => {
    await bootPage({
      '/api/generated': () => ({
        ok: true,
        json: async () => [{ ...RECORD, snapshot_url: null, restore_url: null }],
      }),
    });

    expect(button('View contents')).toBeUndefined();
    expect(button('Use as starting point')).toBeUndefined();
    expect(button('Download PowerPoint')).toBeDefined();
  });
});

describe('reusing a generated deck in the builder', () => {
  it('creates a new draft from the saved revision and opens that exact service', async () => {
    await bootPage({
      '/api/generated': listingOk,
      [RECORD.restore_url]: () => ({
        ok: true,
        json: async () => ({ ...SNAPSHOT, id: 'restored service/2', name: 'Copy of Morning service' }),
      }),
    });

    const anchors = [];
    const createElement = document.createElement.bind(document);
    vi.spyOn(document, 'createElement').mockImplementation(tag => {
      const element = createElement(tag);
      if (tag === 'a') {
        element.click = vi.fn();
        anchors.push(element);
      }
      return element;
    });

    button('Use as starting point').click();
    await vi.waitFor(() => expect(anchors).toHaveLength(1));
    const [, options] = global.fetch.mock.calls.find(([url]) => url === RECORD.restore_url);
    expect(options.method).toBe('POST');
    expect(options.headers.get('x-csrf-token')).toBe('test-csrf');
    expect(new URL(anchors[0].href).searchParams.get('service')).toBe('restored service/2');
    expect(anchors[0].click).toHaveBeenCalled();
  });

  it('reports a restore failure and leaves the action available to retry', async () => {
    await bootPage({
      '/api/generated': listingOk,
      [RECORD.restore_url]: () => ({
        ok: false,
        status: 404,
        json: async () => ({ error: 'saved settings are unavailable' }),
      }),
    });

    const restore = button('Use as starting point');
    restore.click();
    await vi.waitFor(() =>
      expect(document.getElementById('toast').textContent).toBe('saved settings are unavailable'));
    expect(restore.disabled).toBe(false);
    expect(restore.textContent).toBe('Use as starting point');
  });
});

describe('saved service orders', () => {
  it('lists unfinished orders with a link to resume the same service and no download', async () => {
    const draft = { ...SNAPSHOT, id: 'saved order/1', status: 'draft', name: 'Next Sunday', audit: { ...SNAPSHOT.audit, updated_at: '2026-09-21T10:00:00Z' } };
    await bootPage({
      '/api/generated': listingOk,
      '/api/services': () => ({ ok: true, json: async () => [draft, SNAPSHOT, { ...draft, id: 'archived', status: 'archived' }] }),
    });
    const rows = document.querySelectorAll('.saved-order-row');
    expect(rows).toHaveLength(1);
    expect(rows[0].textContent).toContain('Next Sunday');
    expect(rows[0].textContent).toContain('Draft');
    const resume = rows[0].querySelector('a');
    expect(resume.textContent).toBe('Continue editing');
    expect(new URL(resume.href).searchParams.get('service')).toBe(draft.id);
    expect(rows[0].textContent).not.toContain('Download PowerPoint');
    rows[0].querySelector('button').click();
    expect(rows[0].querySelector('.snapshot-panel').hidden).toBe(false);
    expect(rows[0].textContent).toContain('Tell Out My Soul');
    expect(global.fetch.mock.calls.every(([, options]) => !options.method || options.method === 'GET')).toBe(true);
    expect(document.getElementById('generated-results').textContent).toContain('Download PowerPoint');
  });

  it('shows saved orders even when no PowerPoint has ever been generated', async () => {
    await bootPage({
      '/api/generated': () => ({ ok: true, json: async () => [] }),
      '/api/services': () => ({ ok: true, json: async () => [{ ...SNAPSHOT, status: 'draft' }] }),
    });
    expect(document.getElementById('saved-orders-count').textContent).toBe('1 saved order');
    expect(document.querySelector('.saved-order-row a').textContent).toBe('Continue editing');
    expect(document.getElementById('generated-results').textContent).toContain('No PowerPoints generated yet');
  });

  it('keeps downloads available if saved orders fail to load and offers a retry', async () => {
    let attempts = 0;
    await bootPage({
      '/api/generated': listingOk,
      '/api/services': () => ++attempts === 1
        ? { ok: false, status: 503, json: async () => ({ error: 'Storage unavailable' }) }
        : { ok: true, json: async () => [] },
    });
    expect(button('Download PowerPoint')).toBeDefined();
    const results = document.getElementById('saved-orders-results');
    expect(results.textContent).toContain('could not be loaded');
    results.querySelector('button').click();
    await vi.waitFor(() => expect(results.textContent).toContain('No unfinished service orders'));
  });
});
