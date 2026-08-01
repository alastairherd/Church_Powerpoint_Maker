import { describe, expect, it, vi } from 'vitest';
import { createGetRequestCache } from '../crates/server/static/app.js';
import { deferred, errorResponse, jsonResponse } from './helpers/editor-fixture.js';

describe('read-only request cache', () => {
  it('shares an in-flight GET while giving each caller a readable response body', async () => {
    const request = vi.fn(async () => jsonResponse({ slides: ['one', 'two'] }));
    const cache = createGetRequestCache(request);

    const [first, second] = await Promise.all([
      cache.get('/api/psalm?reference=Psalm%2023'),
      cache.get('/api/psalm?reference=Psalm%2023'),
    ]);

    expect(request).toHaveBeenCalledOnce();
    await expect(first.json()).resolves.toEqual({ slides: ['one', 'two'] });
    await expect(second.json()).resolves.toEqual({ slides: ['one', 'two'] });
  });

  it('does not retain failed responses and expires successful responses', async () => {
    let currentTime = 1_000;
    const request = vi.fn()
      .mockResolvedValueOnce(errorResponse('temporary failure', 503))
      .mockResolvedValue(jsonResponse({ ok: true }));
    const cache = createGetRequestCache(request, { ttlMs: 50, now: () => currentTime });

    expect((await cache.get('/api/scripture?reference=John%203%3A16')).status).toBe(503);
    expect((await cache.get('/api/scripture?reference=John%203%3A16')).status).toBe(200);
    currentTime += 51;
    expect((await cache.get('/api/scripture?reference=John%203%3A16')).status).toBe(200);
    expect(request).toHaveBeenCalledTimes(3);
  });

  it('lets one consumer time out without cancelling the shared background GET', async () => {
    const response = deferred();
    const request = vi.fn(() => response.promise);
    const cache = createGetRequestCache(request);
    const abortController = new AbortController();

    const first = cache.get('/api/teaching?source=wsc&selection=1', { signal: abortController.signal });
    abortController.abort();
    await expect(first).rejects.toMatchObject({ name: 'AbortError' });

    const second = cache.get('/api/teaching?source=wsc&selection=1');
    response.resolve(jsonResponse({ question: 'Question', answer: 'Answer' }));
    await expect((await second).json()).resolves.toEqual({ question: 'Question', answer: 'Answer' });
    expect(request).toHaveBeenCalledOnce();
  });
});
