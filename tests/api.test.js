import { beforeEach, describe, expect, it, vi } from 'vitest';
import { createApiRequest } from '../crates/server/static/api.js';

const response = (status, data) => new Response(JSON.stringify(data), { status });
const rejected = () => response(403, { error: 'CSRF token is missing or invalid' });
beforeEach(() => { document.head.innerHTML = '<meta name="csrf-token" content="old">'; });

describe('session recovery', () => {
  it('refreshes a stale token and preserves the export request when retrying', async () => {
    const calls = [];
    const fetchImpl = vi.fn(async (url, options) => {
      calls.push({ url, ...options, headers: new Headers(options.headers) });
      return [rejected(), response(200, { csrf: 'fresh' }), response(200, {})][calls.length - 1];
    });
    await createApiRequest({ fetchImpl })('/api/services/one/generate', { method: 'POST', body: 'payload' });
    expect(calls.map(call => call.url)).toEqual(['/api/services/one/generate', '/api/session', '/api/services/one/generate']);
    expect(calls[0].headers.get('x-csrf-token')).toBe('old');
    expect(calls[1].cache).toBe('no-store');
    expect(calls[2].headers.get('x-csrf-token')).toBe('fresh');
    expect(calls[2].body).toBe('payload');
    expect(document.querySelector('meta').content).toBe('fresh');
  });

  it('recovers when the page has no token', async () => {
    document.head.innerHTML = '';
    const fetchImpl = vi.fn().mockResolvedValueOnce(rejected())
      .mockResolvedValueOnce(response(200, { csrf: 'fresh' })).mockResolvedValueOnce(response(200, {}));
    await createApiRequest({ fetchImpl })('/api/logout', { method: 'POST' });
    expect(document.querySelector('meta[name="csrf-token"]').content).toBe('fresh');
  });

  it('stops after one retry if the session changes again', async () => {
    const fetchImpl = vi.fn().mockResolvedValueOnce(rejected())
      .mockResolvedValueOnce(response(200, { csrf: 'fresh' })).mockResolvedValueOnce(rejected());
    await expect(createApiRequest({ fetchImpl })('/api/services', { method: 'POST' })).rejects.toMatchObject({ status: 403 });
    expect(fetchImpl).toHaveBeenCalledTimes(3);
  });

  it.each([401, 403, 409, 500])('never retries an unrelated %s response', async status => {
    const fetchImpl = vi.fn().mockResolvedValue(response(status, { error: 'other error' }));
    await expect(createApiRequest({ fetchImpl })('/api/services', { method: 'POST' })).rejects.toMatchObject({ status });
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('does not retry after a network failure with an unknown mutation outcome', async () => {
    const fetchImpl = vi.fn().mockRejectedValue(new TypeError('network error'));
    await expect(createApiRequest({ fetchImpl })('/api/services', { method: 'POST' })).rejects.toThrow('network error');
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('keeps expired-session recovery explicit without replaying the mutation', async () => {
    const fetchImpl = vi.fn().mockResolvedValueOnce(rejected()).mockResolvedValueOnce(response(401, {}));
    await expect(createApiRequest({ fetchImpl })('/api/services', { method: 'POST' })).rejects.toThrow('sign in again');
    expect(fetchImpl).toHaveBeenCalledTimes(2);
  });
});
