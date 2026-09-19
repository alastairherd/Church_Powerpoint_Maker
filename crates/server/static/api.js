// Retry only an explicit CSRF rejection: middleware rejected it before any mutation.
export function createApiRequest({ document: doc = globalThis.document, fetchImpl = globalThis.fetch } = {}) {
  async function checked(response) {
    if (response.ok) return response;
    const data = await response.json().catch(() => ({}));
    const error = new Error(response.status === 401
      ? 'Your session has expired. Keep any unsaved text and sign in again.'
      : data.error || `Request failed (${response.status})`);
    error.status = response.status;
    error.body = data;
    throw error;
  }

  return async function request(url, options = {}) {
    const headers = new Headers(options.headers || {});
    const mutation = !['GET', 'HEAD', 'OPTIONS'].includes((options.method || 'GET').toUpperCase());
    let meta = doc.querySelector('meta[name="csrf-token"]');
    if (mutation && meta?.content) headers.set('x-csrf-token', meta.content);
    const send = () => fetchImpl(url, { ...options, headers, credentials: 'same-origin', cache: 'no-store' });
    let response = await send();
    if (mutation && response.status === 403) {
      const data = await response.clone().json().catch(() => ({}));
      if (data.error === 'CSRF token is missing or invalid') {
        const sessionResponse = await checked(await fetchImpl('/api/session', {
          credentials: 'same-origin', cache: 'no-store', signal: options.signal,
        }));
        const session = await sessionResponse.json();
        if (!session.csrf) throw new Error('Could not refresh your session. Keep any unsaved text and sign in again.');
        if (!meta) {
          meta = doc.createElement('meta');
          meta.name = 'csrf-token';
          doc.head.append(meta);
        }
        meta.content = session.csrf;
        headers.set('x-csrf-token', session.csrf);
        response = await send();
      }
    }
    return checked(response);
  };
}
