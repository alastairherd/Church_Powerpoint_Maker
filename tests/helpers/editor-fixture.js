export function makeService(overrides = {}) {
  const reading = {
    id: 'reading-1',
    type: 'reading',
    heading: 'New Testament Reading',
    reference: 'Luke 7:11–17',
    bible_page: 842,
  };
  const psalm = {
    id: 'psalm-1',
    type: 'psalm',
    heading: 'Psalm',
    reference: 'Psalm 23:1–6',
    show_verse_numbers: true,
    slide_breaks: ['The LORD is my shepherd'],
  };
  const call = {
    id: 'call-1',
    type: 'call_to_worship',
    heading: 'Call to Worship',
    reference: 'Psalm 96:2',
    text: 'Sing to the LORD.',
    external_source_failed: false,
  };
  return {
    id: 'service-1',
    name: 'Morning service',
    date: '2026-07-19',
    preset: 'am',
    status: 'draft',
    revision: 4,
    audit: { created_at: '2026-07-19T08:00:00Z', created_by: 'Test Staff', updated_at: '2026-07-19T08:00:00Z', updated_by: 'Test Staff' },
    components: [reading, psalm, call],
    ...overrides,
  };
}

export function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

export function jsonResponse(body, { status = 200 } = {}) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

export function errorResponse(message, status) {
  return jsonResponse({ error: message }, { status });
}

export function installBuilderDom(document) {
  document.body.innerHTML = `
    <main>
      <a class="wordmark" href="/">Home</a>
      <a class="nav-link" href="/">Services</a>
       <a class="nav-link" href="/library">Song library</a>
       <a class="nav-link" href="/generated">Generated PowerPoints</a>
       <a class="nav-link" href="/admin">Administration</a>
      <button id="new-service"></button>
      <button id="create-service"></button>
      <button id="review-service"></button>
      <button id="generate-service"></button>
      <button id="review-generate"></button>
      <button id="add-component"></button>
      <button id="sign-out"></button>
      <button id="save-now"></button>
      <div id="save-state" aria-live="polite"><span></span></div>
      <p id="save-help" hidden></p>
      <input id="service-name">
      <input id="service-date">
      <select id="service-preset"></select>
      <h1 id="service-heading"></h1>
      <strong id="crumb-name"></strong>
      <ol id="component-list"></ol>
      <strong id="component-count"></strong>
      <strong id="slide-count"></strong>
      <strong id="review-slides"></strong>
      <section id="editor-panel"></section>
      <section id="validation-list"></section>
      <strong id="readiness-score"></strong>
      <span id="readiness-bar"></span>
      <dialog id="new-dialog"></dialog>
      <div id="preset-choices"></div>
      <dialog id="review-dialog"></dialog>
      <h2 id="review-title"></h2>
      <div id="full-review"></div>
      <div id="toast"></div>
    </main>`;
}
