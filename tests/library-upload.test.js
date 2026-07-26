import { beforeAll, describe, expect, it, vi } from 'vitest';

// library.js wires itself to the page and loads the catalogue as soon as it is imported, so the
// markup it expects has to exist first.
beforeAll(async () => {
  document.head.innerHTML = '<meta name="csrf-token" content="test-csrf">';
  document.body.innerHTML = `
    <input id="song-search">
    <div id="library-results"></div>
    <div id="song-preview"></div>
    <p id="song-count"></p>
    <button class="sign-out"></button>
    <button id="add-song"></button>
    <dialog id="add-song-dialog"><form id="add-song-form">
      <p id="add-song-error" hidden></p>
      <button id="add-song-submit" type="submit"></button>
      <button type="button" data-close></button>
    </form></dialog>
    <input id="replace-song-file" type="file">
    <div id="toast"></div>
  `;
  global.fetch = vi.fn().mockResolvedValue({ ok: true, json: async () => [] });
});

describe('parseAliases', () => {
  it('splits on commas and drops blanks and stray spacing', async () => {
    const { parseAliases } = await import('../crates/server/static/library.js');
    expect(parseAliases(' And can it be ,, Amazing love ')).toEqual([
      'And can it be',
      'Amazing love',
    ]);
    expect(parseAliases('')).toEqual([]);
    expect(parseAliases(undefined)).toEqual([]);
  });
});

describe('describeFileProblem', () => {
  const file = (name, size) => ({ name, size });

  it('accepts a reasonable .pptx', async () => {
    const { describeFileProblem } = await import('../crates/server/static/library.js');
    expect(describeFileProblem(file('And Can it Be.pptx', 400_000))).toBeNull();
    expect(describeFileProblem(file('SHOUTING.PPTX', 400_000))).toBeNull();
  });

  it('explains the mistakes worth catching before uploading', async () => {
    const { describeFileProblem } = await import('../crates/server/static/library.js');
    expect(describeFileProblem(null)).toMatch(/Choose a PowerPoint/);
    expect(describeFileProblem(file('song.ppt', 1000))).toMatch(/\.pptx/);
    expect(describeFileProblem(file('song.pdf', 1000))).toMatch(/\.pptx/);
    expect(describeFileProblem(file('song.pptx', 0))).toMatch(/empty/);
    expect(describeFileProblem(file('song.pptx', 26 * 1024 * 1024))).toMatch(/25 MB/);
  });
});

describe('headerSafeFilename', () => {
  it('keeps the name latin-1 so it can travel in a header', async () => {
    const { headerSafeFilename } = await import('../crates/server/static/library.js');
    expect(headerSafeFilename('And Can it Be.pptx')).toBe('And Can it Be.pptx');
    expect(headerSafeFilename('O’Come — all ye.pptx')).toBe('OCome  all ye.pptx');
    expect(headerSafeFilename('☃.pptx')).toBe('.pptx');
    expect(headerSafeFilename('')).toBe('song.pptx');
  });
});
