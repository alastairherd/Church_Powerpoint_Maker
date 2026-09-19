// @vitest-environment node
import { afterEach, describe, expect, it } from 'vitest';
import { mkdtempSync, readFileSync, writeFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';

const directories = [];
afterEach(() => directories.splice(0).forEach(path => rmSync(path, { recursive: true, force: true })));
function run({ failures = 1, env = {} } = {}) {
  const dir = mkdtempSync(join(tmpdir(), 'church-dns-'));
  directories.push(dir);
  const resolver = join(dir, 'resolv.conf');
  const original = 'nameserver 127.0.0.11\nsearch internal\n';
  writeFileSync(resolver, original);
  const source = readFileSync(new URL('../scripts/docker-entrypoint.sh', import.meta.url), 'utf8');
  writeFileSync(join(dir, 'entrypoint.sh'), source.replaceAll('/etc/resolv.conf', resolver));
  writeFileSync(join(dir, 'getent'), `#!/bin/sh
count=0
[ ! -f "$DNS_TEST_DIR/count" ] || count=$(cat "$DNS_TEST_DIR/count")
count=$((count + 1))
printf '%s' "$count" > "$DNS_TEST_DIR/count"
[ "$count" -gt "$DNS_TEST_FAILURES" ]
`, { mode: 0o755 });
  const result = spawnSync('/bin/sh', [join(dir, 'entrypoint.sh'), '/bin/sh', '-c', 'echo app-started'], {
    encoding: 'utf8', env: { ...process.env, PATH: `${dir}:${process.env.PATH}`, DNS_TEST_DIR: dir,
      DNS_TEST_FAILURES: String(failures), OBJECT_STORE: 'r2', R2_ACCOUNT_ID: 'test-account', ...env },
  });
  return { ...result, original, resolver: readFileSync(resolver, 'utf8') };
}

describe('container DNS recovery', () => {
  it('preserves working Docker DNS', () => {
    const result = run({ failures: 0 });
    expect(result.status).toBe(0);
    expect(result.resolver).toBe(result.original);
    expect(result.stdout).toContain('app-started');
  });
  it('recovers broken DNS while retaining internal resolver entries', () => {
    const result = run();
    expect(result.status).toBe(0);
    expect(result.resolver).toBe(`nameserver 1.1.1.1\nnameserver 9.9.9.9\n${result.original}`);
    expect(result.stderr).toContain('recovered');
  });
  it('restores original DNS when public resolvers also fail', () => {
    const result = run({ failures: 2 });
    expect(result.status).toBe(0);
    expect(result.resolver).toBe(result.original);
    expect(result.stderr).toContain('check container networking');
  });
  it('allows administrators to disable or customise fallbacks', () => {
    expect(run({ env: { APP_DNS_FALLBACKS: '' } }).resolver).toBe('nameserver 127.0.0.11\nsearch internal\n');
    expect(run({ env: { APP_DNS_FALLBACKS: '10.0.0.53' } }).resolver).toContain('nameserver 10.0.0.53\n');
  });
  it('skips DNS changes for memory storage', () => {
    const result = run({ env: { OBJECT_STORE: 'memory' } });
    expect(result.resolver).toBe(result.original);
    expect(result.status).toBe(0);
  });
  it('rejects invalid resolver input without changing DNS', () => {
    const result = run({ env: { APP_DNS_FALLBACKS: 'invalid/host' } });
    expect(result.status).toBe(1);
    expect(result.resolver).toBe(result.original);
  });
});
