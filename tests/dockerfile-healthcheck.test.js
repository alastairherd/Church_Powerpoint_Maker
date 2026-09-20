// @vitest-environment node
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';

const dockerfile = readFileSync(new URL('../Dockerfile', import.meta.url), 'utf8');
const runtimeStage = dockerfile.slice(dockerfile.indexOf('FROM debian:stable-slim AS runtime'));

describe('container healthcheck support', () => {
  it('installs curl in the runtime image for Coolify HTTP probes', () => {
    expect(runtimeStage).toMatch(/apt-get install[^\n]*(?:\\\n[^\n]*)*\bcurl\b/);
  });
});
