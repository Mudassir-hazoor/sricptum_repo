import { defineConfig } from 'astro/config';

export default defineConfig({
  site: 'https://example.github.io',
  base: '/scriptorium',
  output: 'static',
  build: {
    format: 'directory',
  },
});
