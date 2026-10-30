import js from '@eslint/js';
import astro from 'eslint-plugin-astro';

export default [
  js.configs.recommended,
  ...astro.configs['flat/recommended'],
  {
    files: ['**/*.js', '**/*.mjs', '**/*.ts'],
    languageOptions: {
      ecmaVersion: 'latest',
      sourceType: 'module',
      globals: {
        process: 'readonly',
      },
    },
  },
];
