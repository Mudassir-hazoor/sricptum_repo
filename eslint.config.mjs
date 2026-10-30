import js from "@eslint/js";
import astro from "eslint-plugin-astro";

export default [
  {
    ignores: [".astro/**", "dist/**", "node_modules/**", "coverage/**"],
  },
  js.configs.recommended,
  ...astro.configs["flat/recommended"],
  {
    files: ["**/*.js", "**/*.mjs", "**/*.ts"],
    languageOptions: {
      ecmaVersion: "latest",
      sourceType: "module",
      globals: {
        process: "readonly",
      },
    },
  },
];
