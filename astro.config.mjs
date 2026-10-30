import { defineConfig } from "astro/config";

export default defineConfig({
  site: "https://mudassir-hazoor.github.io",
  base: "/sricptum_repo",
  output: "static",
  build: {
    format: "directory",
  },
});
