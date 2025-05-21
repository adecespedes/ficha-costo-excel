import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import path from "path";
import { fileURLToPath } from "url";

// https://vite.dev/config/
export default defineConfig({
  base: "./",
  plugins: [react()],
  build: {
    outDir: "dist",
  },
  resolve: {
    alias: {
      "@": path.resolve(path.dirname(fileURLToPath(import.meta.url)), "./src"),
    },
  },
});
