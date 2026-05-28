import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig(({ isSsrBuild }) => ({
  plugins: [react()],

  resolve: {
    alias: {
      sheetjs: "xlsx",
    },
  },

  build: {
    // Client bundle → dist/   |   SSR bundle → dist/server/
    outDir: isSsrBuild ? "dist/server" : "dist",
    // Keep SSR bundle as ESM so the prerender script can import() it
    ...(isSsrBuild && { rollupOptions: { output: { format: "es" } } }),
  },

  ssr: {
    // Bundle xlsx (CommonJS) into the SSR output so Node can load it
    noExternal: ["xlsx"],
  },
}));
