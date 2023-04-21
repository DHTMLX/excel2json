import { resolve } from "path";
import { defineConfig } from 'vite'

import wasm from "vite-plugin-wasm";
import topLevelAwait from "vite-plugin-top-level-await";

export default defineConfig({
  base:"./",
  plugins: [
    wasm()
  ],
  format: "esm",
  build:{
    target: "esnext",
    rollupOptions: {
      input:{
        main: resolve(__dirname, "index.html"),
        worker: resolve(__dirname, "worker.html"),
      },
      output: {
        entryFileNames: `assets/[name].js`,
        chunkFileNames: `assets/[name].js`,
        assetFileNames: `assets/[name].[ext]`,
      }
    }
  },
  worker:{
    plugins: [
      wasm(),
      topLevelAwait()
    ],
    format: "esm",
    rollupOptions: {
      output: {
        entryFileNames: `assets/[name].js`,
        chunkFileNames: `assets/[name].js`,
        assetFileNames: `assets/[name].[ext]`,
      }
    }
  },
});