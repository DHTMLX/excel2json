import { resolve } from "path";
import { defineConfig } from 'vite'
import wasm from "vite-plugin-wasm";
import topLevelAwait from "vite-plugin-top-level-await";

const baseConfig = {
  base: "./",
  plugins: [
    wasm(),
    topLevelAwait()
  ],
  build: {
    target: "esnext",
    emptyOutDir: false,
  }
};

// Use mode to determine which config to use
export default defineConfig(({ mode }) => {
  if (mode === 'module') {
    return {
      ...baseConfig,
      build: {
        ...baseConfig.build,
        lib: {
          entry: resolve(__dirname, "./js/module.js"),
          formats: ['es'],
          fileName: () => 'module.js',
        },
        rollupOptions: {
          output: {
            entryFileNames: `[name].js`,
            chunkFileNames: `[name].js`,
            assetFileNames: `[name].[ext]`
          }
        }
      }
    }
  }

  if (mode === 'worker') {
    return {
      ...baseConfig,
      build: {
        ...baseConfig.build,
        lib: {
          entry: resolve(__dirname, "./js/worker.js"),
          formats: ['cjs'],
          fileName: () => 'worker.js',
        },
        rollupOptions: {
          output: {
            entryFileNames: `[name].js`,
            chunkFileNames: `[name].js`,
            assetFileNames: `[name].[ext]`
          }
        }
      }
    }
  }

  throw new Error('Please specify mode: module or worker');
});