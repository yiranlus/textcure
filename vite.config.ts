import path from "node:path";
import { readFileSync, writeFileSync } from "fs";
import { defineConfig } from "vite";

import pkg from "./package.json";

export default defineConfig({
  base: "./",
  build: {
    // minify: false,
    outDir: "textcure",
    rollupOptions: {
      input: {
        index: path.resolve(__dirname, "index.html"),
        settings: path.resolve(__dirname, "settings.html"),
        about: path.resolve(__dirname, "about.html"),
        connectionError: path.resolve(__dirname, "connection-error.html"),
      },
      external: [/Asc/],
      output: {
        entryFileNames: "scripts/[name].js",
      },
    },
  },
  plugins: [
    {
      name: "generate-config",
      writeBundle() {
        const configPath = path.join(__dirname, "config.json");
        const configData = JSON.parse(readFileSync(configPath, "utf8"));
        configData.version = pkg.version;
        configData.offered = pkg.author;

        writeFileSync(
          "./textcure/config.json",
          JSON.stringify(configData, null, 2),
        );
      },
    },
  ],
});
