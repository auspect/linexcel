import { build } from "vite";
import { readFile, writeFile } from "node:fs/promises";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const root = dirname(fileURLToPath(import.meta.url));
const output = resolve(root, "../src/linexcel/assets");
const result = await build({
  configFile: false,
  root,
  logLevel: "warn",
  define: { "process.env.NODE_ENV": '"production"' },
  build: {
    write: false,
    minify: true,
    target: "es2022",
    lib: {
      entry: resolve(root, "src/graph.jsx"),
      name: "LinexcelGraph",
      formats: ["iife"],
      fileName: () => "app-graph.js",
      cssFileName: "app-graph",
    },
    rollupOptions: { output: { inlineDynamicImports: true } },
  },
});
for (const file of (Array.isArray(result) ? result : [result]).flatMap(
  (item) => item.output,
)) {
  const bytes = Buffer.from(file.type === "asset" ? file.source : file.code);
  const path = resolve(output, file.fileName);
  if (process.argv.includes("--check")) {
    if (!bytes.equals(await readFile(path)))
      throw new Error(
        `${file.fileName} is stale: run npm ci && npm run build in frontend/`,
      );
  } else await writeFile(path, bytes);
}
console.log(
  process.argv.includes("--check")
    ? "Committed frontend bundle matches sources."
    : "Offline frontend bundle built.",
);
