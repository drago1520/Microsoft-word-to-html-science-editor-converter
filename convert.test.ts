import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

const cwd = await mkdtemp(join(tmpdir(), "docx-convert-"));
try {
  const result = Bun.spawnSync([
    process.execPath, join(import.meta.dir, "convert.ts"),
    `--input=${join(import.meta.dir, "article1.docx")}`,
  ], { cwd, stderr: "pipe" });
  assert.equal(result.exitCode, 0, result.stderr.toString());
  assert.match(await readFile(join(cwd, "public/article1/index.html"), "utf8"), /<html/);
} finally {
  await rm(cwd, { recursive: true, force: true });
}
