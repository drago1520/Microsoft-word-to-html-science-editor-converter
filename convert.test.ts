import assert from 'node:assert/strict'
import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const cwd = await mkdtemp(join(tmpdir(), 'docx-convert-'))
try {
  const result = Bun.spawnSync(
    [process.execPath, join(import.meta.dir, 'convert.ts'), `--input=${join(import.meta.dir, 'article1.docx')}`],
    { cwd, stderr: 'pipe' }
  )
  assert.equal(result.exitCode, 0, result.stderr.toString())
  const html = await readFile(join(cwd, 'public/article1/index.html'), 'utf8')
  assert.doesNotMatch(html, /style="\s*;/)
  const cell = html.match(/<td\b([^>]*)>\s*<div[^>]*>\s*<p><strong>Tumorigenesis<\/strong>/)?.[1]
  assert.ok(cell, 'Tumorigenesis cell exists')
  assert.match(cell, /border-top: none;/)
  assert.match(cell, /border-bottom: none;/)
  assert.match(cell, /border-left: 0\.75pt solid #000000;/)
  const article2 = Bun.spawnSync(
    [process.execPath, join(import.meta.dir, 'convert.ts'), `--input=${join(import.meta.dir, 'article2.docx')}`],
    { cwd, stderr: 'pipe' }
  )
  assert.equal(article2.exitCode, 0, article2.stderr.toString())
  const html2 = await readFile(join(cwd, 'public/article2/index.html'), 'utf8')
  const equation = html2.match(/<th\s+style="([^"]*)"><math[^>]*><semantics><mrow><mo[^>]*>\[<\/mo><mi>A<\/mi>/)?.[1]
  assert.ok(equation, 'Equation cell exists')
  for (const side of ['top', 'right', 'bottom', 'left']) assert.ok(equation.includes(`border-${side}: none;`))
  assert.match(html2, /table tbody \{ border: none; \}/)
} finally {
  await rm(cwd, { recursive: true, force: true })
}
