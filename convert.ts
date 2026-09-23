/** @usage bun run convert.ts --input="./article1.docx" */
//@ts-expect-error no types for this lib. Dw
import { convert } from 'pandoc-wasm'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { basename, dirname, extname, join } from 'node:path'
import { parseArgs } from 'node:util'
import { wordCellBorders, wordStylesToCss } from './utils/custom-styles-css'

const { values } = parseArgs({
  args: Bun.argv.slice(2),
  options: {
    input: { type: 'string' }
  }
})

if (!values.input) throw new Error('Required: --input=<path.docx>')
const ARTICLE_PATH = values.input
const inputName = basename(ARTICLE_PATH)
const OUT = join('public', basename(ARTICLE_PATH, extname(ARTICLE_PATH)))

const docxBytes = await readFile(ARTICLE_PATH)
const docx = new Blob([docxBytes])

const result = await convert(
  {
    from: 'docx+styles',
    to: 'html5',
    standalone: true,
    'input-files': [inputName],
    'extract-media': '.',
    'section-divs': true,
    // Default HTML math is "plain", which silently dumps raw LaTeX ($$...$$) for anything it cannot fake with spans (fractions, matrices, \left|).
    'html-math-method': { method: 'mathml' }
  },
  null,
  // Key must match the name in input-files: it is the filename inside pandoc's virtual FS.
  { [inputName]: docx }
)

if (result.stderr) console.error(result.stderr)

let html = result.stdout

/** Written webp path -> its final pixel size, for width/height attributes. */
const sizes = new Map<string, { width: number; height: number }>()

// Pandoc copies embedded media out byte-for-byte, so recompress here.
for (const [name, blob] of Object.entries<Blob>(result.mediaFiles)) {
  // Path stays relative to the html, so the src attributes keep working.
  const webp = join(dirname(name), basename(name, extname(name)) + '.webp')
  await mkdir(join(OUT, dirname(webp)), { recursive: true })
  //prettier-ignore
  const img = new Bun.Image(blob)
    .resize(1200, undefined, { withoutEnlargement: true })
    .webp({ quality: 80 })
  await img.write(join(OUT, webp))
  sizes.set(webp, { width: img.width, height: img.height })
}

/** Default grid; explicit Word cell borders override it below. */
const extraCss = `
<style>
${wordStylesToCss(docxBytes)}
</style>`

/** @description Change all img src extentions --> .webp, because I transformed the images */
const cellBorders = wordCellBorders(docxBytes)
const matchingCells = cellBorders.length === [...html.matchAll(/<t[dh]\b/g)].length
let cellIndex = 0
html = new HTMLRewriter()
  .on('td, th', {
    element(el) {
      const border = cellBorders[cellIndex++]
      if (matchingCells && border)
        el.setAttribute('style', [el.getAttribute('style')?.replace(/;\s*$/, ''), border].filter(Boolean).join('; '))
    }
  })
  .on('head', {
    element(el) {
      el.append(extraCss, { html: true })
    }
  })
  .on('img', {
    element(el) {
      const src = el.getAttribute('src')
      if (!src) return
      // join() also drops pandoc's leading "./", matching the sizes keys.
      const webp = join(dirname(src), basename(src, extname(src)) + '.webp')
      el.setAttribute('src', webp)
      el.setAttribute('loading', 'lazy')
      el.setAttribute('decoding', 'async')

      // Intrinsic size reserves layout space -> no reflow when images load.
      const size = sizes.get(webp)
      if (size) {
        el.setAttribute('width', String(size.width))
        el.setAttribute('height', String(size.height))
      }
    }
  })
  .transform(html)

await mkdir(OUT, { recursive: true })
await writeFile(join(OUT, 'index.html'), html)

console.log(`wrote ${OUT}/index.html + ${Object.keys(result.mediaFiles).length} webp images`)
//I can pass a reference .docx with all title, h1, h2, h3 etc. custom styled and pandoc will pick them up.
//TODO: only medium problem is indentation on normal text. The .xml (.docx) contains the info for identation but kinda hard to parse.
