import { convert } from "pandoc-wasm";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { basename, dirname, extname, join } from "node:path";
import sharp from "sharp";

const OUT = "public";

const docx = new Blob([await readFile("article1.docx")]);

const result = await convert(
  {
    from: "docx+styles",
    to: "html5",
    standalone: true,
    "input-files": ["article1.docx"],
    "extract-media": ".",
    "section-divs": true,
  },
  null,
  { "article1.docx": docx },
);

if (result.stderr) console.error(result.stderr);

let html = result.stdout;

/** Written webp path -> its final pixel size, for width/height attributes. */
const sizes = new Map<string, { width: number; height: number }>();

// Pandoc copies embedded media out byte-for-byte, so recompress here.
for (const [name, blob] of Object.entries<Blob>(result.mediaFiles)) {
  // Path stays relative to the html, so the src attributes keep working.
  const webp = join(dirname(name), basename(name, extname(name)) + ".webp");
  await mkdir(join(OUT, dirname(webp)), { recursive: true });
  const info = await sharp(Buffer.from(await blob.arrayBuffer()))
    .resize({ width: 1200, withoutEnlargement: true })
    .webp({ quality: 80 })
    .toFile(join(OUT, webp));
  sizes.set(webp, { width: info.width, height: info.height });
}

/** Pandoc's table model has no borders, so restore Word's Table Grid look. */
const extraCss = `
<style>
  table th,
  table td {
    border: 1px solid #abababff;
    padding: 0.4em 0.6em;
  }
</style>`;

/** @description Change all img src extentions --> .webp, because I transformed the images */
html = new HTMLRewriter()
  .on("head", {
    element(el) {
      el.append(extraCss, { html: true });
    },
  })
  .on("img", {
    element(el) {
      const src = el.getAttribute("src");
      if (!src) return;
      // join() also drops pandoc's leading "./", matching the sizes keys.
      const webp = join(dirname(src), basename(src, extname(src)) + ".webp");
      el.setAttribute("src", webp);
      el.setAttribute("loading", "lazy");
      el.setAttribute("decoding", "async");

      // Intrinsic size reserves layout space -> no reflow when images load.
      const size = sizes.get(webp);
      if (size) {
        el.setAttribute("width", String(size.width));
        el.setAttribute("height", String(size.height));
      }
    },
  })
  .transform(html);

await writeFile(join(OUT, "article1.html"), html);

console.log(
  `wrote ${OUT}/article1.html + ${Object.keys(result.mediaFiles).length} webp images`,
);
//I can pass a reference .docx with all title, h1, h2, h3 etc. custom styled and pandoc will pick them up.
//TODO: only medium problem is indentation on normal text. The .xml (.docx) contains the info for identation but kinda hard to parse. 