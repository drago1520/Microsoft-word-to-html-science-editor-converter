import { convert } from "pandoc-wasm";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname } from "node:path";

const docx = new Blob([await readFile("article1.docx")]);

const result = await convert(
  {
    from: "docx",
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
await writeFile("article1.html", result.stdout);

for (const [name, blob] of Object.entries<Blob>(result.mediaFiles)) {
  await mkdir(dirname(name), { recursive: true });
  await writeFile(name, Buffer.from(await blob.arrayBuffer()));
}

console.log(
  `wrote article1.html + ${Object.keys(result.mediaFiles).length} media files`,
);
//TODO: convert .png -> optimized .webp images. Pandoc can't natively. Use sharp